<#
.SYNOPSIS
	Move and link the CCMCache and SoftwareDistribution folders out of C drive
.DESCRIPTION
	This script is used to clear the CCMCache, move the CCMCache and SoftwareDistribution folders to another drive (D drive by default), and create the juncion links in the original locations.
	This is useful if there is not enough space in C drive to store the cache and perform Windows Updates
.PARAMETER DestDrive
	Optional.
	Specify the drive letter that the cache folder should be moved to.
	Default is D if this parameter is not specified.
.PARAMETER MinDays
	Optional.
	Specify the maximum age, in number of days, of the cache files that should be kept.
	Default is 0, which means all cache will be removed.
.PARAMETER ClearCacheOnly
	Switch.
	Specify this if you only want to clear the cache without moving and linking the cache folder to another location.
.INPUTS
	None
.OUTPUTS
	None
.EXAMPLE
	.\MoveCCM.ps1
	This command clears the CCMCache, moves the 2 folders to D drive, and creates the junction links in the original locations
.EXAMPLE
	.\MoveCCM.ps1 -DestDrive E
	This command clears the CCMCache, moves the 2 folders to E drive, and creates the junction links in the original locations
.EXAMPLE
	.\MoveCCM.ps1 -ClearCacheOnly
	This commands only clear the CCMCache without moving the folders
.NOTES
	Don't forget to run this script again after the Client Health Check script re-created the SoftwareDistribution folder.
	It doesn't hurt to run this script even if the folders have already been moved.
#>

[CmdletBinding(DefaultParameterSetName = 'SetA')]
Param (
	[Parameter(Mandatory=$false, ParameterSetName="SetA")]
	[ValidatePattern("^[A-Z]|[a-z]$")]
	[string]$DestDrive = 'D',
	[Parameter(Mandatory=$false, ParameterSetName="SetB")]
	[Switch]$ClearCacheOnly,
	[Parameter(Mandatory=$false)]
	[int]$MinDays = 0
)

################ Adapted from Client Health Script - Start ################

$ServiceRunning = $true
$CCMFolders = @("SoftwareDistribution", "ccmcache")

Function Remove-CcmCache {
    Param(
        [Parameter(Mandatory=$false)]
        [int]$MinDays
    )
    $Count = 0
    $UIResourceMgr = New-Object -ComObject UIResource.UIResourceMgr
    $Cache = $UIResourceMgr.GetCacheInfo()
    $Cache.GetCacheElements() |
    where-object {[datetime]$_.LastReferenceTime -lt (get-date).Adddays(-$Mindays) -and (($_.PersistinCache -ne 1) -or ($_.Peercaching -ne 1))} |
    ForEach-Object {
        If ((Get-ItemProperty -Path $_.Location).LastWriteTime -le (get-date).AddDays(-$MinDays)){
            Write-Host "Attempting to remove $($_.Location)" -ForegroundColor Gray
            $Cache.DeleteCacheElement($_.CacheElementID)
            $Count ++
        }
    }
	if ($Count -gt 0) {
		Write-Host ''
	}
	return $Count
}

Function Remove-CCMOrphanedCache {
	try {
		 $CacheElements =  get-wmiobject -query "SELECT * FROM CacheInfoEx" -namespace "ROOT\ccm\SoftMgmtAgent"
		$ElementGroup = $CacheElements | Group-Object ContentID
		[int]$Cleaned = 0;

		#Cleanup CacheItems where ContentFolder does not exist
		$CacheElements | Where-Object {!(Test-Path $_.Location)} | ForEach-Object { $_.Delete(); $Cleaned++ }

		foreach ($ElementID in $ElementGroup) 
		{
			if ($ElementID.Count -gt 1) 
			{
				$max = ($ElementID.Group.ContentVer| Measure-Object -Maximum).Maximum

				$ElementsToRemove = $CacheElements | Where-Object {$_.contentid -eq $ElementID.Name -and $_.ContentVer-ne $Max}
				foreach ($Element in $ElementsToRemove) 
				{
					Write-Host “Deleting”$Element.ContentID”with version”$Element.ContentVersion -Color Gray

					Remove-Item $Element.Location -recurse
					$Element.Delete()
					$Cleaned++
				}
			} 
		}

		#Cleanup Orphaned Folders in ccmcache
		$UsedFolders = $CacheElements | ForEach-Object { Select-Object -inputobject $_.Location }
		[string]$CCMCache = ([wmi]"ROOT\ccm\SoftMgmtAgent:CacheConfig.ConfigKey='Cache'").Location
		if($CCMCache.EndsWith('ccmcache'))
		{
			Get-ChildItem($CCMCache) |  Where-Object{ $_.PSIsContainer } | Where-Object { $UsedFolders -notcontains $_.FullName } | ForEach-Object { Remove-Item $_.FullName -recurse ; $Cleaned++ }
		}
	}
	catch { 
		Write-Host "Failed Clearing ConfigMgr orphaned Cache items." -ForegroundColor Red
	}
	Write-Output $Cleaned
}

################ Adapted from Client Health Script - End ################

Function MoveAndLinkFolder {
	Param(
		[Parameter(Mandatory=$true)]
		[string]$FolderName
	)

	StopServicesIfRunning
	if($FolderName -eq "ccmcache") {
		takeown /a /r /d Y /f "C:\Windows\ccmcache" > $null
	}

	if (Test-Path ${DestDrive}:\CCM\$Folder -PathType Container) {
		Write-Host "${DestDrive}:\CCM\$Folder already exist. Deleting..." -ForegroundColor Yellow
		Remove-Item ${DestDrive}:\CCM\$Folder -Recurse -Force
	}

	Write-Host "Moving $FolderName..." -ForegroundColor Green
	Move-Item -Path C:\Windows\$FolderName -Destination ${DestDrive}:\CCM\$FolderName
	Write-Host "Creating junction point for $FolderName..." -ForegroundColor Green
	New-Item -ItemType Junction -Path C:\Windows\$FolderName -Target ${DestDrive}:\CCM\$FolderName > $null
}

Function StopServicesIfRunning {
	if($ServiceRunning) {
		Write-Host "Stopping the following services: CcmExec, wuauserv, BITS, cryptsvc..." -ForegroundColor Green
		Stop-Service -name CcmExec, wuauserv, BITS, cryptsvc -Force
		Write-Host "Services stopped. Sleeping for 5 seconds..." -ForegroundColor DarkGray
		Start-Sleep -Seconds 5
		Write-Host "Woke up" -ForegroundColor DarkGray
		Write-Host ""
		$ServiceRunning = $false
	}
}

############################## End of function declarations ##############################



$Count = Remove-CcmCache -MinDays $MinDays
Write-Host "Removed $Count CCMCache items" -ForegroundColor Green

$Count = Remove-CCMOrphanedCache
Write-Host "Removed $Count orphaned CCMCache" -ForegroundColor Green
Write-Host ''
	
if ($ClearCacheOnly) {
	exit
}

# If SoftwareDistributionOld is an existing junction point, it probably means that the SoftwareDistribution folder has been re-created in C drive by the Client Health script
# Need to remove the current folder in the destination drive, and move it again
if ((Test-Path C:\Windows\SoftwareDistributionOld -PathType Container) -and (Get-Item C:\Windows\SoftwareDistributionOld).Attributes.ToString().Contains("ReparsePoint")) {
	Write-Host 'C:\Windows\SoftwareDistributionOld folder is detected. Cleaning up...' -ForegroundColor Yellow
	Write-Host ''
	StopServicesIfRunning
	Remove-Item C:\Windows\SoftwareDistributionOld -Force -Recurse
	Remove-Item ${DestDrive}:\CCM\SoftwareDistribution -Force -Recurse
	Write-Host ''
}

foreach ($Folder in $CCMFolders) {
	if (Test-Path C:\Windows\$Folder -PathType Container) {
		if(-Not (Get-Item C:\Windows\$Folder).Attributes.ToString().Contains("ReparsePoint")) {
			Write-Host "C:\Windows\$Folder folder is found. Moving and linking..." -ForegroundColor Green
			Write-Host ''
			MoveAndLinkFolder $Folder
		}
		else {
			Write-Host "$Folder folder is already a juntion point" -ForegroundColor Green
		}
	}
	else {
		if (Test-Path ${DestDrive}:\CCM\$Folder -PathType Container) {
			Write-Host "$Folder folder has already been moved to $DestDrive drive" -ForegroundColor Green
		}
		else {
			Write-Host "$Folder folder is not found in C drive and $DestDrive drive" -ForegroundColor Red
		}
	}
	Write-Host ''
}

if (-Not $ServiceRunning) {
	Write-Host "Starting the following services: CcmExec, wuauserv, BITS, cryptsvc..." -ForegroundColor Green
	Start-Service -name CcmExec, wuauserv, BITS, cryptsvc
}
