# Requires -version 2.0
#Requires -version 7.0

$DestinationRootPath = "C:\Users\DennisKuhn\OneDrive - Dennis'es Services\Videos\Waka Ama"

$GoProPath = "GoPro MTP Client Disk Volume\DCIM\100GOPRO"

$ConnectedRetries = 5
$ConnectedRetrySleep = 0.5

$FoldersRetries = 10
$FoldersRetrySleep = 0.5


$RenameRetries = 5
$RenameRetrySleep = 0.5

# GoPro Vendor ID = 2672
$script:VendorID = 2672

# GoPro Service for MTP Device (Storage)
$script:Service = "WUDFWpdMtp"

$InformationPreference = "Continue"
$ErrorActionPreference = "Stop"

$script:Log = 'Application'
$script:LogSource = $MyInvocation.MyCommand.Name
$script:UseSystemLog = $false
$script:InitialisedUseSystemLog = $false

function Register-EventLog {
    if ($false -eq $script:InitialisedUseSystemLog) {
        $script:InitialisedUseSystemLog = $true
        $script:UseSystemLog = $false

        # When debugging stops with error if not using -ErrorAction Continue
        Import-WinModule Microsoft.PowerShell.Management -ErrorAction SilentlyContinue

        # Check if Log exists
        # Ref: http://msdn.microsoft.com/en-us/library/system.diagnostics.eventlog.exists(v=vs.110).aspx
        if ( [System.Diagnostics.EventLog]::Exists($script:Log) -eq $false ) {
            Write-MessageWarning "${Log} event log does not exist"
        }
        # Ref: http://msdn.microsoft.com/en-us/library/system.diagnostics.eventlog.sourceexists(v=vs.110).aspx
        # Check if Source exists
        else {
            $SourceExists = $null;
            try {
                $SourceExists = [System.Diagnostics.EventLog]::SourceExists($script:LogSource)
            } catch {
                $SourceExists = $null
            }
            if ( $SourceExists -eq $false ) {
                New-EventLog –LogName $script:Log –Source $script:LogSource -ErrorAction "Continue"
                Write-MessageInformation "Created log ${Log}/${LogSource}"
                $script:UseSystemLog = $true
            } elseif ($null -eq $SourceExists) {
                Write-MessageWarning "Run this script as administrator once if you want to use the system event log"
            } else {
                Write-MessageInformation "Use system event log ${Log}/${LogSource}"
                $script:UseSystemLog = $true
            }
        }
    }
}

function Write-MessageError {
    param($message)

    Register-EventLog
    if ( $script:UseSystemLog -eq $true) {
        Write-EventLog -LogName $script:Log -Source $script:LogSource -EntryType Error -EventID 1 -Message $message
    }
    else {
        Write-Error "$(get-date -format s): $message"
    }
}

function Write-MessageWarning {
    param($message)

    Register-EventLog
    if ( $script:UseSystemLog -eq $true) {
        Write-EventLog -LogName $script:Log -Source $script:LogSource -EntryType Warning -EventID 1 -Message $message
    }
    else {
        Write-Warning "$(get-date -format s): $message"
    }
}

function Write-MessageInformation {
    param($message)

    Register-EventLog
    if ( $script:UseSystemLog -eq $true) {
        Write-EventLog -LogName $script:Log -Source $script:LogSource -EntryType Information -EventID 1 -Message $message
    }
    else {
        Write-Information "$(get-date -format s): $message"
    }
}

function Get-ShellProxy {
    if ( -not $global:ShellProxy) {
        $global:ShellProxy = new-object -com Shell.Application
    }
    $global:ShellProxy
}
 
function Get-GoPro {
    param($deviceName)
    $shell = Get-ShellProxy
    # 17 (0x11) = ssfDRIVES from the ShellSpecialFolderConstants (https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096(v=vs.85).aspx)
    # => "My Computer" — the virtual folder that contains everything on the local computer: storage devices, printers, and Control Panel.
    # This folder can also contain mapped network drives.
    $shellItem = $shell.NameSpace(17).self

    $tries = 0;
    do {
        $tries++

        $device = $shellItem.GetFolder.items() | Where-Object { $_.name -eq $deviceName }

        If ( -not( $device )) {
            If ($tries -gt $ConnectedRetries) {
                Write-MessageError  "GoPro Device Folder '${deviceName}' not found"
            }
            else {
                Write-MessageWarning  "GoPro Device Folder '${deviceName}' not available, try #${tries}"
                Start-Sleep -Seconds $ConnectedRetrySleep
            }            
        }
    } until (
        $device -or ( $tries -gt $ConnectedRetries )
    )
    return $device
}
 
function Get-SubFolder {
    param($parent, [string]$path)
    $pathParts = @( $path.Split([system.io.path]::DirectorySeparatorChar) )
    # Get firstPathPart as it may need retries
    $firstPathPart = $pathParts | Select-Object -First 1
    # Get lastPathPart to NOT throw error, as GoPro is empty
    $lastPathPart = $pathParts | Select-Object -Last 1
    $current = $parent
    foreach ($pathPart in $pathParts) {
        if ($pathPart) {
            $tries = 0
            do {
                if ($tries -gt 0) {
                    Write-MessageWarning "GoPro Device sub folder '${pathPart}' not available, try #${tries}"
                    Start-Sleep $FoldersRetrySleep
                }
                $tries++          
                $new = $current.GetFolder.items() | Where-Object { $_.Name -eq $pathPart }
            } while (
                ($tries -le $FoldersRetries) -and
                (-not $new) -and
                ($pathPart -eq $firstPathPart)
            )
            $current = $new
            if (-not $new) {
                if ($pathPart -eq $lastPathPart) {
                    Write-MessageWarning "GoPro Device sub folder '${pathPart}' does not exist -  probably empty and reconnected"
                }
                else {
                    Write-MessageError "GoPro Device sub folder '${pathPart}' does not exist"
                }
            }
        }
    }
    return $current
}

function Get-GoProFiles {
    param($deviceName)

    Write-MessageInformation "Get-GoProFiles from '${deviceName}'"
    $GoPro = Get-GoPro -deviceName $deviceName
    $GoProFolder = Get-SubFolder -parent $GoPro -path $GoProPath

    If ($GoProFolder) {
        $GoProFolderName = $GoProFolder.Name
        $items = @( $GoProFolder.GetFolder.items() | Where-Object { $PSItem.Name -match $filter } )
        if ( -not $items) {
            Write-MessageWarning "No items in items folder"
        }
        else {
            $totalItems = $items.count
            if ($totalItems -lt 1) {
                Write-MessageWarning "Found ${totalItems} items"
            }
            else {
                Write-MessageInformation "Found ${totalItems} files"

                $TimestampText = get-date -format "yyyy-MM-dd"

                $DestinationDayPath = Join-Path -Path $DestinationRootPath -ChildPath $TimestampText
                $DestinationGoProPath = Join-Path -Path $DestinationRootPath -ChildPath $GoProFolderName
            
                if ( Test-Path $DestinationDayPath ) {
                    Write-MessageError "Destination path already exists: '${DestinationDayPath}'"
                }
                elseif ( Test-Path $DestinationGoProPath) {
                    Write-MessageError "Destination temporary path already exists: '${DestinationDayPath}'"
                }
                else {
                    Write-MessageInformation "Move ${totalItems} items: '${deviceName}\${GoProPath}' -> '${DestinationDayPath}'"

                    $shell = Get-ShellProxy
                    $destinationFolder = $shell.Namespace($DestinationRootPath).self
                    $destinationFolder.GetFolder.MoveHere($GoProFolder)

                    if ( Test-Path $DestinationGoProPath ) {
                        Rename-Item -Path $DestinationGoProPath -NewName $TimestampText

                        $tries = 0

                        while ( ($tries -lt $RenameRetries) -and -not (Test-Path $DestinationDayPath) ) {
                            $tries++
                            Write-MessageInformation "Waiting for renaming of temporary folder, try #${tries}"
                            Start-Sleep -Seconds $RenameRetrySleep
                        }
                        if ( Test-Path $DestinationDayPath) {
                            Write-MessageInformation "Moved content to: '${DestinationDayPath}'"
                        }
                        else {
                            Write-MessageError "Destination folder doesn't exist after renaming: '${DestinationGoProPath}' -> '${DestinationDayPath}'"
                        }                
                    }
                    else {
                        Write-MessageError "Destination temporary path doesn't exist after moving: '${DestinationGoProPath}'"
                    }
                }
            }
        }
    }
}

function Main {
    Write-MessageInformation "Beginning script..."

    Register-CimIndicationEvent -ClassName Win32_DeviceChangeEvent

    $GoProInfo = Get-PnpDevice -PresentOnly | Select-Object * | Where-Object { $PSItem.InstanceId -match "^USB\\VID_$script:VendorID" -and $PSItem.Service -eq $script:Service }
    if ( $GoProInfo ) {
        $GoProPresent = $true
        $GoProName = $GoProInfo.FriendlyName
        Write-MessageInformation "GoPro '${GoProName}' present"

        Get-GoProFiles -deviceName $GoProName
    }
    else {
        Write-MessageInformation "GoPro absent"
        $GoProPresent = $false
    }

    $run = $true
    do {
        # $newEvent = Wait-Event -SourceIdentifier volumeChange
        $newEvent = Wait-Event
        $eventType = $newEvent.SourceEventArgs.NewEvent.EventType
        $eventTypeName = switch ($eventType) {
            1 { "Configuration changed" }
            2 { "Device arrival" }
            3 { "Device removal" }
            4 { "docking" }
        }

        if ($GoProPresent) {
            if ($eventType -eq 3) {
                $GoProInfo = Get-PnpDevice -PresentOnly | Select-Object * | Where-Object { $PSItem.InstanceId -match "^USB\\VID_$script:VendorID" -and $PSItem.Service -eq $script:Service }

                if ( $GoProInfo ) {
                    $GoProName = $GoProInfo.FriendlyName
                    Write-MessageInformation "${eventTypeName}: GoPro ${GoProName} still present"
                }
                else {
                    $GoProPresent = $false
                    Write-MessageInformation "${eventTypeName}: GoPro ${GoProName} removed"
                }
            }
        }
        else {
            if ($eventType -eq 2) {
                $GoProInfo = Get-PnpDevice -PresentOnly | Select-Object * | Where-Object { $PSItem.InstanceId -match "^USB\\VID_$script:VendorID" -and $PSItem.Service -eq $script:Service }

                if ( $GoProInfo ) {
                    $GoProPresent = $true
                    $GoProName = $GoProInfo.FriendlyName
                    Write-MessageInformation "${eventTypeName}: GoPro ${GoProName} connected"

                    Get-GoProFiles -deviceName $GoProName
                }
                else {
                    Write-MessageWarning "${eventTypeName}: No GoPro storage found"
                }
            }
        }
        Remove-Event -SourceIdentifier $newEvent.SourceIdentifier
    } while ($run -eq $true) #Loop until next event
    # Unregister-Event -SourceIdentifier volumeChange
}

Main
