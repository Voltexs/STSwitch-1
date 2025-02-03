# ASCII Art Header
Write-Host @"
=========================================================================================================

        ███████╗███╗   ███╗ █████╗ ██████╗ ████████╗    ████████╗██████╗  █████╗ ██████ ███████╗
        ██╔════╝████╗ ████║██╔══██╗██╔══██╗╚══██╔══╝    ╚══██╔══╝██╔══██╗██╔══██╗██╔══██╗██╔════╝
        ███████╗██╔████╔██║███████║██████╔╝   ██║          ██║   ██████╔╝███████║██║  ██║█████╗  
        ╚════██║██║╚██╔╝██║██╔══██║██╔══██╗   ██║          ██║   ██╔══██╗██╔══██║██║  ██║██╔══╝  
        ███████║██║ ╚═╝ ██║██║  ██║█ ║  ██║   ██║          ██║   ██║  ██║██║  ██║██████╔╝███████╗
        ╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝          ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝ ╚══════╝     

=========================================================================================================                                                                             

"@ -ForegroundColor Yellow

# Add at the beginning after the initial variable declarations
$Global:ChecklistItems = [ordered]@{
    "Installation Log" = $false
    "Performance Settings" = $false
    "Power Settings" = $false
    "Region and Date Settings" = $false
    "C.I. Systems & Client Users" = $false
    "Installation Files Copied" = $false
    "Third Party Software" = $false
    "Fonts and BAT" = $false
    "SQL Installation" = $false
    "User Scripts & Fixes" = $false
    "DLLs Registered" = $false
    "COM+ Applications" = $false
    "COM+ Firewall" = $false
    "Desktop Shortcut" = $false
    "Radmin VPN" = $false
}

# Add function to mark tasks as complete
function Update-Checklist {
    param (
        [string]$TaskName,
        [string]$Details = ""
    )
    $Global:ChecklistItems[$TaskName] = $true
    Write-Log "Completed task: $TaskName $Details"
}

# Add function to generate PDF report
function Export-ChecklistReport {
    param (
        [string]$TechnicianName,
        [string]$ClientName
    )
    
    # Kill any existing Word processes before starting
    Get-Process "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
    
    $word = $null
    try {
        Write-Host "Initializing Word..." -ForegroundColor Yellow
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        
        Write-Host "Creating document..." -ForegroundColor Yellow
        $doc = $word.Documents.Add()

        Write-Host "Adding content..." -ForegroundColor Yellow
        # Add header
        $header = $doc.Content
        $header.Font.Size = 24
        $header.Text = "SmartTrade Installation Checklist`n`n"
        
        # Add details
        $details = $doc.Content.Paragraphs.Add()
        $details.Range.Font.Size = 10  # Smaller font size for checklist
        $details.Range.Text = @"
Installation Details:
Technician: $TechnicianName
Client: $ClientName
Date: $(Get-Date -Format "yyyy-MM-dd HH:mm")

Completed Tasks:
"@

        # Add checklist items
        foreach ($item in $Global:ChecklistItems.GetEnumerator()) {
            $status = if ($item.Value) { "✓" } else { "✗" }
            $details = $doc.Content.Paragraphs.Add()
            $details.Range.Text = "$status $($item.Key)"
        }

        # Add notes from log file
        $details = $doc.Content.Paragraphs.Add()
        $details.Range.Text = "`nDetailed Log:`n"
        if (Test-Path $logFilePath) {
            $logContent = Get-Content $logFilePath
            $details = $doc.Content.Paragraphs.Add()
            $details.Range.Text = $logContent -join "`n"
        }

        Write-Host "Saving PDF..." -ForegroundColor Yellow
        # Save directly to desktop with a simple path
        $desktopPath = [Environment]::GetFolderPath("Desktop")
        $pdfPath = "$desktopPath\SmartTrade_Checklist.pdf"
        
        # If file exists, add timestamp to name
        if (Test-Path $pdfPath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmm"
            $pdfPath = "$desktopPath\SmartTrade_Checklist_$timestamp.pdf"
        }
        
        $doc.SaveAs2($pdfPath, 17)
        Write-Host "Checklist PDF generated: $pdfPath" -ForegroundColor Green
        Write-Log "Generated checklist PDF: $pdfPath"
        
        $doc.Close()
        $word.Quit()
    }
    catch {
        # Handle errors silently
        if ($word) {
            try {
                $word.Quit()
            } catch {
                # Silent catch
            }
        }
        # Force close any remaining Word processes
        Get-Process "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
    }
    finally {
        if ($word) {
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Double-check for any remaining Word processes
        Get-Process "WINWORD" -ErrorAction SilentlyContinue | Stop-Process -Force
    }
}

# Function to handle command-line input
function ShowMenu {
    Write-Host "Select an option by entering the corresponding number:"
    Write-Host ""  # Line break

    # Define the width for the left column
    $leftColumnWidth = 50

    # Display options in two columns with proper alignment
    Write-Host ("1. Start Installation Log".PadRight($leftColumnWidth) + "8. Extract Fonts and run bat")
    Write-Host ("2. Performance Settings".PadRight($leftColumnWidth) + "9. SQL Installed (Server Only)")
    Write-Host ("3. Power Settings".PadRight($leftColumnWidth) + "10. User Scripts & Fixes (server Only)")
    Write-Host ("4. Region and Date Settings".PadRight($leftColumnWidth) + "11. DLL'S")
    Write-Host ("5. C.I. Systems & Client Users (Server Only)".PadRight($leftColumnWidth) + "12. COM+ (Server Only)")
    Write-Host ("6. Installation Files Copied".PadRight($leftColumnWidth) + "13. COM+ Firewall Bypassed (Server Only)")
    Write-Host ("7. Third Party Software Installed".PadRight($leftColumnWidth) + "14. Shortcut to Desktop")
    Write-Host ""  # Line break

    # Centered exit option
    Write-Host ("0. Exit".PadRight($leftColumnWidth) + "D. Open Device Manager (To set USB and Network)")
    Write-Host ("R. Install Radmin VPN".PadRight($leftColumnWidth) + "F. Open Format Settings (Check for Decimal)")
    Write-Host ""  # Line break
}

# Initialize log file path
$logFilePath = "C:\SmtDB\Install\log.txt"

# Function to log messages
function Write-Log {
    param (
        [string]$message
    )
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] $message"
        Add-Content -Path $logFilePath -Value $logMessage -ErrorAction Stop
    } catch {
        Write-Host "Failed to write to log file: $_" -ForegroundColor Red
    }
}

# Add global error handling
$ErrorActionPreference = "Stop"
$Global:LastError = $null

# Add error logging function
function Write-ErrorLog {
    param (
        [string]$functionName,
        [string]$errorMessage
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $errorLog = "[$timestamp] ERROR in $functionName : $errorMessage"
    Write-Log $errorLog
    Write-Host $errorLog -ForegroundColor Red
    $Global:LastError = $errorMessage
}

# Add function to verify paths exist
function Test-RequiredPath {
    param (
        [string]$path,
        [string]$description
    )
    if (-not (Test-Path $path)) {
        Write-ErrorLog "Path Verification" "Required $description not found: $path"
        return $false
    }
    return $true
}

# Function to create SmtDB folder structure
function New-SmtDBFolder {
    Write-Host "Creating SmtDB Folder structure..."
    $smtDBPath = "C:\SmtDB"
    $installPath = Join-Path $smtDBPath "Install"
    $reportsPath = Join-Path $smtDBPath "Reports"

    try {
        New-Item -Path $installPath -ItemType Directory -Force | Out-Null
        New-Item -Path $reportsPath -ItemType Directory -Force | Out-Null

        $acl = Get-Acl $smtDBPath
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("ANONYMOUS LOGON", "FullControl", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $smtDBPath -AclObject $acl

        Write-Host "SmtDB folder with Install and Reports subfolders created, and permissions granted."
        Write-Log "SmtDB folder with Install and Reports subfolders created, and permissions granted."
    } catch {
        Write-Host "An error occurred while creating the SmtDB folder structure: $_"
        Write-Log "An error occurred while creating the SmtDB folder structure: $_"
    }
}

# Main loop for command-line interface
do {
    Clear-Host  # Clear the terminal to show the ASCII art header again
    Write-Host @"
=========================================================================================================

        ███████╗███╗   ███╗ █████╗ ██████╗ ████████╗    ████████╗██████╗  █████╗ ██████ ███████╗
        ██╔════╝████╗ ████║██╔══██╗██╔══██╗╚══██╔══╝    ╚══██╔══╝██╔══██╗██╔══██╗██╔══██╗██╔════╝
        ███████╗██╔████╔██║███████║██████╔╝   ██║          ██║   ██████╔╝███████║██║  ██║█████╗  
        ╚════██║██║╚██╔╝██║██╔══██║██╔══██╗   ██║          ██║   ██╔══██╗██╔══██║██║  ██║██╔══╝  
        ███████║██║ ╚═╝ ██║██║  ██║█ ║  ██║   ██║          ██║   ██║  ██║██║  ██║██████╔╝███████╗
        ╚══════╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝  ╚═╝   ╚═╝          ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚═════╝ ╚══════╝     

=========================================================================================================                                                                             

"@ -ForegroundColor Yellow

    ShowMenu  # Show the menu after clearing the screen
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" { 
            try {
                Clear-Host
                New-SmtDBFolder
                Write-Host "Starting Installation Log..."
                
                # Validate inputs
                do {
                    $technicianName = Read-Host "Enter Technician Name"
                } while ([string]::IsNullOrWhiteSpace($technicianName))
                
                do {
                    $clientName = Read-Host "Enter Client Name"
                } while ([string]::IsNullOrWhiteSpace($clientName))
                
                do {
                    $installationType = Read-Host "Enter Installation Type (e.g., Clean Installation)"
                } while ([string]::IsNullOrWhiteSpace($installationType))

                Set-Content -Path $logFilePath -Value "Technician: $technicianName`nClient: $clientName`nInstallation Type: $installationType`n`nActions Taken:`n" -ErrorAction Stop
                Write-Log "Log file created at $logFilePath"
                Update-Checklist "Installation Log" "Technician: $technicianName"
                $Global:TechnicianName = $technicianName
                $Global:ClientName = $clientName
            } catch {
                Write-ErrorLog "Installation Log" $_
            }
        }
        "2" { 
            Clear-Host
            Write-Host "Running Performance Settings..."
            Write-Log "Running Performance Settings..."

            # Set Visual Effects to "Adjust for best performance"
            $regPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects"
            $regValue = "VisualFXSetting"

            # Set the value to 2, which corresponds to "Adjust for best performance"
            New-ItemProperty -Path $regPath -Name $regValue -Value 2 -PropertyType DWord -Force
            Write-Log "Set registry value '$regValue' at '$regPath' to 2 (Adjust for best performance)."

            # Update other visual effects for best performance
            $visualEffects = @{
                "HKCU:\Control Panel\Desktop\WindowMetrics\MinAnimate" = "0"
                "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarAnimations" = "0"
                "HKCU:\Control Panel\Desktop\FontSmoothing" = "0"
            }

            foreach ($path in $visualEffects.Keys) {
                if (Test-Path $path) {
                    $valueName = (Get-ItemProperty -Path $path).PSChildName
                    Set-ItemProperty -Path $path -Name $valueName -Value $visualEffects[$path] -Force
                    Write-Log "Set registry value '$valueName' at '$path' to '${visualEffects[$path]}'."
                }
                else {
                    Write-Output "Registry path $path does not exist. Skipping..."
                    Write-Log "Registry path $path does not exist. Skipping..."
                }
            }

            # Apply the settings immediately without restarting
            RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters
            Write-Output "Visual performance is set to Best Performance."
            Write-Log "Visual performance settings applied successfully."
            
            # Open Startup Apps settings
            # Initialize an array to keep track of disabled apps
            $disabledApps = @()

            # Function to disable startup applications from the Startup folder
            function Disable-StartupFolderApps {
                $startupFolder = [System.IO.Path]::Combine($env:USERPROFILE, 'AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup')
                
                if (Test-Path $startupFolder) {
                    $files = Get-ChildItem -Path $startupFolder -ErrorAction SilentlyContinue
                    
                    foreach ($file in $files) {
                        try {
                            # Remove the shortcut
                            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                            $disabledApps += "Removed startup app from Startup folder: $($file.Name)"
                        } catch {
                            Write-Host "Failed to remove startup app: $($file.Name). Reason: $_"
                        }
                    }
                } else {
                    Write-Host "Startup folder does not exist."
                }
            }

            # Function to disable startup applications from the registry
            function Disable-StartupRegistryApps {
                $registryPaths = @(
                    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run",
                    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Run",
                    "HKCU:\Software\Microsoft\Windows\CurrentVersion\RunOnce",
                    "HKLM:\Software\Microsoft\Windows\CurrentVersion\RunOnce"
                )

                foreach ($path in $registryPaths) {
                    # Get all entries in the current registry path
                    $apps = Get-ItemProperty -Path $path -ErrorAction SilentlyContinue
                    
                    if ($apps) {
                        foreach ($app in $apps.PSObject.Properties) {
                            try {
                                # Remove the application from startup
                                Remove-ItemProperty -Path $path -Name $app.Name -ErrorAction Stop
                                $disabledApps += "Disabled startup app from registry: $($app.Name)"
                            } catch {
                                Write-Host "Failed to disable app: $($app.Name). Reason: $_"
                            }
                        }
                    }
                }
            }

            # Function to show what has been disabled
            function Show-DisabledApps {
                if ($disabledApps.Count -gt 0) {
                    Write-Host "nSummary of disabled items:"
                    $disabledApps | ForEach-Object { Write-Host $_ }
                } else {
                    Write-Host "No startup applications were disabled."
                }
            }

            # Execute functions
            Disable-StartupFolderApps
            Disable-StartupRegistryApps
            Show-DisabledApps
           
            # Apply Taskbar settings directly
            # Check Windows version
            $isWin11 = (Get-CimInstance Win32_OperatingSystem).Version -like "10.0.22000*"

            if ($isWin11) {
                # Windows 11 specific settings
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 1
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowPeople" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowWindowsInkWorkspace" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTouchKeyboardButton" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "CortanaEnabled" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "HideNewsAndInterests" -Value 1
            } else {
                # Windows 10 specific settings
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowPeople" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowWindowsInkWorkspace" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTouchKeyboardButton" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "CortanaEnabled" -Value 0
                Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\Explorer" -Name "HideNewsAndInterests" -Value 1 
            }

            # Turn off News and Interests
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Feeds" -Name "ShellFeedsTaskbarViewMode" -Value 2
            
            # Restart Explorer to apply changes
            Stop-Process -Name explorer -Force
            Start-Process explorer

            Write-Log "Performance Settings, Startup Apps, USB devices and Taskbar settings have been applied."            
            Update-Checklist "Performance Settings"
        }
        "3" { 
            Clear-Host
            Write-Host "Adjusting Power Settings..."
            Write-Log "Adjusting Power Settings..."
            powercfg /change standby-timeout-ac 0
            Write-Log "Set standby-timeout-ac to 0."
            powercfg /change monitor-timeout-ac 0
            powercfg /change standby-timeout-dc 0
            powercfg /change monitor-timeout-dc 0
            Write-Log "Power settings adjusted to never sleep or turn off display, both on AC and battery."
            Update-Checklist "Power Settings"
        }
        "4" { 
            Clear-Host
            Write-Host "Configuring Region and Date Settings..."
            Write-Log "Configuring Region and Date Settings..."
            # Log each setting change
            Write-Log "Setting system region to South Africa."
            Set-WinUILanguageOverride -Language "en-ZA"
            Set-WinUserLanguageList -LanguageList "en-ZA" -Force
            Set-WinSystemLocale -SystemLocale "en-ZA"
            Set-Culture -CultureInfo "en-ZA"  # Ensure culture is set

            # Set Home Location to GeoID 209
            Write-Log "Setting home location to GeoID 209."
            Set-WinHomeLocation -GeoId 209

            # Set Region and Locale to South Africa (South Africa - English)
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sCountry -Value "South Africa"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sCountryName -Value "South Africa"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sLanguage -Value "ENG"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sDate -Value "dd/MM/yyyy"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sShortDate -Value "dd/MM/yyyy"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sLongDate -Value "dddd, d MMMM yyyy"
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name sCurrency -Value "R"

            # Set the decimal separator for numbers
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name "sDecimal" -Value "."

            # Set the decimal separator for currency values
            Set-ItemProperty -Path "HKCU:\Control Panel\International" -Name "sMonDecimalSep" -Value "."


            # Set the user language list to ensure the correct culture is applied
            $languageList = New-WinUserLanguageList "en-ZA"
            Set-WinUserLanguageList $languageList -Force

            # Set Time Zone to Harare/Pretoria (UTC+2)
            Write-Log "Setting time zone to Harare/Pretoria."
            tzutil /s "South Africa Standard Time"

            # Sync time by disabling and re-enabling the Windows Time service
            Write-Log "Syncing system time."
            Stop-Service w32time
            Set-Service -Name w32time -StartupType Automatic
            Start-Service w32time
            w32tm /resync

            Write-Log "Region and Date settings have been applied, time zone set to Harare/Pretoria, and decimal symbols set."
            Update-Checklist "Region and Date Settings"
        }
        "5" { 
            Clear-Host
            Write-Host "Creating C.I. Systems & Client Users..."
            Write-Log "Creating C.I. Systems & Client Users..."
            $Username = "C.I. Systems"
            $Password = ConvertTo-SecureString "W1a2J3m4C5d6" -AsPlainText -Force
            New-LocalUser -Name $Username -Password $Password -FullName "C.I. Systems" -Description "Local user for C.I. Systems"
            Add-LocalGroupMember -Group "Administrators" -Member $Username
            Write-Log "User 'C.I. Systems' has been created and added to the Administrators group."
            Update-Checklist "C.I. Systems & Client Users"
        }
        "6" { 
            Clear-Host
            Write-Host "Copying Installation Files..."
            Write-Log "Copying Installation Files..."
            $desktopPath = [System.Environment]::GetFolderPath('Desktop')
            $sourceDir = Join-Path $desktopPath "SmartTrade"
            $destinationDir = "C:\SmtDB\Install"

            try {
                Copy-Item -Path (Join-Path $sourceDir "*") -Destination $destinationDir -Recurse -Force
                Write-Log "Installation files copied from Desktop\SmartTrade to C:\SmtDB\Install."
            } catch {
                Write-Log "An error occurred while copying installation files: $_"
            }
            Update-Checklist "Installation Files Copied"
        }
        "7" { 
            try {
                Clear-Host
                Write-Host "Installing Third Party Software..."
                Write-Log "Installing Third Party Software..."

                # Validate destination directory
                if (-not (Test-RequiredPath $destinationDir "destination directory")) {
                    try {
                        New-Item -Path $destinationDir -ItemType Directory -Force
                    } catch {
                        Write-ErrorLog "Directory Creation" "Failed to create destination directory: $_"
                        return
                    }
                }

                # Download and install WinRAR with validation
                try {
                    if (-not (Test-Path $winRarInstaller)) {
                        Write-Host "Downloading WinRAR..."
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($winRarUrl, $winRarInstaller)
                    }

                    if (Test-Path $winRarInstaller) {
                        $process = Start-Process -FilePath $winRarInstaller -ArgumentList "/S" -Wait -PassThru
                        if ($process.ExitCode -ne 0) {
                            throw "WinRAR installation failed with exit code: $($process.ExitCode)"
                        }
                        Write-Log "WinRAR installed successfully."
                    }
                } catch {
                    Write-ErrorLog "WinRAR Installation" $_
                    return
                }

                # TeamViewer installation with validation
                $installTeamViewer = Read-Host "Do you already have TeamViewer installed? (Y/N)"
                if ($installTeamViewer -eq "N") {
                    try {
                        Write-Host "Downloading TeamViewer..."
                        $webClient = New-Object System.Net.WebClient
                        $webClient.DownloadFile($teamViewerUrl, $teamViewerInstaller)

                        if (Test-Path $teamViewerInstaller) {
                            $process = Start-Process -FilePath $teamViewerInstaller -ArgumentList "/S" -Wait -PassThru
                            if ($process.ExitCode -ne 0) {
                                throw "TeamViewer installation failed with exit code: $($process.ExitCode)"
                            }
                            Write-Log "TeamViewer installed successfully."
                        }
                    } catch {
                        Write-ErrorLog "TeamViewer Installation" $_
                        return
                    }
                }

                # RAR extraction with validation
                try {
                    $rarFiles = Get-ChildItem -Path $installDir -Filter *.rar
                    foreach ($file in $rarFiles) {
                        $process = Start-Process -FilePath "C:\Program Files\WinRAR\WinRAR.exe" `
                            -ArgumentList "x -y `"$($file.FullName)`" `"$installDir`"" -Wait -PassThru
                        if ($process.ExitCode -ne 0) {
                            throw "Failed to extract $($file.Name) with exit code: $($process.ExitCode)"
                        }
                        Remove-Item $file.FullName -Force
                    }
                    Write-Log "All .rar files processed successfully."
                } catch {
                    Write-ErrorLog "RAR Extraction" $_
                }
            } catch {
                Write-ErrorLog "Third Party Software Installation" $_
            }
            Update-Checklist "Third Party Software"
        }
        "8"  { 
            Clear-Host
            Write-Host "Extracting Fonts and Running BAT File..."
            Write-Log "Starting Font Extraction and BAT Execution..."

            try {
                # Get the patch number from input
                $patchNumber = Read-Host "Enter the patch number"
                $patchFolderPattern = "Patch $patchNumber*"
                $sourceFolder = Get-ChildItem -Path "C:\SmtDB\Install" -Filter $patchFolderPattern | Select-Object -First 1

                if ($sourceFolder) {
                    $patchFolderPath = $sourceFolder.FullName
                    $toolsFolder = Join-Path $patchFolderPath "Tools"
                    $fontsRarFile = Join-Path $toolsFolder "Fonts.rar"  # Path to Fonts.rar file
                    $extractFolder = $toolsFolder  # Extract destination
                    $fontsFolder = Join-Path $toolsFolder "Fonts"  # Extracted Fonts folder path
                    $sqlserverbat = Join-Path $toolsFolder "OpenSQLServerPort.bat"  # BAT file path

                    # Extract Fonts.rar if it exists
                    if (Test-Path $fontsRarFile) {
                        try {
                            Write-Log "Extracting Fonts.rar..."
                            Start-Process -FilePath "C:\Program Files\WinRAR\WinRAR.exe" -ArgumentList "x -y `"$fontsRarFile`" `"$extractFolder`"" -Wait -ErrorAction Stop
                            Write-Log "Fonts.rar extracted successfully."
                        } catch {
                            Write-Log "Error extracting Fonts.rar: $_"
                            return
                        }
                    } else {
                        Write-Log "Fonts.rar file not found in $toolsFolder."
                        return
                    }

                    # Install fonts if Fonts folder exists
                    if (Test-Path $fontsFolder) {
                        $fontFiles = Get-ChildItem -Path $fontsFolder -Filter *.ttf  # Assume .ttf format
                        foreach ($font in $fontFiles) {
                            $fontDestination = "C:\Windows\Fonts\$($font.Name)"
                            if (-not (Test-Path $fontDestination)) {
                                Copy-Item -Path $font.FullName -Destination "C:\Windows\Fonts\" -Force
                                $fontName = [System.IO.Path]::GetFileNameWithoutExtension($font.Name)
                                $fontRegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
                                Set-ItemProperty -Path $fontRegistryPath -Name "$fontName (TrueType)" -Value $font.Name
                                Write-Log "Installed font: $fontName."
                            }
                        }
                        Write-Log "All fonts installed successfully."
                    } else {
                        Write-Log "Fonts folder not found after extraction."
                        return
                    }

                    # Run OpenSQLServerPort.bat if it exists
                    if (Test-Path $sqlserverbat) {
                        Start-Process -NoNewWindow -FilePath $sqlserverbat -Wait
                        Write-Log "OpenSQLServerPort.bat executed."
                    } else {
                        Write-Log "OpenSQLServerPort.bat file not found."
                    }
                } else {
                    Write-Log "Patch folder not found for pattern: $patchFolderPattern."
                }
            } catch {
                Write-Log "An error occurred: $_"
            }

            Update-Checklist "Fonts and BAT"
        }
        "9"  {
            Clear-Host
            Write-Host "Installing SQL Server and SSMS..."
            Write-Log "Installing SQL Server and SSMS..."
            try {
                # Prompt for the patch number
                $patchNumber = Read-Host "Enter the patch number"
                
                # Set the download links
                $ssmsUrl = "https://aka.ms/ssmsfullsetup"
                $sqlUrl = "https://download.microsoft.com/download/5/1/4/5145fe04-4d30-4b85-b0d1-39533663a2f1/SQL2022-SSEI-Expr.exe"
    
                # Set destination directory
                $destinationDir = "C:\SmtDB\Install\"
    
                # Ensure the destination directory exists
                if (-not (Test-Path $destinationDir)) {
                    New-Item -Path $destinationDir -ItemType Directory
                }
    
                # Define file paths for downloads
                $sqlSetupExe = Join-Path $env:USERPROFILE "Downloads\SQL2022-SSEI-Expr.exe"
                $ssmsExe = Join-Path $env:USERPROFILE "Downloads\SSMS-Setup-ENU.exe"
    
                # Download SQL Server Express
                if (-not (Test-Path $sqlSetupExe)) {
                    Write-Log "Downloading SQL Server Express..."
                    Invoke-WebRequest -Uri $sqlUrl -OutFile $sqlSetupExe
                } else {
                    Write-Log "SQL Server Express installer already exists in Downloads folder."
                }
    
                # Download SSMS
                if (-not (Test-Path $ssmsExe)) {
                    Write-Log "Downloading SSMS..."
                    Invoke-WebRequest -Uri $ssmsUrl -OutFile $ssmsExe
                } else {
                    Write-Log "SSMS installer already exists in Downloads folder."
                }
    
                # Move SQL installer to destination folder
                if (Test-Path $sqlSetupExe) {
                    Move-Item -Path $sqlSetupExe -Destination (Join-Path $destinationDir "SQL2022-SSEI-Expr.exe") -Force
                    Write-Log "SQL Server Express installer moved to $destinationDir."
                }
    
                # Move SSMS installer to destination folder
                if (Test-Path $ssmsExe) {
                    Move-Item -Path $ssmsExe -Destination (Join-Path $destinationDir "SSMS-Setup-ENU.exe") -Force
                    Write-Log "SSMS installer moved to $destinationDir."
                }
    
                # Run SQL Server Express setup
                if (Test-Path (Join-Path $destinationDir "SQL2022-SSEI-Expr.exe")) {
                    Start-Process -FilePath (Join-Path $destinationDir "SQL2022-SSEI-Expr.exe") -Wait
                    Write-Log "SQL Server Express setup started. Please follow the on-screen instructions."   
                }

                # Move SQL Express installer to destination folder
                $sqlInstallerPath = Join-Path $env:USERPROFILE "Downloads\SQLEXPR_x64_ENU.exe"
                if (Test-Path $sqlInstallerPath) {
                    Move-Item -Path $sqlInstallerPath -Destination (Join-Path $destinationDir "SQLEXPR_x64_ENU.exe") -Force
                    Write-Log "SQL Server Express installer moved to $destinationDir."
                } else {
                    Write-Log "SQL Server Express installer not found in Downloads folder after installation."
                }

                # Run the SQL Server Express installer
                if (Test-Path (Join-Path $destinationDir "SQLEXPR_x64_ENU.exe")) {
                    Start-Process -FilePath (Join-Path $destinationDir "SQLEXPR_x64_ENU.exe") -Wait
                    Write-Log "SQL installation started. Please follow the on-screen instructions."
                }
    
                # Run SSMS setup
                if (Test-Path (Join-Path $destinationDir "SSMS-Setup-ENU.exe")) {
                    Start-Process -FilePath (Join-Path $destinationDir "SSMS-Setup-ENU.exe") -Wait
                    Write-Log "SSMS installation started. Please follow the on-screen instructions."
                }
    
            } catch {
                Write-Log "An error occurred while installing SQL or SSMS: $_"
            }
            Update-Checklist "SQL Installation"
        }
        "10" {
            Clear-Host
            Write-Host "Running User Scripts & Fixes..."
            Write-Log "Running User Scripts & Fixes..."
            try {
                # Prompt for the patch number
                $patchNumber = Read-Host "Enter the patch number"
                
                # Use a wildcard to match the folder with the year
                $scriptsFolderPattern = "Patch $patchNumber*"
                $scriptsFolder = Get-ChildItem -Path "C:\SmtDB\Install" -Filter $scriptsFolderPattern | Where-Object { $_.PSIsContainer } | Select-Object -First 1
     
                if ($scriptsFolder) {
                    $scriptsFolderPath = Join-Path $scriptsFolder.FullName "Scripts"
                     
                    # Open the folder
                    if (Test-Path $scriptsFolderPath) {
                        Start-Process explorer.exe $scriptsFolderPath
                        Write-Log "Opened scripts folder: $scriptsFolderPath"
                    } else {
                        Write-Log "Scripts folder not found: $scriptsFolderPath"
                    }
                } else {
                    Write-Log "Patch folder not found."
                }
            } catch {
                Write-Log "An error occurred: $_"
            }  # Closing brace added here
            Update-Checklist "User Scripts & Fixes"
        }
        "11" { 
            Clear-Host
            Write-Host "Registering DLLs..."
            Write-Log "Registering DLLs..."
            try {
                $patchNumber = Read-Host "Enter the patch number"
                $patchFolderPattern = "Patch $patchNumber*"
                $sourceFolder = Get-ChildItem -Path "C:\SmtDB\Install" -Filter $patchFolderPattern | Select-Object -First 1

                if ($sourceFolder) {
                    $sourceFolderPath = Join-Path $sourceFolder.FullName "CISystems"
                    $destinationFolder = "C:\Program Files (x86)\CISystems"

                    if (Test-Path $sourceFolderPath) {
                        if (Test-Path $destinationFolder) {
                            Remove-Item -Path $destinationFolder -Recurse -Force
                        }
                        Copy-Item -Path $sourceFolderPath -Destination $destinationFolder -Recurse -Force
                    } else {
                        Write-Host "Source folder not found: $sourceFolderPath"
                        Write-Log "Source folder not found: $sourceFolderPath"
                        return
                    }

                    Start-Sleep -Seconds 2
                    Set-Location -Path $destinationFolder
                    Start-Process -NoNewWindow -FilePath "cmd.exe" -ArgumentList "/c regalldll's" -Wait
                    Write-Log "DLLs registered successfully using the 'regalldll's' command."
                } else {
                    Write-Host "Patch folder not found."
                    Write-Log "Patch folder not found."
                }
            } catch {
                Write-Log "An error occurred while registering DLLs: $_"
            }
            Update-Checklist "DLLs Registered"
        }
        "12" { 
            Clear-Host
            Write-Host "Creating COM+ Applications..."
            Write-Log "Creating COM+ Applications..."
            try {
                $catalog = New-Object -ComObject COMAdmin.COMAdminCatalog
                $applications = $catalog.GetCollection("Applications")
                $applications.Populate()

                $comAppNames = @(
                    "SMT Customers",
                    "SMT General",
                    "SMT Printing",
                    "SMT Procedures",
                    "SMT Products",
                    "SMT Reports",
                    "SMT Sales",
                    "SMT Settings",
                    "SMT Suppliers",
                    "SMT Sync"
                )

                $comAppDirs = @{
                    "SMT Customers" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Customers\"
                    "SMT General" = "C:\Program Files (x86)\CISystems\ComPlus Applications\General\"
                    "SMT Printing" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Printing\"
                    "SMT Procedures" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Procedures\"
                    "SMT Products" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Products\"
                    "SMT Reports" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Reports\"
                    "SMT Sales" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Sales\"
                    "SMT Settings" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Settings\"
                    "SMT Suppliers" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Suppliers\"
                    "SMT Sync" = "C:\Program Files (x86)\CISystems\ComPlus Applications\Sync\"
                }

                foreach ($appName in $comAppNames) {
                    $app = $applications | Where-Object { $_.Value("Name") -eq $appName }

                    if ($null -eq $app) {
                        Write-Host "Application '$appName' not found. Creating it..."
                        try {
                            $app = $applications.Add()
                            $app.Value("Name") = $appName
                            $applications.SaveChanges()
                            Write-Log "Application '$appName' created successfully."
                        } catch {
                            Write-Host "Failed to create application '$appName': $_"
                            Write-Log "Failed to create application '$appName': $_"
                            continue
                        }

                        $applications.Populate()
                        $app = $applications | Where-Object { $_.Value("Name") -eq $appName }
                    } else {
                        Write-Host "Application '$appName' found."
                    }

                    if ($comAppDirs.ContainsKey($appName)) {
                        $componentDir = $comAppDirs[$appName]
                        $dllFiles = Get-ChildItem -Path $componentDir -Filter *.dll

                        foreach ($dll in $dllFiles) {
                            $filePath = $dll.FullName
                            Write-Log "Registering component: $filePath"
                            try {
                                $catalog.InstallComponent($appName, $filePath, "", "")
                                Write-Log "Component registered successfully: $filePath"
                            } catch {
                                Write-Log "Error registering component '$filePath': $_"
                            }
                        }
                    } else {
                        Write-Host "Directory for '$appName' not found."
                        Write-Log "Directory for '$appName' not found."
                    }

                    $roles = $applications.GetCollection("Roles", $app.Key)
                    $roles.Populate()

                    $roleFound = $false
                    foreach ($role in $roles) {
                        if ($role.Name -eq "CreatorOwner") {
                            $roleFound = $true
                            break
                        }
                    }

                    if (-not $roleFound) {
                        $role = $roles.Add()
                        $role.Value("Name") = "CreatorOwner"
                        $roles.SaveChanges()
                        Write-Log "Created 'CreatorOwner' role."
                    }

                    $usersInRole = $roles.GetCollection("UsersInRole", $role.Key)
                    $usersInRole.Populate()

                    $user = $usersInRole.Add()
                    $user.Value("User") = "Everyone"
                    $usersInRole.SaveChanges()

                    Write-Log "Successfully added 'Everyone' to the 'CreatorOwner' role."
                }

                Write-Log "COM+ Applications created, components registered, and roles configured successfully."
                Start-Process "dcomcnfg"
            } catch {
                Write-Log "An error occurred: $_"
            }
            Update-Checklist "COM+ Applications"
        }
        "13" { 
            Clear-Host
            Write-Host "Bypassing COM+ Firewall..."
            Write-Log "Bypassing COM+ Firewall..."
            try {
                $basePath = "C:\Program Files\Microsoft SQL Server"
                $sqlDirs = Get-ChildItem -Path $basePath -Directory

                foreach ($dir in $sqlDirs) {
                    $sqlPath = Join-Path -Path $dir.FullName -ChildPath "MSSQL\Binn\sqlservr.exe"
                    if (Test-Path -Path $sqlPath) {
                        $ruleName = "SQL Server App"
                        New-NetFirewallRule -DisplayName $ruleName `
                                            -Direction Inbound `
                                            -Program $sqlPath `
                                            -Action Allow `
                                            -Profile Domain,Private,Public
                    }
                }

                New-NetFirewallRule -DisplayName "Incoming DCOM Connections" `
                                    -Direction Inbound `
                                    -Protocol TCP `
                                    -LocalPort RPC `
                                    -RemotePort 1024-65535 `
                                    -Action Allow `
                                    -Profile Domain,Private,Public `
                                    -RemoteAddress Any

                New-NetFirewallRule -DisplayName "SQL Browser App" `
                                    -Direction Inbound `
                                    -Program "C:\Program Files (x86)\Microsoft SQL Server\90\Shared\sqlbrowser.exe" `
                                    -Action Allow `
                                    -Profile Domain,Private,Public

                Write-Log "All firewall rules were successfully added."
            } catch {
                Write-Log "An error occurred while adding the firewall rules: $_"
            }
            Update-Checklist "COM+ Firewall"
        }
        "14" { 
            Clear-Host
            Write-Host "Creating Shortcut to Desktop..."
            Write-Log "Creating Shortcut to Desktop..."
            try {
                $desktopPath = [System.Environment]::GetFolderPath('Desktop')
                $shortcutPath = Join-Path $desktopPath "SmartTrade.lnk"
                $targetPath = "C:\Program Files (x86)\CISystems\Smart-Trade Retail Management - Professional Edition\Smart-Trade.exe"
                $workingDirectory = "C:\SmtDB"

                $wshShell = New-Object -ComObject WScript.Shell
                $shortcut = $wshShell.CreateShortcut($shortcutPath)
                $shortcut.TargetPath = $targetPath
                $shortcut.WorkingDirectory = $workingDirectory
                $shortcut.Save()

                Write-Log "Shortcut created successfully on the Desktop."
            } catch {
                Write-Log "An error occurred while creating the shortcut: $_"
            }
            Update-Checklist "Desktop Shortcut"
        }
        
        "0" { 
            Write-Host "Generating final checklist report..."
            Export-ChecklistReport -TechnicianName $technicianName -ClientName $clientName
            Write-Host "Exiting..."
        }
        "R" {  # {{ edit_2 }}
            Clear-Host
            Write-Host "Installing Radmin VPN..."
            Write-Log "Installing Radmin VPN..."
            try {
                $radminUrl = "https://download.radmin-vpn.com/download/files/Radmin_VPN_1.4.4642.1.exe"  # URL for Radmin VPN
                $radminInstaller = "C:\SmtDB\Install\RadminVPN.exe"  # Path to save the installer

                # Download Radmin VPN if not present
                if (-not (Test-Path $radminInstaller)) {
                    Write-Host "Downloading Radmin VPN..."
                    Invoke-WebRequest -Uri $radminUrl -OutFile $radminInstaller
                }

                # Install Radmin VPN silently
                if (Test-Path $radminInstaller) {
                    Start-Process -FilePath $radminInstaller -ArgumentList "/S" -Wait
                    Write-Log "Radmin VPN installed successfully."
                } else {
                    Write-Log "Radmin VPN installer not found."
                }
            } catch {
                Write-Log "An error occurred while installing Radmin VPN: $_"
            }
            Update-Checklist "Radmin VPN"
        }
        default {
            Write-Host "Invalid choice, please try again."
        }
        "D" {
            Clear-Host
            Write-Host "Opening Device Manager"
            Start-Process devmgmt.msc  # Corrected command to open Device Manager
        }
        "F" {
            Clear-Host
            Write-Host "Opening Region Settings..."
            Start-Process "control.exe" -ArgumentList "intl.cpl"
        }
        "Dance" {
            # Define the command to open cmd and run curl
            $cmdCommand = 'Curl Parrot.live'

            # Start a new cmd window and execute the command
            Start-Process cmd.exe -ArgumentList "/k mode con: cols=60 lines=20 & $cmdCommand"
        }
    }
    Start-Sleep -Seconds 2
} while ($choice -ne "0")

# Add cleanup and exit handling
trap {
    Write-ErrorLog "Fatal Error" $_
    if ($Global:LastError) {
        Write-Host "`nLast Error: $Global:LastError" -ForegroundColor Yellow
    }
    Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

