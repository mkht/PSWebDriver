function Set-WebDriverEnvironment {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [ValidateSet('Chrome', 'Firefox', 'Edge', 'EdgeChromium', 'HeadlessChrome', 'HeadlessFirefox', 'IE', 'InternetExplorer')]
        [String[]]
        $TargetBrowser,

        [Parameter(ValueFromPipelineByPropertyName)]
        [string]
        $BrowserBinaryPath
    )

    Process {
        foreach ($target in $TargetBrowser) {
            $strictTargetBrowser = switch ($target) {
                'Chrome' { 'Chrome'; break }
                'Firefox' { 'Firefox'; break }
                'Edge' { 'Edge'; break }
                'EdgeChromium' { 'EdgeChromium'; break }
                'HeadlessChrome' { 'Chrome'; break }
                'HeadlessFirefox' { 'Firefox'; break }
                'IE' { 'InternetExplorer'; break }
                'InternetExplorer' { 'InternetExplorer'; break }
                Default {}
            }

            switch ($strictTargetBrowser) {
                'Chrome' { 
                    Set-ChromeEnvironment -BrowserBinaryPath $BrowserBinaryPath
                    break
                }
                'Firefox' { 
                    Set-FirefoxEnvironment
                    break
                }
                'Edge' { 
                    Set-EdgeEnvironment
                    break
                }
                'EdgeChromium' { 
                    Set-EdgeChromiumEnvironment -BrowserBinaryPath $BrowserBinaryPath
                    break
                }
                'InternetExplorer' { 
                    Set-InternetExplorerEnvironment
                    break
                }
                Default {}
            }
        }
    }

}


function Set-ChromeEnvironment {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $BrowserBinaryPath
    )
    
    begin {
        $DriverSavePath = Join-Path $env:USERPROFILE '.webdriver\chromedriver'
        $WebDriverQueryLocation = 'https://chromedriver.storage.googleapis.com/LATEST_RELEASE'
        $WebDriverDownloadLocation = 'https://chromedriver.storage.googleapis.com/'
        $SearchBrowserPaths = if ($BrowserBinaryPath) {
            @($BrowserBinaryPath)
        }
        else {
            @(
                (Join-Path $env:APPDATA '\Local\Google\Chrome\Application\chrome.exe'),
                (Join-Path $env:ProgramFiles '\Google\Chrome\Application\chrome.exe'),
                (Join-Path ${env:ProgramFiles(x86)} '\Google\Chrome\Application\chrome.exe')
            )
        }
    }
    
    process {
        $BrowserVersion = $null
        if ($BrowserBinaryPath) {
            if (-not (Test-Path -LiteralPath $BrowserBinaryPath -PathType Leaf)) {
                Write-Error "$BrowserBinaryPath is not exist"
                return
            }
        }
        
        # Search chrome.exe
        foreach ($testPath in $SearchBrowserPaths) {
            if (Test-Path -LiteralPath $testPath -PathType Leaf) {
                $VersionInfo = (Get-Item $testPath).VersionInfo.ProductVersion
                if ([version]::TryParse($VersionInfo, [ref]$BrowserVersion)) {
                    Write-Verbose ('The chrome.exe is found in "{0}"' -f $testPath)
                    Write-Verbose ('Target browser version is {0}' -f $BrowserVersion.ToString())
                    break
                }
            }
        }

        # chrome.exe does not found
        if ($null -eq $BrowserVersion) {
            Write-Error 'chrome.exe does not found on the system. You should install Google Chrome or Specify a path of the chrome.exe to the BrowserBinaryPath parameter.'
            return
        }

        # Query a recommend version of the webdriver
        $QueryURL = $WebDriverQueryLocation + ('_{0}.{1}.{2}' -f $BrowserVersion.Major, $BrowserVersion.Minor, $BrowserVersion.Build)
        $DriverVersion = (Invoke-WebRequest -Uri $QueryURL -UseBasicParsing -ErrorAction Ignore).Content

        # recommend driver does not found
        if (-not $DriverVersion) {
            Write-Error 'Could not found proper version of the WebDriver. Confirm network connection'
            return
        }

        # Test current driver is exist or not, and version is proper or not
        $IsCurrentDriverPresent = $false
        if (Test-Path -LiteralPath (Join-Path $DriverSavePath 'chromedriver.exe') -PathType Leaf) {
            $currentVersion = & (Join-Path $DriverSavePath 'chromedriver.exe') --version
            Write-Verbose $currentVersion
            if ($currentVersion -match $DriverVersion) {
                $IsCurrentDriverPresent = $true
                Write-Verbose 'Current driver is proper version.'
            }
            else {
                $IsCurrentDriverPresent = $false
                Write-Verbose 'Current driver is not proper version.'
            }
        }

        # download web driver
        if (-not $IsCurrentDriverPresent) {
            $TempPath = Join-Path $env:TEMP 'chromedriver_win32.zip'
            $WebDriverDownloadLocation = $WebDriverDownloadLocation + ('{0}/{1}' -f $DriverVersion, 'chromedriver_win32.zip')
            Write-Verbose "Download driver from $WebDriverDownloadLocation"
            try {
                Invoke-WebRequest -Uri $WebDriverDownloadLocation -OutFile $TempPath -UseBasicParsing -ErrorAction Stop
                Expand-Archive -Path $TempPath -DestinationPath $DriverSavePath -Force -ErrorAction Stop
            }
            catch {
                Write-Error 'Failed to download driver'
                return
            }
            finally {
                Remove-Item -Path $TempPath -Force -ErrorAction Ignore
            }
            Write-Verbose "Chrome driver is downloaded to $DriverSavePath"
        }
    }
    
    end {
        # Add PATH
        $CurrentPath = $env:PATH -split ';'
        if (-not ($DriverSavePath -in $CurrentPath)) {
            Add-PathEnvironment -Path $DriverSavePath -Scope User -ErrorAction SilentlyContinue
        }
    }
}


function Set-InternetExplorerEnvironment {
    [CmdletBinding()]
    param (
    )
    
    begin {
        $DriverSavePath = Join-Path $env:USERPROFILE '.webdriver\iedriver'
        $WebDriverDownloadLocation = 'https://selenium-release.storage.googleapis.com/3.150/IEDriverServer_Win32_3.150.1.zip'
    }
    
    process {
        # Test current driver exist or not
        $IsCurrentDriverPresent = $false
        if (Test-Path -LiteralPath (Join-Path $DriverSavePath 'IEDriverServer.exe') -PathType Leaf) {
            $currentVersion = & (Join-Path $DriverSavePath 'IEDriverServer.exe') --version
            Write-Verbose $currentVersion
            $IsCurrentDriverPresent = $true
            Write-Verbose 'Current driver exist.'
        }
        else {
            $IsCurrentDriverPresent = $false
            Write-Verbose 'Current driver does not exist.'
        }

        # download web driver
        if (-not $IsCurrentDriverPresent) {
            $TempPath = Join-Path $env:TEMP 'iedriver.zip'
            Write-Verbose "Download driver from $WebDriverDownloadLocation"
            try {
                Invoke-WebRequest -Uri $WebDriverDownloadLocation -OutFile $TempPath -UseBasicParsing -ErrorAction Stop
                Expand-Archive -Path $TempPath -DestinationPath $DriverSavePath -Force -ErrorAction Stop
            }
            catch {
                Write-Error 'Failed to download driver'
                return
            }
            finally {
                Remove-Item -Path $TempPath -Force -ErrorAction Ignore
            }
            Write-Verbose "IE driver is downloaded to $DriverSavePath"
        }

        # Change IE settings
        # Stop iexplore processes
        Get-Process -Name iexplore -ErrorAction Ignore | Stop-Process
        
        # Enable protected mode on all zones
        Write-Verbose 'Change IE settings: Enable protected mode on all zones'
        $ZoneRegistry = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\'
        (1..4) | % {
            Set-ItemProperty -Path (Join-Path $ZoneRegistry $_) -Name 2500 -Value 0 -Force
        }
        
        # Disable Enchanced protected mode
        Write-Verbose 'Change IE settings: Disable Enchanced protected mode'
        $MainRegistry = 'HKCU:\Software\Microsoft\Internet Explorer\Main'
        Set-ItemProperty -Path $MainRegistry -Name 'Isolation' -Value 'PMIL' -Force

        # Enable URL-based basic authentication
        # https://docs.microsoft.com/en-us/troubleshoot/browsers/name-and-password-not-supported-in-website-address
        Write-Verbose 'Change IE settings: Enable URL-based basic authentication'
        $FEATURERegistry = 'HKCU:\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_HTTP_USERNAME_PASSWORD_DISABLE'
        if (-not (Test-Path -LiteralPath $FEATURERegistry)) {
            $null = New-Item -Path $FEATURERegistry -Force
        }
        Set-ItemProperty -Path $FEATURERegistry -Name 'iexplore.exe' -Value 0 -Force
    }
    
    end {
        # Add PATH
        $CurrentPath = $env:PATH -split ';'
        if (-not ($DriverSavePath -in $CurrentPath)) {
            Add-PathEnvironment -Path $DriverSavePath -Scope User -ErrorAction SilentlyContinue
        }
    }
}

function Set-EdgeEnvironment {
    [CmdletBinding()]
    param (
    )
    
    process {
        $Capability = Get-WindowsCapability -Name 'Microsoft.WebDriver~~~~0.0.1.0' -Online -ErrorAction Stop
        if ($Capability.State -eq 'Installed') {
            Write-Verbose 'Edge driver is already installed.'
            return
        }
        else {
            Write-Verbose 'Installing Edge driver.'
            $result = Add-WindowsCapability -Name $Capability.Name -Online
            if ($result.RestartNeeded) {
                Write-Warning 'You should restart computer to enable this feature.'
            }
        }
    }
}

function Set-EdgeChromiumEnvironment {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $BrowserBinaryPath
    )
    
    begin {
        $DriverSavePath = Join-Path $env:USERPROFILE '.webdriver\edgedriver'
        $WebDriverQueryLocation = 'https://msedgewebdriverstorage.z22.web.core.windows.net/'
        $WebDriverDownloadLocation = 'https://msedgedriver.azureedge.net/'
        $SearchBrowserPaths = if ($BrowserBinaryPath) {
            @($BrowserBinaryPath)
        }
        else {
            @(
                (Join-Path ${env:ProgramFiles(x86)} '\Microsoft\Edge\Application\msedge.exe'),
                (Join-Path $env:ProgramFiles '\Microsoft\Edge\Application\msedge.exe')
            )
        }

        $Platform = switch ($env:PROCESSOR_ARCHITECTURE) {
            'AMD64' { 'win64' }
            'x86' { 'win32' }
            Default {
                Write-Error 'Invalid platfrom'
                return
            }
        }
    }
    
    process {
        $BrowserVersion = $null
        if ($BrowserBinaryPath) {
            if (-not (Test-Path -LiteralPath $BrowserBinaryPath -PathType Leaf)) {
                Write-Error "$BrowserBinaryPath is not exist"
                return
            }
        }
        
        # Search msedge.exe
        foreach ($testPath in $SearchBrowserPaths) {
            if (Test-Path -LiteralPath $testPath -PathType Leaf) {
                $VersionInfo = (Get-Item $testPath).VersionInfo.ProductVersion
                if ([version]::TryParse($VersionInfo, [ref]$BrowserVersion)) {
                    Write-Verbose ('The msedge.exe is found in "{0}"' -f $testPath)
                    Write-Verbose ('Target browser version is {0}' -f $BrowserVersion.ToString())
                    break
                }
            }
        }

        # msedge.exe does not found
        if ($null -eq $BrowserVersion) {
            Write-Error 'msedge.exe does not found on the system. You should install Edge (Chromium) or Specify a path of the msedge.exe to the BrowserBinaryPath parameter.'
            return
        }

        # Query webdriver url
        $QueryURL = $WebDriverQueryLocation + ('?prefix={0}' -f $BrowserVersion.ToString())
        $Request = Invoke-WebRequest -Uri $QueryURL -UseBasicParsing -ErrorAction Ignore
        if ($Request.StatusCode -ne 200) {
            Write-Error 'Could not found proper version of the WebDriver. Confirm network connection'
            return
        }

        # Test current driver is exist or not, and version is proper or not
        $IsCurrentDriverPresent = $false
        if (Test-Path -LiteralPath (Join-Path $DriverSavePath 'msedgedriver.exe') -PathType Leaf) {
            $currentVersion = & (Join-Path $DriverSavePath 'msedgedriver.exe') --version
            Write-Verbose $currentVersion
            if ($currentVersion -match $BrowserVersion.ToString()) {
                $IsCurrentDriverPresent = $true
                Write-Verbose 'Edge driver is proper version.'
            }
            else {
                $IsCurrentDriverPresent = $false
                Write-Verbose 'Current driver is not proper version.'
            }
        }

        # download web driver
        if (-not $IsCurrentDriverPresent) {
            $TempPath = Join-Path $env:TEMP 'edgedriver.zip'
            $WebDriverDownloadLocation = $WebDriverDownloadLocation + ('{0}/edgedriver_{1}.zip' -f $BrowserVersion.ToString(), $Platform)
            Write-Verbose "Download driver from $WebDriverDownloadLocation"
            try {
                Invoke-WebRequest -Uri $WebDriverDownloadLocation -OutFile $TempPath -UseBasicParsing -ErrorAction Stop
                Expand-Archive -Path $TempPath -DestinationPath $DriverSavePath -Force -ErrorAction Stop
            }
            catch {
                Write-Error 'Failed to download driver'
                return
            }
            finally {
                Remove-Item -Path $TempPath -Force -ErrorAction Ignore
            }
            Write-Verbose "Edge driver is downloaded to $DriverSavePath"
        }
    }
    
    end {
        # Add PATH
        $CurrentPath = $env:PATH -split ';'
        if (-not ($DriverSavePath -in $CurrentPath)) {
            Add-PathEnvironment -Path $DriverSavePath -Scope User -ErrorAction SilentlyContinue
        }
    }
}


function Set-FirefoxEnvironment {
    [CmdletBinding()]
    param (
    )
    
    begin {
        $DriverSavePath = Join-Path $env:USERPROFILE '.webdriver\geckodriver'
        $WebDriverQueryLocation = 'https://api.github.com/repos/mozilla/geckodriver/releases/latest'
        $Platform = switch ($env:PROCESSOR_ARCHITECTURE) {
            'AMD64' { 'win64' }
            'x86' { 'win32' }
            Default {
                Write-Error 'Invalid platfrom'
                return
            }
        }
    }
    
    process {
        # Query latest driver url
        $LatestRelease = (Invoke-WebRequest -Uri $WebDriverQueryLocation -UseBasicParsing -ErrorAction Ignore).Content | ConvertFrom-Json
        if (-not $LatestRelease.name) {
            Write-Error 'Could not found latest release of the GeckoDriver. Confirm network connection'
            return
        }
        $AssetUrl = $LatestRelease.assets_url
        $LatestAssets = (Invoke-WebRequest -Uri $AssetUrl -UseBasicParsing -ErrorAction Ignore).Content | ConvertFrom-Json
        if (-not $LatestAssets) {
            Write-Error 'Could not found latest release of the GeckoDriver. Confirm network connection'
            return
        }

        # Test current driver is exist or not, and version is proper or not
        $IsCurrentDriverPresent = $false
        if (Test-Path -LiteralPath (Join-Path $DriverSavePath 'geckodriver.exe') -PathType Leaf) {
            $currentVersion = (& (Join-Path $DriverSavePath 'geckodriver.exe') --version)[0]
            Write-Verbose $currentVersion
            if ($currentVersion -match $LatestRelease.name) {
                $IsCurrentDriverPresent = $true
                Write-Verbose 'Current driver is latest version.'
            }
            else {
                $IsCurrentDriverPresent = $false
                Write-Verbose 'Current driver is not latest version.'
            }
        }

        # download web driver
        if (-not $IsCurrentDriverPresent) {
            $TempPath = Join-Path $env:TEMP 'geckodriver.zip'
            $WebDriverDownloadLocation = ($LatestAssets | Where-Object { $_.name -match $Platform } | Select-Object -ExpandProperty browser_download_url)
            Write-Verbose "Download driver from $WebDriverDownloadLocation"
            try {
                Invoke-WebRequest -Uri $WebDriverDownloadLocation -OutFile $TempPath -UseBasicParsing -ErrorAction Stop
                Expand-Archive -Path $TempPath -DestinationPath $DriverSavePath -Force -ErrorAction Stop
            }
            catch {
                Write-Error 'Failed to download driver'
                return
            }
            finally {
                Remove-Item -Path $TempPath -Force -ErrorAction Ignore
            }
            Write-Verbose "Gecko driver is downloaded to $DriverSavePath"
        }
    }
    
    end {
        # Add PATH
        $CurrentPath = $env:PATH -split ';'
        if (-not ($DriverSavePath -in $CurrentPath)) {
            Add-PathEnvironment -Path $DriverSavePath -Scope User -ErrorAction SilentlyContinue
        }
    }
}


function Send-SettingChange {
    if (-not ('Win32.NativeMethods' -as [type])) {
        Add-Type -Namespace Win32 -Name NativeMethods -MemberDefinition @'
[DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, string lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);
'@
    }

    $HWND_BROADCAST = [IntPtr]0xffff
    $WM_SETTINGCHANGE = 0x001A
    $result = [UIntPtr]::Zero

    $SendMessageResult = [Win32.NativeMethods]::SendMessageTimeout($HWND_BROADCAST, $WM_SETTINGCHANGE, [UIntPtr]::Zero, 'Environment', 0x0002, 5000, [ref]$result)

    if ($SendMessageResult -eq 0) {
        Write-Error -Message 'SendMessageTimeout returns error'
    }
}

function Add-PathEnvironment {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline)]
        [string[]]$Path,

        [Parameter(Position = 1)]
        [ValidateSet('User', 'Machine', 'Process')]
        [Alias('Target')]
        [string]$Scope = 'Process'
    )

    Begin {
        if ($Scope -eq 'Machine') {
            # Check administrator privileges
            $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
            if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
                $er = [System.Management.Automation.ErrorRecord]::new(
                    'The requested operation requires administrator privileges.',
                    'Add-PathEnvironment',
                    [System.Management.Automation.ErrorCategory]::PermissionDenied,
                    $null
                )
                $PSCmdlet.ThrowTerminatingError($er)
            }
        }
    }

    Process {
        foreach ($p in $Path) {
            Write-Verbose ('Add {0} to the PATH' -f $p)
            $EnvPath = [Environment]::GetEnvironmentVariable('PATH', $Scope) -split ';'

            if ($p -in $EnvPath) {
                Write-Error -Message 'Specified path already exists in the PATH variable.'
                return
            }
            
            $EnvPath += $p
            [Environment]::SetEnvironmentVariable('PATH', ($EnvPath -join ';'), $Scope)

            if ($Scope -ne 'Process') {
                # Broadcast WM_SETTINGCHANGE
                Send-SettingChange
            }
        }
    }
}
