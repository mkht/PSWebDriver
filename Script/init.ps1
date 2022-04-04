#Require -Version 5.0

$PSModuleRoot = Split-Path $PSScriptRoot -Parent
# $LibPath = Join-Path $PSModuleRoot '\Lib'
Write-Debug ('$PSModuleRoot:{0}' -f $PSModuleRoot)

# Load Selenium WebDriver DLL
if (!("OpenQA.Selenium.By" -as [type])) {
    if (!($SeleniumPath = Resolve-Path "$PSModuleRoot\Lib\Selenium.WebDriver.*\lib\netstandard2.0" -ea SilentlyContinue)) {
        Write-Error "Couldn't find WebDriver.dll"
    }
    # Load Selenium
    try {
        Add-Type -Path (Join-Path $SeleniumPath WebDriver.dll) -ErrorAction Stop
    }
    catch {
        Write-Error "Couldn't load Selenium WebDriver"
    }
}

if (("OpenQA.Selenium.By" -as [type]) -and !("OpenQA.Selenium.Support.UI.SelectElement" -as [type])) {
    if (!($SeleniumPath = Resolve-Path "$PSModuleRoot\Lib\Selenium.Support.*\lib\netstandard2.0" -ea SilentlyContinue)) {
        Write-Error "Couldn't find WebDriver.Support.dll"
    }
    # Load Selenium Support
    try {
        Add-Type -Path (Join-Path $SeleniumPath 'WebDriver.Support.dll') -ErrorAction Stop
    }
    catch {
        Write-Error "Couldn't load Selenium Support"
    }
}

# Import Assert function
if ($AssertPath = Resolve-Path "$PSModuleRoot\Function\Assert.psm1" -ea SilentlyContinue) {
    Import-Module $AssertPath -Force
}