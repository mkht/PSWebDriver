#Require -Version 5.0

$PSModuleRoot = Split-Path $PSScriptRoot -Parent
# $LibPath = Join-Path $PSModuleRoot '\Lib'
Write-Debug ('$PSModuleRoot:{0}' -f $PSModuleRoot)

# Load Selenium WebDriver DLL
if (!("OpenQA.Selenium.By" -as [type])) {
    if (!($SeleniumPath = Resolve-Path "$PSModuleRoot\Lib\Selenium.WebDriver.*\lib\net40" -ea SilentlyContinue)) {
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
    if (!($SeleniumPath = Resolve-Path "$PSModuleRoot\Lib\Selenium.Support.*\lib\net40" -ea SilentlyContinue)) {
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

if (!('Microsoft.Edge.SeleniumTools.EdgeDriver' -as [type])) {
    if (!($SeleniumPath = Resolve-Path "$PSModuleRoot\Lib\microsoft.edge.seleniumtools.*\lib\netstandard2.0" -ea SilentlyContinue)) {
        throw "Couldn't find Microsoft.Edge.SeleniumTools.dll"
    }
    # Load Microsoft.Edge.SeleniumTools
    try {
        Add-Type -Path (Join-Path $SeleniumPath 'Microsoft.Edge.SeleniumTools.dll') -ErrorAction Stop
    }
    catch {
        throw "Couldn't load Microsoft.Edge.SeleniumTools"
    }
}

# Import Assert function
if ($AssertPath = Resolve-Path "$PSModuleRoot\Function\Assert.psm1" -ea SilentlyContinue) {
    Import-Module $AssertPath -Force
}