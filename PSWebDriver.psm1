. $PSScriptRoot\Script\init.ps1

# Load Classes
$ClassList = @(
    'PSWebDriver.ps1'
)
foreach ($class in $ClassList) {
    . $PSScriptRoot\Class\$class
}

# Load Functions
$FunctionList = @(
    'Set-WebDriverEnvironment.ps1'
)
foreach ($function in $FunctionList) {
    . $PSScriptRoot\Function\$function
}

Export-ModuleMember -Function ('New-PSWebDriver', 'New-Selector', 'Set-WebDriverEnvironment')
