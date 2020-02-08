. $PSScriptRoot\Script\init.ps1

# Load Classes
$ClassList = @(
    'PSWebDriver.ps1'
)
foreach ($class in $ClassList) {
    . $PSScriptRoot\Class\$class
}

Export-ModuleMember -Function ('New-PSWebDriver', 'New-Selector')
