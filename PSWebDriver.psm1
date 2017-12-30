# Load Classes
$ClassList = @(
    'PSWebDriver.ps1'
)
foreach ($class in $ClassList) {
    . $PSScriptRoot\Class\$class
}
