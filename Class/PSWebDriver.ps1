#Require -Version 5.0
using namespace OpenQA.Selenium

Enum ImageFormat{
    Png = 0
    Jpeg = 1
    Gif = 2
    Tiff = 3
    Bmp = 4
}

Enum SelectorType{
    None
    Id
    Name
    Tag
    ClassName
    Link
    XPath
    Css
}

class Selector {
    [string]$Expression
    [SelectorType]$Type = [SelectorType]::None

    Selector() {
    }

    Selector([string]$Expression) {
        $this.Expression = $Expression
    }

    Selector([string]$Expression, [SelectorType]$Type) {
        $this.Expression = $Expression
        $this.Type = $Type
    }

    [string]ToString() {
        return $this.Expression
    }

    static [Selector]Parse([string]$Expression) {
        $local:ret = switch -Regex ($Expression) {
            '^id=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Id) }
            '^name=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Name) }
            '^tag=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Tag) }
            '^className=(.+)' { [Selector]::new($Matches[1], [SelectorType]::ClassName) }
            '^link=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Link) }
            '^xpath=(.+)' { [Selector]::new($Matches[1], [SelectorType]::XPath) }
            '^//.+' { [Selector]::new($Matches[0], [SelectorType]::XPath) }
            '^css=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Css) }
            Default {[Selector]::new($Expression)}
        }
        return $ret
    }
}

class SpecialKeys {
    [hashtable]$KeyMap

    SpecialKeys() {
        $PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.KeyMap = (ConvertFrom-StringData (Get-Content (Join-Path $PSModuleRoot "Static\KEYMAP.txt") -raw))
    }

    [string]ConvertSeleniumKeys([string]$key) {
        if (!$this.KeyMap) {return ''}
        if ($this.KeyMap.ContainsKey($key)) {
            $tmp = $this.KeyMap.$key
            return [string](iex '[OpenQA.Selenium.keys]::($tmp)')
        }
        else {
            return ''
        }
    }
}

class PSWebDriver {
    #Properties
    [ValidateSet("Chrome", "Firefox", "Edge", "HeadlessChrome", "IE", "InternetExplorer")]
    [string]$BrowserName
    [SpecialKeys]$SpecialKeys
    $Driver

    Hidden [string] $StrictBrowserName
    Hidden [string] $DriverPackage
    Hidden [string] $PSModuleRoot


    # Constructor
    PSWebDriver([string]$Browser) {
        $this.PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.BrowserName = $Browser
        $this.StrictBrowserName = $this._ParseBrowserName($Browser)
        $this.DriverPackage = $this._ParseDriverPackage($Browser)
        $this.SpecialKeys = [SpecialKeys]::New()
        $this._LoadSelenium()
        $this._LoadWebDriver()
    }

    [void]Start() {
        if ($this.Driver) {
            $this.Quit()
        }

        $Options = $null
        # for Headless Chrome
        if ($this.BrowserName -eq 'HeadlessChrome') {
            $Options = New-Object Chrome.ChromeOptions
            $Options.AddArgument("--headless")
        }

        if($this.StrictBrowserName -eq 'IE'){
            $local:tmp = 'OpenQA.Selenium.IE.InternetExplorerDriver'
        }
        else{
            $local:tmp = [string]('OpenQA.Selenium.{0}.{0}{1}' -f $this.StrictBrowserName, "Driver")
        }
        #Start browser
        if (!$Options) {
            $this.Driver = New-Object $tmp
        }
        else {
            $this.Driver = New-Object $tmp($Options)
        }
    }

    [void]Start([Uri]$URL) {
        $this.Start()
        $this.Open($URL)
    }

    [void]Quit() {
        if ($this.Driver) {
            try {
                $this.Driver.Quit()
                Write-Verbose 'Browser terminated successfully.'
            }
            catch {
                throw 'Failed to terminate browser.'
            }
            finally {
                $this.Driver = $null
            }
        }
        else {
            $this._WarnBrowserNotStarted()
        }
    }

    # Aliase of Quit()
    [void]Close() {
        $this.Quit()
    }

    [void]Open([Uri]$URL) {
        if ($this.Driver) {
            $this.Driver.Navigate().GoToUrl($URL)
        }
        else {
            $this._WarnBrowserNotStarted()
        }
    }

    [string]GetTitle() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return [string]$this.Driver.Title
        }
    }

    [Object]FindElement([Selector]$Selector) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            $local:SelectorObj =
            switch ($Selector.Type) {
                'Id' { iex '[OpenQA.Selenium.By]::Id($Selector.ToString())' }
                'Name' { iex '[OpenQA.Selenium.By]::Name($Selector.ToString())'}
                'Tag' { iex '[OpenQA.Selenium.By]::TagName($Selector.ToString())' }
                'ClassName' { iex '[OpenQA.Selenium.By]::ClassName($Selector.ToString())'}
                'Link' { iex '[OpenQA.Selenium.By]::LinkText($Selector.ToString())' }
                'XPath' { iex '[OpenQA.Selenium.By]::XPath($Selector.ToString())' }
                'Css' { iex '[OpenQA.Selenium.By]::CssSelector($Selector.ToString())'}
                Default {
                    Write-Error 'Undefind selector type'
                    return $null
                }
            }
            if ($SelectorObj) {
                try {
                    return $this.Driver.FindElement($SelectorObj)

                }
                catch {
                    if ($_.Exception.InnerException.getType().FullName -eq "OpenQA.Selenium.NoSuchElementException") {
                        Write-Verbose ('No element found.')
                        return $null
                    }
                    else {
                        throw $_.Exception
                    }
                }
            }
            return $null
        }
    }

    [Object]FindElement([string]$SelectorExpression) {
        return $this.FindElement([Selector]::Parse($SelectorExpression))
    }

    [Object]FindElement([string]$SelectorExpression, [SelectorType]$Type) {
        return $this.FindElement([Selector]::New($SelectorExpression, $Type))
    }

    # [bool]IsElementPresent([Selector]$Selector) {
    #     return [bool]($this.FindElement($Selector))
    # }

    [bool]IsElementPresent([string]$SelectorExpression) {
        return [bool]($this.FindElement([Selector]::Parse($SelectorExpression)))
    }

    # [bool]IsElementPresent([string]$SelectorExpression, [SelectorType]$Type) {
    #     return [bool]($this.FindElement([Selector]::New($SelectorExpression, $Type)))
    # }

    [bool]IsAlertPresent() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $false
        }
        try {
            $this.Driver.SwitchTo().Alert()
            return $true
        }
        catch {
            if ($_.Exception.InnerException.getType().FullName -eq "OpenQA.Selenium.NoAlertPresentException") {
                Write-Verbose ('No Alert open.')
                return $false
            }
            else {
                throw $_.Exception
            }
        }
    }



    [void]SendKeys([string]$Target, [string]$Value) {
        $element = $this.FindElement($Target)
        if ($element) {
            if (($Value -match '\$\{(KEY_.+)\}') -and ($this.SpecialKeys)) {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys($Matches[1])
                $Value = ($Value -replace '\$\{KEY_.+\}', $Spec)
            }
            $element.SendKeys($Value)
        }
    }

    [void]ClearAndType([string]$Target, [string]$Value) {
        $element = $this.FindElement($Target)
        if ($element) {
            $element.Clear()
            if (($Value -match '\$\{(KEY_.+)\}') -and ($this.SpecialKeys)) {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys($Matches[1])
                $Value = ($Value -replace '\$\{KEY_.+\}', $Spec)
            }
            $element.SendKeys($Value)
        }
    }

    [void]Click([string]$Target) {
        $element = $this.FindElement($Target)
        if ($element) {
            $element.Click()
        }
    }

    [string]CloseAlertAndGetText([bool]$Accept) {
        [string]$AlertText = ''
        try {
            $Alert = $this.Driver.SwitchTo().Alert()
            $AlertText = [string]$Alert.Text
            if ($Accept) {
                $Alert.Accept()
            }
            else {
                $Alert.Dismiss()
            }
        }
        catch {}
        return $AlertText
    }

    [void]CloseAlert() {
        [void]$this.CloseAlertAndGetText($true)
    }


    [bool]WaitForElementPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.IsElementPresent($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForElementNotPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {!$this.IsElementPresent($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.FindElement($Target).GetAttribute('value')) -like $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.FindElement($Target).GetAttribute('value')) -notlike $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.FindElement($Target).Text) -like $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.FindElement($Target).Text) -notlike $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {($this.FindElement($Target).Displayed)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {!($this.FindElement($Target).Displayed)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.GetTitle()) -like $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {[string]($this.GetTitle()) -notlike $Value}
        $sb = $this._ConvertSeleniumPattern($sb, $Value)
        return $this._WaitForBase($sb, $Timeout)
    }

    [void]Pause([int]$WaitTimeInMilliSeconds) {
        [System.Threading.Thread]::Sleep($WaitTimeInMilliSeconds)
    }

    [string]ExecuteScript([string]$Script){
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        return [string]($this.Driver.ExecuteScript($Script))
    }

    [string]ExecuteScript([string]$Target, [string]$Script){
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        if($element = $this.FindElement($Target)){
            return [string]($this.Driver.ExecuteScript($Script, $element))
        }
        else{
            return $null
        }
    }

    [void]SaveScreenShot([string]$FileName, [ImageFormat]$ImageFormat) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            #TODO:To alternate [System.Drawing.Image] class
            iex '$ScrrenShot = [OpenQA.Selenium.Screenshot]$this.Driver.GetScreenShot()'
            iex '$ScrrenShot.SaveAsFile($FileName, [OpenQA.Selenium.ScreenshotImageFormat]$ImageFormat)'
        }
    }

    [void]SaveScreenShot([string]$FileName) {
        $this.SaveScreenShot($FileName, [ImageFormat]::Png)
    }


    # Private Method
    Hidden [void]_LoadSelenium() {
        $LibPath = Join-Path $this.PSModuleRoot '\Lib'
        if (!("OpenQA.Selenium.By" -as [type])) {
            if (!($SeleniumPath = Resolve-Path "$LibPath\Selenium.WebDriver.*\lib\net40" -ea SilentlyContinue)) {
                Write-Error "Couldn't find WebDriver.dll"
            }
            # Load Selenium
            try {
                Add-Type -Path (Join-Path $SeleniumPath 'WebDriver.dll') -ErrorAction Stop
            }
            catch {
                Write-Error "Couldn't load Selenium WebDriver"
            }
        }

        if (("OpenQA.Selenium.By" -as [type]) -and !("OpenQA.Selenium.Support.UI.SelectElement" -as [type])) {
            if (!($SeleniumPath = Resolve-Path "$LibPath\Selenium.Support.*\lib\net40" -ea SilentlyContinue)) {
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
    }

    Hidden [void]_LoadWebDriver() {
        $local:dir = [string]$this.DriverPackage
        if($this.StrictBrowserName -eq 'IE'){
            $local:exe = 'IEDriverServer.exe'
        }
        else{
            $local:exe = $dir + '.exe'
        }
        if (($this.StrictBrowserName -ne "Firefox") -and !(Get-Command $exe -ErrorAction SilentlyContinue)) {
            $DriverPath = Join-Path $this.PSModuleRoot "\Bin\$dir"
            if (Resolve-Path $DriverPath -ErrorAction SilentlyContinue) {
                if ($Env:Path.EndsWith(';')) {
                    $Env:Path += $DriverPath
                }
                else {
                    $Env:Path += (';' + $DriverPath)
                }
            }
            else {
                Write-Error "Couldn't find $exe"
            }
        }
    }

    Hidden [string]_ParseBrowserName([string]$BrowserName) {
        [string]$local:tmp = switch ($BrowserName) {
            "InternetExplorer" { "IE" }
            "HeadlessChrome" { "Chrome" }
            Default {$_}
        }
        return $tmp
    }

    Hidden [string]_ParseDriverPackage([string]$BrowserName) {
        [string]$local:tmp = switch ($BrowserName) {
            "Edge" { "MicrosoftWebDriver" }
            "IE" { "IEDriver" }
            "InternetExplorer" { "IEDriver" }
            "Chrome" { "ChromeDriver" }
            "HeadlessChrome" { "ChromeDriver" }
            default { "${_}Driver" }
        }
        return $tmp
    }

    Hidden [void]_WarnBrowserNotStarted([string]$Message) {
        Write-Warning $Message
    }

    Hidden [void]_WarnBrowserNotStarted() {
        $Message = 'Browser is not started.'
        $this._WarnBrowserNotStarted($Message)
    }


    Hidden [bool]_WaitForBase([ScriptBlock]$Expression, [int]$Timeout) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $false
        }

        $sec = 0;
        [bool]$ret = $false
        do {
            if ($sec -ge $Timeout) {
                $ret = $false
                Write-Error 'Timeout'
                break
            }
            try {
                if ($Expression.Invoke()) {
                    $ret = $true
                    break
                }
            }
            catch {$ret = $false}
            [System.Threading.Thread]::Sleep(1000)
            $sec++
        } while ($true)
        return $ret
    }

    # Selenium IDEの検索パターン文字列をPowershell比較演算子に置き換える
    Hidden [scriptblock]_ConvertSeleniumPattern([scriptblock]$ScriptBlock, [string]$Pattern) {
        # 入力する$ScriptBlockのフォーマット制限がきつい。比較対象の変数名は必ず$Value
        # 比較演算子は-likeもしくは-notlikeでなければ正しく動作しない
        $sbstr = $ScriptBlock.ToString()
        switch -Regex ($Pattern) {
            '^regexp:(.+)' {
                $sbstr = $sbstr -replace '-(|not)like', '-$1match'
                $sbstr = $sbstr.Replace('$Value', ("'{0}'" -f $Matches[1]))
            }
            '^glob:(.+)' {
                $sbstr = $sbstr.Replace('$Value', ("'{0}'" -f $Matches[1]))
            }
            '^exact:(.+)' {
                $sbstr = $sbstr.Replace('-like', '-eq')
                $sbstr = $sbstr.Replace('-notlike', '-ne')
                $sbstr = $sbstr.Replace('$Value', ("'{0}'" -f $Matches[1]))
            }
            Default {
                return $ScriptBlock
            }
        }
        return [scriptblock]::Create($sbstr)
    }

}

# #forDebug
# $obj = New-Object PSWebDriver('Chrome')
# $obj.Start('http://www.google.co.jp')
