#Require -Version 5.0
using namespace OpenQA.Selenium

#region Enum:ImageFormat
Enum ImageFormat{
    Png = 0
    Jpeg = 1
    Gif = 2
    Tiff = 3
    Bmp = 4
}
#endregion

#region Enum:SelectorType
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
#endregion

#region Class:Selector
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
            '^/.+' { [Selector]::new($Matches[0], [SelectorType]::XPath) }
            '^css=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Css) }
            Default {[Selector]::new($Expression)}
        }
        return $ret
    }
}
#endregion

#region Class:SpecialKeys
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
            return ('${{{0}}}' -f $key)
        }
    }
}
#endregion

#region Class:PSWebDriver
class PSWebDriver {
    #region Public Properties
    $Driver
    #endregion

    #region Hidden properties
    [ValidateSet("Chrome", "Firefox", "Edge", "HeadlessChrome", "IE", "InternetExplorer")]
    Hidden [string] $BrowserName
    Hidden [string] $InstanceId
    Hidden [SpecialKeys] $SpecialKeys
    Hidden [string] $StrictBrowserName
    Hidden [string] $DriverPackage
    Hidden [string] $PSModuleRoot
    Hidden [int] $ImplicitWait = 0
    Hidden [int] $PageLoadTimeout = 30
    Hidden [System.Timers.Timer]$Timer
    Hidden [int]$RecordInterval = 5000
    #endregion

    #region Constructor:PSWebDriver
    PSWebDriver([string]$Browser) {
        $this.PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.InstanceId = [string]( -join ((1..4) | % {Get-Random -input ([char[]]((48..57) + (65..90) + (97..122)))})) #4-digits random id
        $this.BrowserName = $Browser
        $this.StrictBrowserName = $this._ParseBrowserName($Browser)
        $this.DriverPackage = $this._ParseDriverPackage($Browser)
        $this.SpecialKeys = [SpecialKeys]::New()
        $this._LoadSelenium()
        $this._LoadWebDriver()
    }
    #endregion

    #region Method:SetImplicitWait()
    [void]SetImplicitWait([int]$TimeoutInSeconds) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            [int]$local:tmp = $this.ImplicitWait
            try {
                if ($TimeoutInSeconds -lt 0) {
                    $TimeSpan = [System.Threading.Timeout]::InfiniteTimeSpan
                }
                else {
                    $TimeSpan = New-TimeSpan -Seconds $TimeoutInSeconds -ea Stop
                }
                $this.Driver.Manage().Timeouts().ImplicitWait = $TimeSpan
                $this.ImplicitWait = $TimeoutInSeconds
            }
            catch {
                $this.ImplicitWait = $tmp
            }
        }
    }
    #endregion

    #region Method:Get/SetWindowSize()
    [System.Drawing.Size]GetWindowSize() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return $this.Driver.Manage().Window.Size
        }
    }

    [void]SetWindowSize([System.Drawing.Size]$Size) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $this.Driver.Manage().Window.Size = $Size
        }
    }

    [void]SetWindowSize([int]$Width, [int]$Height) {
        $this.SetWindowSize([System.Drawing.Size]::New($Width, $Height))
    }
    #endregion

    #region Method:Start()
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

        if ($this.StrictBrowserName -eq 'IE') {
            $local:tmp = 'OpenQA.Selenium.IE.InternetExplorerDriver'
        }
        else {
            $local:tmp = [string]('OpenQA.Selenium.{0}.{0}{1}' -f $this.StrictBrowserName, "Driver")
        }
        #Start browser
        if (!$Options) {
            $this.Driver = New-Object $tmp
        }
        else {
            $this.Driver = New-Object $tmp($Options)
        }
        #Set default implicit wait
        if ($this.Driver) {$this.SetImplicitWait($this.ImplicitWait)}
    }

    [void]Start([Uri]$URL) {
        $this.Start()
        $this.Open($URL)
    }
    #endregion

    #region Method:Quit()
    [void]Quit() {
        # Stop animation recorder if running
        if ($this.Timer) {
            try {
                $this._DisposeRecorder()
            }
            catch {}
        }

        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
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
    }
    #endregion

    #region Method:Close()
    [void]Close() {
        if (!$this.Driver) {
            $this.Quit()
        }
        elseif ($this.Driver.WindowHandles.Count -le 1) {
            $this.Quit()
        }
        else {
            $this.Driver.Close()
            if (!$this.Driver.WindowHandles) {
                $this.Quit()
            }
        }
    }
    #endregion

    #region Method:Open()
    [void]Open([Uri]$URL) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $this.Driver.Navigate().GoToUrl($URL)
        }
    }
    #endregion

    #region Method:GetTitle()
    [string]GetTitle() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return [string]$this.Driver.Title
        }
    }
    #endregion

    #region Method:Get/SetLocation()
    [string]GetLocation() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            return [string]$this.Driver.Url
        }
    }

    # Aliase of Open()
    [void]SetLocation([Uri]$URL) {
        $this.Open($URL)
    }
    #endregion

    [string]GetAttribute([string]$Target, [string]$Attribute) {
        return [string]($this.FindElement($Target).GetAttribute($Attribute))
    }

    [string]GetText([string]$Target) {
        return [string]($this.FindElement($Target).Text)
    }

    [bool]IsVisible([string]$Target) {
        return [bool]($this.FindElement($Target).Displayed)
    }

    #region Method:FindElement()
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
                    throw 'Undefind selector type'
                    return $null
                }
            }
            if ($SelectorObj) {
                $this.WaitForPageToLoad($this.PageLoadTimeout)
                return $this.Driver.FindElement($SelectorObj)
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
    #endregion

    #region Method:IsElementPresent()
    [bool]IsElementPresent([string]$SelectorExpression) {
        [int]$tmpWait = $this.ImplicitWait
        try {
            # Set implicit wait to 0 sec temporally.
            if ($this.Driver) {$this.SetImplicitWait(0)}
            return [bool]($this.FindElement([Selector]::Parse($SelectorExpression)))
        }
        catch {
            return $false
        }
        finally {
            # Reset implicit wait
            if ($this.Driver) {$this.SetImplicitWait($tmpWait)}
        }
    }
    #endregion

    #region Method:Sendkeys() & ClearAndType()
    [void]SendKeys([string]$Target, [string]$Value) {
        if ($element = $this.FindElement($Target)) {
            $Ret = $Value
            $local:regex = [regex]'\$\{KEY_.+?\}'
            $regex.Matches($Value) | % {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys(($_.Value).SubString(2, ($_.Value.length - 3)))
                $Ret = $Ret.Replace($_.Value, $Spec)
            }
            $element.SendKeys($Ret)
        }
    }

    [void]ClearAndType([string]$Target, [string]$Value) {
        if ($element = $this.FindElement($Target)) {
            $element.Clear()
            if (($Value -match '\$\{(KEY_.+)\}') -and ($this.SpecialKeys)) {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys($Matches[1])
                $Value = ($Value -replace '\$\{KEY_.+\}', $Spec)
            }
            $element.SendKeys($Value)
        }
    }
    #endregion

    #region Method:Click()
    [void]Click([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $element.Click()
        }
    }
    #endregion

    #region Method:Select()
    [void]Select([string]$Target, [string]$Value) {
        if ($SelectElement = $this._GetSelectElement($Target)) {
                #TODO: Implement SelectByIndex
                #TODO: Implement SelectByValue
                $SelectElement.SelectByText($Value)
            }
        }
    #endregion

    #region Switch window
    [void]SelectWindow([string]$Title) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $FoundFlag = $false
            $Pattern = $this._PerseSeleniumPattern($Title)
            $CurrentWindow = $this.Driver.CurrentWindowHandle
            $AllWindow = $this.Driver.WindowHandles
            #Enumerate all windows
            :SWLOOP foreach ($window in $AllWindow) {
                if ($window -eq $CurrentWindow) {
                    $title = $this.Driver.Title
                }
                else {
                    $title = $this.Driver.SwitchTo().Window($window).Title
                }

                switch ($Pattern.Matcher) {
                    'Like' {
                        if ($title -like $Pattern.Pattern) {
                            $FoundFlag = $true
                            break SWLOOP
                        }
                    }
                    'RegExp' {
                        if ($title -match $Pattern.Pattern) {
                            $FoundFlag = $true
                            break SWLOOP
                        }
                    }
                    'Equal' {
                        if ($title -eq $Pattern.Pattern) {
                            $FoundFlag = $true
                            break SWLOOP
                        }
                    }
                }
            }

            if (!$FoundFlag) {
                if ($this.Driver.CurrentWindowHandle -ne $CurrentWindow) {
                    #Retrun current window
                    $this.Driver.SwitchTo().Window($CurrentWindow)
                }
                #throw NoSuchWindowException
                $this.Driver.SwitchTo().Window([System.Guid]::NewGuid().ToString)
            }
        }
    }
    #endregion

    #region Method:SelectFrame()
    [void]SelectFrame([string]$Name) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            $this.Driver.SwitchTo().DefaultContent()
            $this.Driver.SwitchTo().Frame($Name)
        }
    }
    #endregion

    #region HTTP Status Code (Invoke-WebRequest)
    [int]GetHttpStatusCode([Uri]$URL) {
        try {
            $response = Invoke-WebRequest -Uri $URL -UseBasicParsing -ErrorAction Stop
        }
        catch {
            $response = $_.Exception.Response
            if (!$response) {
                throw $_.Exception
            }
        }

        if ($response.StatusCode -as [int]) {
            return [int]$response.StatusCode
        }
        else {
            throw [System.Exception]::new('Unexpected Exception')
        }
    }

    # Assertion
    [void]AssertHttpStatusCode([Uri]$URL, [int]$Value) {
        $this.GetHttpStatusCode($URL) | Assert -Expected $Value
    }

    [void]AssertNotHttpStatusCode([Uri]$URL, [int]$Value) {
        $this.GetHttpStatusCode($URL) | Assert -Not -Expected $Value
    }
    #endregion

    #region Alert handling
    #region Method:IsAlertPresent()
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
    #endregion

    #region Method:CloseAlertAndGetText()
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
    #endregion

    #region Method:CloseAlertAndGetText()
    [void]CloseAlert() {
        [void]$this.CloseAlertAndGetText($true)
    }
    #endregion
    #endregion Alert handling

    #region WaitFor*
    #region Method:WaitForPageToLoad()
    [bool]WaitForPageToLoad([int]$Timeout) {
        $sb = [ScriptBlock] {
            if (!($this.ExecuteScript('return document.readyState;') -eq 'complete')) {
                throw [System.Exception]::new()
            }
        }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForElementPresent()
    [bool]WaitForElementPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertElementPresent($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForElementNotPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertElementNotPresent($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForValue()
    [bool]WaitForValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertValue($Target, $Value)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertNotValue($Target, $Value)}
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForText()
    [bool]WaitForText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertText($Target, $Value)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertNotText($Target, $Value)}
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForVisible()
    [bool]WaitForVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertVisible($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertNotVisible($Target)}
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForTitle()
    [bool]WaitForTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertTitle($Value)}
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] {$this.AssertNotTitle($Value)}
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion
    #endregion WatFor*

    #region Method:Pause()
    [void]Pause([int]$WaitTimeInMilliSeconds) {
        [System.Threading.Thread]::Sleep($WaitTimeInMilliSeconds)
    }
    #endregion

    #region Method:ExecuteScript()
    [string]ExecuteScript([string]$Script) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        return [string]($this.Driver.ExecuteScript($Script))
    }

    [string]ExecuteScript([string]$Target, [string]$Script) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        if ($element = $this.FindElement($Target)) {
            return [string]($this.Driver.ExecuteScript($Script, $element))
        }
        else {
            return $null
        }
    }
    #endregion

    #region Method:SaveScreenShot()
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
    #endregion

    #region Hidden Method
    Hidden [void]_LoadSelenium() {
        $LibPath = Join-Path $this.PSModuleRoot '\Lib'
        if (!("OpenQA.Selenium.By" -as [type])) {
            if (!($SeleniumPath = Resolve-Path "$LibPath\Selenium.WebDriver.*\lib\net40" -ea SilentlyContinue)) {
                throw "Couldn't find WebDriver.dll"
            }
            # Load Selenium
            try {
                Add-Type -Path (Join-Path $SeleniumPath 'WebDriver.dll') -ErrorAction Stop
            }
            catch {
                throw "Couldn't load Selenium WebDriver"
            }
        }

        if (("OpenQA.Selenium.By" -as [type]) -and !("OpenQA.Selenium.Support.UI.SelectElement" -as [type])) {
            if (!($SeleniumPath = Resolve-Path "$LibPath\Selenium.Support.*\lib\net40" -ea SilentlyContinue)) {
                throw "Couldn't find WebDriver.Support.dll"
            }
            # Load Selenium Support
            try {
                Add-Type -Path (Join-Path $SeleniumPath 'WebDriver.Support.dll') -ErrorAction Stop
            }
            catch {
                throw "Couldn't load Selenium Support"
            }
        }
    }

    Hidden [void]_LoadWebDriver() {
        $local:dir = [string]$this.DriverPackage
        if ($this.StrictBrowserName -eq 'IE') {
            $local:exe = 'IEDriverServer.exe'
        }
        else {
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
                throw "Couldn't find $exe"
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
        if (($Timeout -lt 0) -or ($Timeout -gt 3600)) {
            throw [System.ArgumentOutOfRangeException]::New()
        }

        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $false
        }

        [int]$tmpWait = $this.ImplicitWait
        # Set implicit wait to 0 sec temporally.
        if ($this.Driver) {$this.SetImplicitWait(0)}

        $sec = 0;
        [bool]$ret = $false
        do {
            try {
                $Expression.Invoke()
                $ret = $true
                break
            }
            catch {
                $ret = $false
            }
            if ($sec -ge $Timeout) {
                $ret = $false
                throw 'Timeout'
                break
            }
            [System.Threading.Thread]::Sleep(1000)
            $sec++
        } while ($true)

        if ($this.Driver) {$this.SetImplicitWait($tmpWait)}
        return $ret
    }

    Hidden [HashTable]_PerseSeleniumPattern([string]$Pattern) {
        $local:ret = [HashTable]@{
            Matcher = ''
            Pattern = ''
        }

        switch -Regex ($Pattern) {
            '^regexp:(.+)' {
                $ret.Matcher = 'RegExp'
                $ret.Pattern = $Matches[1]
            }
            '^glob:(.+)' {
                $ret.Matcher = 'Like'
                $ret.Pattern = $Matches[1]
            }
            '^exact:(.+)' {
                $ret.Matcher = 'Equal'
                $ret.Pattern = $Matches[1]
            }
            Default {
                $ret.Matcher = 'Like'
                $ret.Pattern = $Pattern
            }
        }
        return $ret
    }

    Hidden [Object]_GetSelectElement([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $SelectElement = $null
            iex '$SelectElement = New-Object "OpenQA.Selenium.Support.UI.SelectElement" $element' -ea Stop
            return $SelectElement
        }
        else {
            return $null
        }
    }
    #endregion Hidden Method

    #region [Experimental] Animated GIF Recorder
    Hidden [void]_InitRecorder() {
        $this._DisposeRecorder()

        # Load AnimatedGifWrapper class
        $LibPath = Join-Path $this.PSModuleRoot '\Lib'
        if (!("AnimatedGifWrapper" -as [type])) {
            $local:CSharpCode = gc (Join-Path $LibPath '\AnimatedGifWrapper\AnimatedGifWrapper.cs') -Raw -ea SilentlyContinue
            $local:DllPath = Resolve-Path "$LibPath\AnimatedGif.*\lib\AnimatedGif.dll" -ea SilentlyContinue
            try {
                if ($CSharpCode) {
                    [void][Reflection.assembly]::LoadFrom($DllPath.ToString())
                    Add-Type -Language CSharp -TypeDefinition $CSharpCode -ReferencedAssemblies ('System.Drawing', 'System.Windows.Forms', $DllPath.ToString()) -ea Stop
                }
            }
            catch {
                throw "Couldn't load AnimatedGifWrapper Class"
            }
        }

        #Create background timer job
        if ("AnimatedGifWrapper" -as [type]) {
            try {
                #Jobから実行できるようにGlobalスコープで作成
                Set-Variable -Name ('Recorder' + $this.InstanceId) -Value (New-Object AnimatedGifWrapper) -Scope Global #他のインスタンスと被らないよう変数名にIDを付ける
                Set-Variable -Name ('Counter' + $this.InstanceId) -Value 0 -Scope Global

                if ($this.Driver) {
                    #Jobから実行できるようにWebDriverをGlobalスコープにする
                    Set-Variable -Name ('WebDriver' + $this.InstanceId) -Value $this.Driver -Scope Global
                }

                $this.Timer = New-Object System.Timers.Timer
                $this.Timer.Interval = $this.RecordInterval #Interval ms

                #forDebug
                # $global:Logging = @()
                $Action = {
                    $Id = [string]$Event.MessageData
                    $localCounter = [int](Get-Variable -Name ('Counter' + $Id) -ea SilentlyContinue).Value
                    #停止し忘れやメモリ食いつぶしを防ぐために最大記録回数を制限する
                    if ($localCounter -gt 1200) {
                        #500ms*1200=10分想定
                        #Stop
                    }
                    else {
                        Set-Variable -Name ('Counter' + $Id) -Value ($localCounter + 1)
                        # $global:Logging += ("ID:{0}" -f $Id) #forDebug
                        if ((Get-Variable ('WebDriver' + $Id) -ea SilentlyContinue) -and (Get-Variable ('Recorder' + $Id) -ea SilentlyContinue)) {
                            try {
                                (Get-Variable ('Recorder' + $Id)).Value.AddFrame([byte[]]((Get-Variable ('WebDriver' + $Id)).Value.GetScreenShot().AsByteArray))
                                #forDebug
                                # $global:Logging += ("ScreenShot ok:{0}" -f [datetime]::Now.ToString())
                            }
                            catch {
                                #forDebug
                                # $global:Logging += ("ScreenShot error:{0}" -f $_Exception.Message)
                            }
                        }
                    }
                }

                Register-ObjectEvent -InputObject $this.Timer -EventName Elapsed -SourceIdentifier $this.InstanceId -Action $Action -MessageData $this.InstanceId> $null
            }
            catch {
                throw 'Failed Initilize GIF Recorder'
                $this._DisposeRecorder()    #Dispose recorder when error occured.
            }
        }
    }

    Hidden [void]_DisposeRecorder() {
        if ($this.Timer) {
            try {
                $this.Timer.Close() #Stop timer
            }
            catch {}
            finally {
                $this.Timer = $null
            }
        }

        # Unregister Event subscriber
        Get-EventSubscriber -SourceIdentifier $this.InstanceId -ea SilentlyContinue | Unregister-Event

        # Dispose Recorder
        if (Get-Variable ('Recorder' + $this.InstanceId) -ea SilentlyContinue) {
            try {
                (Get-Variable ('Recorder' + $this.InstanceId)).Value.Dispose()
            }
            catch {}
            finally {
                Remove-Variable -Name ('Recorder' + $this.InstanceId) -Scope Global
            }
        }

        # Release global scope WebDriver
        if (Get-Variable ('WebDriver' + $this.InstanceId) -ea SilentlyContinue) {
            Remove-Variable -Name ('WebDriver' + $this.InstanceId) -Scope Global
        }
    }

    [void]StartAnimationRecord([int]$Interval) {
        # Minimum interval is 500ms
        if ($Interval -lt 500) {
            throw [System.ArgumentOutOfRangeException]::new("You can't specify Interval less than 500ms")
            return
        }

        #Chack is recoder already started
        if (Get-EventSubscriber -SourceIdentifier $this.InstanceId -ea SilentlyContinue) {
            throw 'Recorder has already started !'
            return
        }

        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return
        }

        $this.RecordInterval = $Interval
        $this._InitRecorder()
        if ($this.Timer) {
            #First capture
            (Get-Variable ('Recorder' + $this.InstanceId)).Value.AddFrame([byte[]]((Get-Variable ('WebDriver' + $this.InstanceId)).Value.GetScreenShot().AsByteArray))
            #Start record
            $this.Timer.start()
        }
    }

    [void]StartAnimationRecord() {
        $this.StartAnimationRecord(1000)  #default interval 1 sec
    }

    [void]StopAnimationRecord([string]$FileName) {
        try {
            if ($this.Timer) {
                $this.Timer.Stop()  #Stop recorder
            }

            # Save Animated GIF file
            if ($FileName) {
                if (Get-Variable ('Recorder' + $this.InstanceId) -ea SilentlyContinue) {
                    (Get-Variable ('Recorder' + $this.InstanceId)).Value.Save($FileName, $this.RecordInterval)
                }
            }

            Unregister-Event $this.InstanceId -ea SilentlyContinue
        }
        catch {}
        finally {
            $this._DisposeRecorder()
        }
    }

    #Only stop record. Don't save file
    [void]StopAnimationRecord() {
        $this.StopAnimationRecord($null)
    }
    #endregion

    #region Assertion
    [void]AssertElementPresent([string]$Selector) {
        $this.IsElementPresent($Selector) | Assert -Expected $true
    }

    [void]AssertElementNotPresent([string]$Selector) {
        $this.IsElementPresent($Selector) | Assert -Expected $false
    }

    [void]AssertAlertPresent() {
        $this.IsAlertPresent() | Assert -Expected $true
    }

    [void]AssertAlertNotPresent() {
        $this.IsAlertPresent() | Assert -Expected $false
    }

    [void]AssertTitle([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetTitle() | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotTitle([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetTitle() | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertLocation([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetLocation() | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotLocation([string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetLocation() | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertAttribute([string]$Target, [string]$Attribute, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetAttribute($Target, $Attribute) | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotAttribute([string]$Target, [string]$Attribute, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetAttribute($Target, $Attribute) | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertValue([string]$Target, [string]$Value) {
        $this.AssertAttribute($Target, 'value', $Value)
    }

    [void]AssertNotValue([string]$Target, [string]$Value) {
        $this.AssertNotAttribute($Target, 'value', $Value)
    }

    [void]AssertText([string]$Target, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetText($Target) | Assert -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertNotText([string]$Target, [string]$Value) {
        $Pattern = $this._PerseSeleniumPattern($Value)
        $this.GetText($Target) | Assert -Not -Expected $Pattern.Pattern -Matcher $Pattern.Matcher
    }

    [void]AssertVisible([string]$Target) {
        $this.IsVisible($Target) | Assert -Expected $true
    }

    [void]AssertNotVisible([string]$Target) {
        $this.IsVisible($Target) | Assert -Expected $false
    }
    #endregion
}
#endregion

# #forDebug
# $obj = New-Object PSWebDriver('Chrome')
# $obj.Start('http://www.google.co.jp')
