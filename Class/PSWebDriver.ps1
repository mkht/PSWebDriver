﻿#Requires -Version 5
using namespace OpenQA.Selenium
using namespace Microsoft.Edge.SeleniumTools

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
    [Object]$By

    Selector() {
    }

    Selector([string]$Expression) {
        $this.Expression = $Expression
    }

    Selector([string]$Expression, [SelectorType]$Type) {
        $this.Expression = $Expression
        $this.Type = $Type
        $this.By = [Selector]::GetSeleniumBy($Expression, $Type)
    }

    [string]ToString() {
        return $this.Expression
    }

    static [Selector]Parse([string]$Expression) {
        $local:ret = switch -Regex ($Expression) {
            '^id=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Id); break }
            '^name=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Name); break }
            '^tag=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Tag); break }
            '^className=(.+)' { [Selector]::new($Matches[1], [SelectorType]::ClassName); break }
            '^link=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Link); break }
            '^xpath=(.+)' { [Selector]::new($Matches[1], [SelectorType]::XPath); break }
            '^/.+' { [Selector]::new($Matches[0], [SelectorType]::XPath); break }
            '^css=(.+)' { [Selector]::new($Matches[1], [SelectorType]::Css); break }
            Default { [Selector]::new($Expression) }
        }
        return $ret
    }

    static Hidden [Object]GetSeleniumBy([string]$Expression, [SelectorType]$Type) {
        $local:SelectorObj =
        switch ($Type) {
            'Id' { Invoke-Expression '[By]::Id($Expression)'; break }
            'Name' { Invoke-Expression '[By]::Name($Expression)'; break }
            'Tag' { Invoke-Expression '[By]::TagName($Expression)'; break }
            'ClassName' { Invoke-Expression '[By]::ClassName($Expression)'; break }
            'Link' { Invoke-Expression '[By]::LinkText($Expression)'; break }
            'XPath' { Invoke-Expression '[By]::XPath($Expression)'; break }
            'Css' { Invoke-Expression '[By]::CssSelector($Expression)'; break }
            Default {
                throw 'Undefined selector type'
            }
        }
        return $SelectorObj
    }
}
#endregion

#region Class:SpecialKeys
class SpecialKeys {
    [hashtable]$KeyMap

    SpecialKeys() {
        $PSModuleRoot = Split-Path $PSScriptRoot -Parent
        $this.KeyMap = (ConvertFrom-StringData (Get-Content (Join-Path $PSModuleRoot 'Static\KEYMAP.txt') -raw))
    }

    [string]ConvertSeleniumKeys([string]$key) {
        if (!$this.KeyMap) { return '' }
        if ($this.KeyMap.ContainsKey($key)) {
            $tmp = $this.KeyMap.$key
            return [string](Invoke-Expression '[keys]::($tmp)')
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
    $Actions
    [ValidateSet('Chrome', 'Firefox', 'Edge', 'EdgeChromium', 'HeadlessChrome', 'HeadlessFirefox', 'IE', 'InternetExplorer')]
    [string] $BrowserName
    [DriverOptions] $BrowserOptions
    [DriverService] $DriverService
    #endregion

    #region Hidden properties
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
        $this.InstanceId = [string]( -join ((1..4) | ForEach-Object { Get-Random -input ([char[]]((48..57) + (65..90) + (97..122))) })) #4-digits random id
        $this.BrowserName = $Browser
        $this.StrictBrowserName = $this._ParseBrowserName($Browser)
        $this.DriverPackage = $this._ParseDriverPackage($Browser)
        $this.SpecialKeys = [SpecialKeys]::New()

        # Add accessor properties
        $this | Add-Member ScriptProperty 'Location' {
            # getter
            if (!$this.Driver) {
                $null
            }
            else {
                [string]$this.Driver.Url
            }
        } {
            # setter
            param ( $arg )
            $this.Open($arg)
        }

        $this._LoadSelenium()
        $this._LoadWebDriver()
        $this.BrowserOptions = $this._NewDriverOptions($Browser)
        $this.DriverService = $this._NewDriverService($Browser)
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

        if ($this.StrictBrowserName -eq 'IE') {
            $local:tmp = 'IE.InternetExplorerDriver'
        }
        elseif ($this.StrictBrowserName -eq 'EdgeChromium') {
            $local:tmp = 'EdgeDriver'
        }
        else {
            $local:tmp = [string]('{0}.{0}{1}' -f $this.StrictBrowserName, 'Driver')
        }

        #Start browser
        if ($null -eq $this.BrowserOptions) {
            if ($null -eq $this.DriverService) {
                $this.Driver = New-Object $tmp
            }
            else {
                $this.Driver = New-Object $tmp($this.DriverService)
            }
        }
        else {
            if ($null -eq $this.DriverService) {
                $this.Driver = New-Object $tmp($this.BrowserOptions)
            }
            else {
                $this.Driver = New-Object $tmp($this.DriverService, $this.BrowserOptions, [TimeSpan]::FromMinutes(3))
            }
        }

        #Set default implicit wait
        if ($this.Driver) { $this.SetImplicitWait($this.ImplicitWait) }

        #Create Action instance
        if ($this.Driver) { $this.Actions = Invoke-Expression '[Interactions.Actions]::New($this.Driver)' }
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
            catch { }
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
            $this.CloseAlert($false)
            $this.Driver.Navigate().GoToUrl($URL)
        }
    }
    #endregion

    #region Method:GetBrowserInfo()
    [HashTable]GetBrowserInfo() {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            $AllInfo = $this.Driver.Capabilities.ToDictionary()
            $ReturnInfo = @{
                browserName    = $AllInfo.browserName
                browserVersion = $AllInfo.browserVersion
                platformName   = $AllInfo.platformName
            }
            return $ReturnInfo
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
            return [string]$this.Location
        }
    }

    # Aliases of Open()
    [void]SetLocation([Uri]$URL) {
        $this.Location = $URL
    }
    #endregion

    [string]GetAttribute([string]$Target, [string]$Attribute) {
        return [string]($this.FindElement($Target).GetAttribute($Attribute))
    }

    [string]GetText([string]$Target) {
        return [string]($this.FindElement($Target).Text)
    }

    [string]GetSelectedLabel([string]$Target) {
        return [string]($this._GetSelectElement($Target).SelectedOption.Text)
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
            if ($Selector.By) {
                $this.WaitForPageToLoad($this.PageLoadTimeout)
                return $this.Driver.FindElement($Selector.By)
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

    #region Method:FindElements()
    [Object[]]FindElements([Selector]$Selector) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        else {
            if ($Selector.By) {
                $this.WaitForPageToLoad($this.PageLoadTimeout)
                return $this.Driver.FindElements($Selector.By)
            }
            return $null
        }
    }

    [Object[]]FindElements([string]$SelectorExpression) {
        return $this.FindElements([Selector]::Parse($SelectorExpression))
    }

    [Object[]]FindElements([string]$SelectorExpression, [SelectorType]$Type) {
        return $this.FindElements([Selector]::New($SelectorExpression, $Type))
    }
    #endregion

    #region Method:IsElementPresent()
    [bool]IsElementPresent([string]$SelectorExpression) {
        [int]$tmpWait = $this.ImplicitWait
        try {
            # Set implicit wait to 0 sec temporally.
            if ($this.Driver) { $this.SetImplicitWait(0) }
            return [bool]($this.FindElement([Selector]::Parse($SelectorExpression)))
        }
        catch {
            return $false
        }
        finally {
            # Reset implicit wait
            if ($this.Driver) { $this.SetImplicitWait($tmpWait) }
        }
    }
    #endregion

    #region Method:Sendkeys() & ClearAndType()
    Hidden [void]_InnerSendKeys([string]$Target, [string]$Value, [bool]$BeforeClear) {
        if ($element = $this.FindElement($Target)) {
            if ($BeforeClear) {
                $element.Clear()
            }
            $private:Ret = $Value
            $private:regex = [regex]'\$\{KEY_.+?\}'
            $regex.Matches($Value) | ForEach-Object {
                $Spec = $this.SpecialKeys.ConvertSeleniumKeys(($_.Value).SubString(2, ($_.Value.length - 3)))
                $Ret = $Ret.Replace($_.Value, $Spec)
            }
            $element.SendKeys($Ret)
        }
    }

    [void]SendKeys([string]$Target, [string]$Value) {
        $this._InnerSendKeys($Target, $Value, $false)
    }

    [void]ClearAndType([string]$Target, [string]$Value) {
        $this._InnerSendKeys($Target, $Value, $true)
    }
    #endregion

    #region Method:Click()
    [void]Click([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $element.Click()
        }
    }

    # Invoke JavaScript click() method
    # Sometimes Selenium Click is not working with IE11 in Windows 10.
    # https://github.com/SeleniumHQ/selenium/issues/4292
    [void]JavaScriptClick([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $this.Driver.ExecuteScript('arguments[0].click();', $element)
        }
    }
    #endregion

    #region Method:DoubleClick()
    [void]DoubleClick([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $this.Actions.DoubleClick($element).build().perform()
        }
    }
    #endregion

    #region Method:RightClick()
    [void]RightClick([string]$Target) {
        if ($element = $this.FindElement($Target)) {
            $this.Actions.ContextClick($element).build().perform()
        }
    }
    #endregion

    #region Method:ClickAt()
    [void]ClickAt([int]$X, [int]$Y) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return
        }
        else {
            $js = @"
            var e = document.createEvent('MouseEvents');
            e.initMouseEvent('click',true,true,window,1,0,0,$X,$Y,false,false,false,false,0,null,);
            document.body.dispatchEvent(e);
"@
            $this.ExecuteScript($js)
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
            $IsWindowFound = $false
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
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                    'RegExp' {
                        if ($title -match $Pattern.Pattern) {
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                    'Equal' {
                        if ($title -eq $Pattern.Pattern) {
                            $IsWindowFound = $true
                            break SWLOOP
                        }
                    }
                }
            }

            if (!$IsWindowFound) {
                if ($this.Driver.CurrentWindowHandle -ne $CurrentWindow) {
                    #Return current window
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
            if ($_.Exception.InnerException.getType().FullName -eq 'OpenQA.Selenium.NoAlertPresentException') {
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
        catch { }

        #hack for Edge
        #In case of Edge, executing the command immediately after closing an alert often fails, so wait a little.
        if ($this.StrictBrowserName -eq 'Edge') {
            [System.Threading.Thread]::Sleep(500)
        }

        return $AlertText
    }
    #endregion

    #region Method:CloseAlert()
    [void]CloseAlert([bool]$Accept) {
        [void]$this.CloseAlertAndGetText($Accept)
    }

    [void]CloseAlert() {
        [void]$this.CloseAlert($true)
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
        $sb = [ScriptBlock] { $this.AssertElementPresent($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForElementNotPresent([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertElementNotPresent($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForValue()
    [bool]WaitForValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertValue($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotValue([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotValue($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForText()
    [bool]WaitForText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertText($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotText([string]$Target, [string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotText($Target, $Value) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForVisible()
    [bool]WaitForVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertVisible($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotVisible([string]$Target, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotVisible($Target) }
        return $this._WaitForBase($sb, $Timeout)
    }
    #endregion

    #region Method:WaitForTitle()
    [bool]WaitForTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertTitle($Value) }
        return $this._WaitForBase($sb, $Timeout)
    }

    [bool]WaitForNotTitle([string]$Value, [int]$Timeout) {
        $sb = [ScriptBlock] { $this.AssertNotTitle($Value) }
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
    [Object]ExecuteScript([string]$Script) {
        return $this.ExecuteScript($Script, $null)
    }

    [Object]ExecuteScript([string]$Script, [Object[]]$Arguments) {
        $Object = $null
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
            return $null
        }
        if ($Arguments) {
            $Object = $this.Driver.ExecuteScript($Script, $Arguments)
        }
        else {
            $Object = $this.Driver.ExecuteScript($Script)
        }

        # Selenium ExecuteScript() seems that return [Array] as [ReadOnlyCollection<T>].
        # For ease of handling, convert to [System.Array]
        if ($Object -and ($Object.GetType().Name -match 'ReadOnlyCollection')) {
            return [Object[]]$Object
        }
        else {
            return $Object
        }
    }
    #endregion

    #region Method:SaveScreenShot()
    [void]SaveScreenShot([string]$FileName, [ImageFormat]$ImageFormat) {
        if (!$this.Driver) {
            $this._WarnBrowserNotStarted()
        }
        else {
            # Path normalization
            if ($global:PSVersionTable.PSVersion -ge 6.1) {
                $NormalizedFilePath = [System.IO.Path]::GetFullPath($FileName, $PWD)  #This method can be used only on .NET Core 2.1+
            }
            else {
                if (([uri]$FileName).IsAbsoluteUri) {
                    $NormalizedFilePath = [System.IO.Path]::GetFullPath($FileName)
                }
                else {
                    $NormalizedFilePath = [System.IO.Path]::GetFullPath((Join-path $PWD $FileName))
                }
            }
            
            $SaveFolder = Split-Path $NormalizedFilePath -Parent
            if ($null -eq $SaveFolder) {
                throw [System.ArgumentNullException]::new()
            }
            elseif (! (Test-Path -LiteralPath $SaveFolder -PathType Container)) {
                New-Item $SaveFolder -ItemType Directory
            }
            #TODO:To alternate [System.Drawing.Image] class
            Invoke-Expression '$ScreenShot = [Screenshot]$this.Driver.GetScreenShot()'
            Invoke-Expression '$ScreenShot.SaveAsFile($NormalizedFilePath, [ScreenshotImageFormat]$ImageFormat)'
        }
    }

    [void]SaveScreenShot([string]$FileName) {
        [ImageFormat]$Format =
        switch ([System.IO.Path]::GetExtension($FileName)) {
            '.jpg' { [ImageFormat]::Jpeg; break }
            '.jpeg' { [ImageFormat]::Jpeg; break }
            '.bmp' { [ImageFormat]::Bmp; break }
            '.gif' { [ImageFormat]::Gif; break }
            '.tif' { [ImageFormat]::Tiff; break }
            '.tiff' { [ImageFormat]::Tiff; break }
            Default { [ImageFormat]::Png }
        }

        $this.SaveScreenShot($FileName, $Format)
    }
    #endregion

    #region Hidden Method
    Hidden [void]_LoadSelenium() {
        $LibPath = Join-Path $this.PSModuleRoot '\Lib'
        if (!('OpenQA.Selenium.By' -as [type])) {
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

        if (('OpenQA.Selenium.By' -as [type]) -and !('OpenQA.Selenium.Support.UI.SelectElement' -as [type])) {
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

        if (($this.Browser -eq 'EdgeChromium') -and !('Microsoft.Edge.SeleniumTools.EdgeDriver' -as [type])) {
            if (!($SeleniumPath = Resolve-Path "$LibPath\microsoft.edge.seleniumtools.*\lib\netstandard2.0" -ea SilentlyContinue)) {
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
    }

    Hidden [void]_LoadWebDriver() {
        $local:dir = [string]$this.DriverPackage
        if ($this.StrictBrowserName -eq 'IE') {
            $local:exe = 'IEDriverServer.exe'
        }
        elseif ($this.StrictBrowserName -eq 'EdgeChromium') {
            $local:exe = 'msedgedriver.exe'
        }
        else {
            $local:exe = $dir + '.exe'
        }

        if (!(Get-Command $exe -ErrorAction SilentlyContinue)) {
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
            'InternetExplorer' { 'IE'; break }
            'HeadlessChrome' { 'Chrome'; break }
            'HeadlessFirefox' { 'Firefox'; break }
            Default { $_ }
        }
        return $tmp
    }

    Hidden [string]_ParseDriverPackage([string]$BrowserName) {
        [string]$local:tmp = switch ($BrowserName) {
            'Firefox' { 'GeckoDriver'; break }
            'HeadlessFirefox' { 'GeckoDriver'; break }
            'Edge' { 'MicrosoftWebDriver'; break }
            'EdgeChromium' { 'EdgeDriver'; break }
            'IE' { 'IEDriver'; break }
            'InternetExplorer' { 'IEDriver'; break }
            'Chrome' { 'ChromeDriver'; break }
            'HeadlessChrome' { 'ChromeDriver'; break }
            default { "${_}Driver" }
        }
        return $tmp
    }

    Hidden [DriverOptions]_NewDriverOptions([string]$BrowserName) {
        $Options = $null

        switch ($BrowserName) {
            'Firefox' {
                $Options = New-Object Firefox.FirefoxOptions
                break
            }
            'HeadlessFirefox' {
                $Options = New-Object Firefox.FirefoxOptions
                $Options.AddArgument('-headless')
                break
            }
            'Edge' {
                $Options = New-Object Edge.EdgeOptions
                break
            }
            'EdgeChromium' {
                $Options = New-Object Microsoft.Edge.SeleniumTools.EdgeOptions
                $Options.UseChromium = $true
                break
            }
            'IE' {
                $Options = New-Object IE.InternetExplorerOptions
                break
            }
            'InternetExplorer' {
                $Options = New-Object IE.InternetExplorerOptions
                break
            }
            'Chrome' {
                $Options = New-Object Chrome.ChromeOptions
                break
            }
            'HeadlessChrome' {
                $Options = New-Object Chrome.ChromeOptions
                $Options.AddArgument('--headless')
                break
            }
            default {
                $Options = $null
            }
        }
        
        return $Options
    }

    Hidden [DriverService]_NewDriverService([string]$BrowserName) {
        $Service = $null

        switch ($BrowserName) {
            'Firefox' {
                $Service = [Firefox.FirefoxDriverService]::CreateDefaultService()
                if ($global:PSEdition -eq 'Core') { $Service.Host = '::1' } # Workaround for performance issue of the Firefox on .NET Core
                break
            }
            'HeadlessFirefox' {
                $Service = [Firefox.FirefoxDriverService]::CreateDefaultService()
                if ($global:PSEdition -eq 'Core') { $Service.Host = '::1' }
                break
            }
            'Edge' {
                $Service = [Edge.EdgeDriverService]::CreateDefaultService()
                break
            }
            'EdgeChromium' {
                $Service = [Microsoft.Edge.SeleniumTools.EdgeDriverService]::CreateChromiumService()
                break
            }
            'IE' {
                $Service = [IE.InternetExplorerDriverService]::CreateDefaultService()
                break
            }
            'InternetExplorer' {
                $Service = [IE.InternetExplorerDriverService]::CreateDefaultService()
                break
            }
            'Chrome' {
                $Service = [Chrome.ChromeDriverService]::CreateDefaultService()
                break
            }
            'HeadlessChrome' {
                $Service = [Chrome.ChromeDriverService]::CreateDefaultService()
                break
            }
            default {
                $Service = $null
            }
        }
        
        return $Service
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
        if ($this.Driver) { $this.SetImplicitWait(0) }

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

        if ($this.Driver) { $this.SetImplicitWait($tmpWait) }
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
                break
            }
            '^glob:(.+)' {
                $ret.Matcher = 'Like'
                $ret.Pattern = $Matches[1]
                break
            }
            '^exact:(.+)' {
                $ret.Matcher = 'Equal'
                $ret.Pattern = $Matches[1]
                break
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
            Invoke-Expression '$SelectElement = New-Object "Support.UI.SelectElement" $element' -ea Stop
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
        if (!('AnimatedGifWrapper' -as [type])) {
            switch ($global:PSEdition) {
                'Core' {
                    #PowerShell >= 7.0
                    $local:DllPath = Resolve-Path "$LibPath\AnimatedGif.1.0.5\lib\AnimatedGif.dll" -ea SilentlyContinue
                    $local:WrapperPath = Resolve-Path "$LibPath\AnimatedGifWrapper\PowerShell\AnimatedGifWrapper.dll" -ea SilentlyContinue
                    break
                }
                Default {
                    #PowerShell <= 5.1
                    $local:DllPath = Resolve-Path "$LibPath\AnimatedGif.1.0.2\lib\AnimatedGif.dll" -ea SilentlyContinue
                    $local:CSharpCode = Get-Content (Join-Path $LibPath '\AnimatedGifWrapper\WindowsPowerShell\AnimatedGifWrapper.cs') -Raw -ea SilentlyContinue
                }
            }
            try {
                [void][Reflection.assembly]::LoadFrom($local:DllPath.ToString())
                if ($local:WrapperPath) {
                    [void][Reflection.assembly]::LoadFrom($local:WrapperPath.ToString())
                }
                elseif ($local:CSharpCode) {
                    Add-Type -Language CSharp -TypeDefinition $local:CSharpCode -ReferencedAssemblies ('System.Drawing', 'System.Windows.Forms', $local:DllPath.ToString()) -ea Stop
                }
            }
            catch {
                throw "Couldn't load AnimatedGifWrapper Class"
            }
        }

        #Create background timer job
        if ('AnimatedGifWrapper' -as [type]) {
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
                $RecordAction = {
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

                Register-ObjectEvent -InputObject $this.Timer -EventName Elapsed -SourceIdentifier $this.InstanceId -Action $RecordAction -MessageData $this.InstanceId> $null
            }
            catch {
                throw 'Failed Initialize GIF Recorder'
                $this._DisposeRecorder()    #Dispose recorder when error occurred.
            }
        }
    }

    Hidden [void]_DisposeRecorder() {
        if ($this.Timer) {
            try {
                $this.Timer.Close() #Stop timer
            }
            catch { }
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
            catch { }
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

        #Check is recorder already started
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
        catch { }
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

function New-PSWebDriver {
    [CmdletBinding()]
    [OutputType([PSWebDriver])]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet('Chrome', 'Firefox', 'Edge', 'EdgeChromium', 'HeadlessChrome', 'HeadlessFirefox', 'IE', 'InternetExplorer')]
        [string]
        $Name
    )

    New-Object PSWebDriver($Name)
}

function New-Selector {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    [OutputType([Selector])]
    param(
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(Mandatory, ParameterSetName = 'WithType')]
        [string]
        $Expression,

        [Parameter(ParameterSetName = 'WithType')]
        [SelectorType]
        $Type = [SelectorType]::None
    )

    if ([string]::IsNullOrEmpty($Expression)) {
        New-Object Selector
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'Default') {
        New-Object Selector($Expression)
    }
    elseif ($PSCmdlet.ParameterSetName -eq 'WithType') {
        New-Object Selector($Expression, $Type)
    }
}
