
<##################################
# Tests for PSWebDriver
###################################>

#Requires -Version 5.0
#Requires -Modules @{ ModuleName="Pester"; ModuleVersion="5.0.2" }

# Import module
$script:moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
Get-Module 'PSWebDriver' | Remove-Module -Force
Import-Module (Join-Path $script:moduleRoot './PSWebDriver.psd1') -Force

# TestData
$script:moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$script:TestData = Join-Path $script:moduleRoot '\Tests\TestData\index.html'

BeforeAll {
    # Specify Browser
    if (-not $env:TARGET_BROWSER) {
        $global:Browser = 'Chrome'
    }
    else {
        $global:Browser = $env:TARGET_BROWSER
    }
}

Describe 'Find element(s)' {

    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    Context 'FindElement([Selector]$Selector)' {
        It 'Return WebElement when element found' {
            $selector = New-Selector -Expression 'btn1' -Type Id
            $Driver.FindElement($selector) | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            $selector = New-Selector -Expression 'notexist' -Type Id
            { $Driver.FindElement($selector) } | Should -Throw -ErrorId "NoSuchElementException"
        }
    }

    Context 'FindElement([string]$SelectorExpression)' {
        It 'Return WebElement when element found' {
            $Driver.FindElement('id=btn1') | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            { $Driver.FindElement('id=notexist') } | Should -Throw -ErrorId "NoSuchElementException"
        }
    }

    Context 'FindElement([string]$SelectorExpression, [SelectorType]$Type)' {
        It 'Return WebElement when element found' {
            $Driver.FindElement('btn1', 'Id') | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            { $Driver.FindElement('notexist', 'Id') } | Should -Throw -ErrorId "NoSuchElementException"
        }
    }

    Context 'FindElements()' {
        It 'Return All WebElements when multiple elements found' {
            ($Driver.FindElements('name=hobby')).Count | Should -Be 3
        }
        
        It 'Return empty array when the element not found' {
            ($Driver.FindElements('name=notexist')).Count | Should -Be 0
        }
    }

    Context 'IsElementPresent()' {
        It 'Return $true when the element found' {
            $Driver.IsElementPresent('id=btn1') | Should -Be $true
        }
        
        It 'Return $false when the element not found' {
            $Driver.IsElementPresent('id=notexist') | Should -Be $false
        }
    }
}

Describe 'Get Browser Info & Title' {

    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    Context 'Browser Info' {
        It 'Return Browser Information' {
            $local:Ret = $Driver.GetBrowserInfo()
            $local:Ret | Should -BeOfType 'HashTable'
            $local:Ret.browserName | Should -Not -BeNullOrEmpty
            $local:Ret.browserVersion | Should -Not -BeNullOrEmpty
            $local:Ret.platformName | Should -Not -BeNullOrEmpty
        }
    }

    Context 'Title' {
        It 'Return Page title' {
            $Driver.GetTitle() | Should -Be 'Test page for PSWebDriver'
        }
    }
}


Describe 'Location' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    BeforeEach {
        $Driver.Open($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Location property should return current url' {
        $Driver.Location | Should -Be ([Uri]$script:TestData).AbsoluteUri
    }
    
    It 'Should move to url when the Location property is set' {
        $NewPage = Join-Path $moduleRoot '\Tests\TestData\newwindow.html'
        $Driver.Location = [Uri]$NewPage
        $Driver.Location | Should -Be ([Uri]$NewPage).AbsoluteUri
    }
    
    It 'GetLocation() method should return current url' {
        $Driver.GetLocation() | Should -Be ([Uri]$script:TestData).AbsoluteUri
    }
    
    It 'SetLocation() method should move to specified url' {
        $NewPage = Join-Path $moduleRoot '\Tests\TestData\newwindow.html'
        $Driver.SetLocation($NewPage)
        $Driver.GetLocation() | Should -Be ([Uri]$NewPage).AbsoluteUri
    }
}

Describe 'GetText()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Get inner text of the element' {
        $Driver.GetText('id=normal_text') | Should -Be 'This is normal text.'
    }
}

Describe 'IsVisible()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Return $true when the element is visible' {
        $Driver.IsVisible('name=first_name') | Should -Be $true
    }
    
    It 'Return $false when the element is hidden' {
        $Driver.IsVisible('name=hidden') | Should -Be $false
    }
}

Describe Click {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }
    
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }
        
    Context 'Click()' {
        It 'Click button element' {
            $Driver.Click('id=btnClearOutput')
            $Driver.Click('id=btn1')
            $Driver.GetText('id=output') | Should -Be 'Button1 Clicked!'
        }
    }

    Context 'JavaScriptClick()' {
        It 'Click button element' {
            $Driver.Click('id=btnClearOutput')
            $Driver.GetText('id=output') | Should -Be ''
            $Driver.JavaScriptClick('id=btn1')
            $Driver.GetText('id=output') | Should -Be 'Button1 Clicked!'
        }
    }

    Context 'DoubleClick()' {
        It 'Double click button element' {
            $Driver.Click('id=btnClearOutput')
            $Driver.DoubleClick('id=btn3')
            $Driver.GetText('id=output') | Should -Be 'Button3 Double Clicked!'
        }
    }

    Context 'RightClick()' {
        It 'Right click button element' {
            $Driver.Click('id=btnClearOutput')
            $Driver.RightClick('id=btn1')
            $Driver.GetText('id=output') | Should -Be 'Button1 Right Clicked!'
        }
    }
}

Describe 'Input keys' {

    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    BeforeEach {
        $Driver.FindElement('name=first_name').Clear()
        $Driver.FindElement('name=last_name').Clear()
    }
        
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }
           
    Context 'SendKeys()' {
        It 'Input "ABC"' {
            $Driver.SendKeys('name=first_name', 'ABC')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be 'ABC'
        }

        It 'Input "ABC" to textbox that has text already' {
            $Driver.SendKeys('name=first_name', 'Already Has Text !')
            $Driver.SendKeys('name=first_name', 'ABC')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be 'Already Has Text !ABC'
        }

        It 'Input double byte chars "今日は"' {
            $Driver.SendKeys('name=first_name', '今日は')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be '今日は'
        }

        It 'Input Single special key "${KEY_N1}", Expect "1"' {
            if ($global:Browser -match 'Firefox') {
                Set-ItResult -Skipped -Because 'Firefox does not recognize NumberPad keys'
            }

            $Driver.SendKeys('name=first_name', '${KEY_N1}')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be '1'
        }

        It 'Input missing special key code "${KEY_NOTHING}", Expect "${KEY_NOTHING}"' {
            $Driver.SendKeys('name=first_name', '${KEY_NOTHING}')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be '${KEY_NOTHING}'
        }

        It 'Input multiple special keys "${KEY_MULTIPLY}${KEY_SEPARATOR}${KEY_N7}${KEY_N7}", Expect "*,77"' {
            if ($global:Browser -match 'Firefox') {
                Set-ItResult -Skipped -Because 'Firefox does not recognize NumberPad keys'
            }

            $Driver.SendKeys('name=first_name', '${KEY_MULTIPLY}${KEY_SEPARATOR}${KEY_N7}${KEY_N7}')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be '*,77'
        }

        It 'Input complex string with multiple special keys "H+${KEY_}llo${KEY_N7}${KEY_N7}@.comA${KEY_BKSP}", Expect "H+${KEY_}llo77@.com"' {
            if ($global:Browser -match 'Firefox') {
                Set-ItResult -Skipped -Because 'Firefox does not recognize NumberPad keys'
            }

            $Driver.SendKeys('name=first_name', 'H+${KEY_}llo${KEY_N7}${KEY_N7}@.comA${KEY_BKSP}')
            $Driver.GetAttribute('name=first_name', 'value') | Should -Be 'H+${KEY_}llo77@.com'
        }
    }

    Context 'ClearAndType()' {
        It 'Input "ABC" to textbox that has text already' {
            $Driver.SendKeys('name=last_name', 'Already Has Text !')
            $Driver.ClearAndType('name=last_name', 'ABC')
            $Driver.GetAttribute('name=last_name', 'value') | Should -Be 'ABC'
        }
    }
}

Describe 'Select() & GetSelectedLabel()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    BeforeEach {
        $Driver.Click('id=reset')
    }
    
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Select drop down' {
        $Driver.Select('name=blood', 'Type-O')
        $Driver.GetSelectedLabel('name=blood') | Should -Be 'Type-O'
    }

    It 'Throw exception when target element is not [select]' {
        { $Driver.Select('id=btn1', 'Type-AB') } | Should -Throw "*Element should have been select but was input*"
        { $Driver.GetSelectedLabel('id=btn1') } | Should -Throw "*Element should have been select but was input*"
    }
}

Describe 'SelectWindow()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }
    
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Open and switch to new window, test page title' {
        $Driver.Click('id=btnNewWindow')
        $Driver.SelectWindow('New Window')
        $Driver.GetTitle() | Should -Be 'New Window'
    }

    It 'When target window not found, Stay current and throw exception' {
        $Current = $Driver.GetTitle()
        { $Driver.SelectWindow('non existence window') } | Should -Throw
        $Driver.GetTitle() | Should -Be $Current
    }
}

Describe 'SelectFrame()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $FramePage = Join-Path $moduleRoot '\Tests\TestData\frame.html'
        $Driver.Start($FramePage)
    }
    
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    It 'Select left frame then get text' {
        $Driver.SelectFrame('leftFrame')
        $Driver.GetText('id=left-text') | Should -Be 'This is left frame.'
    }

    It 'Select right frame then get text' {
        $Driver.SelectFrame('rightFrame')
        $Driver.GetText('id=right-text') | Should -Be 'This is right frame.'
    }

    It 'Select non present frame then throw exception' {
        { $Driver.SelectFrame('notexist') } | Should -Throw '*No frame element found*'
    }
}

Describe 'GetHttpStatusCode()' {
    BeforeAll {
        # Initialization
        $Driver = New-PSWebDriver -Name $global:Browser
    }
    
    AfterAll {
        # Teardown
        $Driver.Quit()
        $Driver = $null
    }

    It 'Return Http Status 404' {
        #Emulate 404 using Mock is difficult...
        $Driver.GetHttpStatusCode('https://httpstat.us/404') | Should -BeExactly 404
    }
    
    It 'Return Http Status 200' {
        $Driver.GetHttpStatusCode('https://httpstat.us/200') | Should -BeExactly 200
    }
}

Describe 'Alert' {

    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    BeforeEach {
        $Driver.CloseAlert()
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    Context 'IsAlertPresent()' {
        It 'Return $true when alert present' {
            $Driver.Click('id=btnAlert')
            $Driver.IsAlertPresent() | Should -Be $true
        }
    
        It 'Return $false when alert not present' {
            $Driver.IsAlertPresent() | Should -Be $false
        }
    
        It 'Return $true when confirm present' {
            $Driver.Click('id=btnConfirm')
            $Driver.IsAlertPresent() | Should -Be $true
        }
    
        It 'Return $true when prompt present' {
            $Driver.Click('id=btnPrompt')
            $Driver.IsAlertPresent() | Should -Be $true
        }
    }
    
    Context 'CloseAlert() & CloseAlertAndGetText()' {
        It 'Close alert' {
            $Driver.Click('id=btnAlert')
            $Driver.IsAlertPresent() | Should -Be $true
            $Driver.CloseAlert()
            $Driver.IsAlertPresent() | Should -Be $false
        }
    
        It 'Get alert text & close' {
            $Driver.Click('id=btnAlert')
            $Driver.IsAlertPresent() | Should -Be $true
            $Driver.CloseAlertAndGetText($true) | Should -Be 'This is Alert!'
            $Driver.IsAlertPresent() | Should -Be $false
        }
    
        It 'Get confirm text & Accept confirm' {
            $Driver.Click('id=btnConfirm')
            $Driver.IsAlertPresent() | Should -Be $true
            $Driver.CloseAlertAndGetText($true) | Should -Be 'Chose an option.'
            $Driver.IsAlertPresent() | Should -Be $false
            $Driver.GetText('id=output') | Should -Be 'Confirmed.'
        }
    
        It 'Get confirm text & Dismiss confirm' {
            $Driver.Click('id=btnConfirm')
            $Driver.IsAlertPresent() | Should -Be $true
            $Driver.CloseAlertAndGetText($false) | Should -Be 'Chose an option.'
            $Driver.IsAlertPresent() | Should -Be $false
            $Driver.GetText('id=output') | Should -Be 'Rejected!'
        }
    }
}

Describe 'ExecuteScript()' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }
    
    It 'Execute Javascript' {
        $Driver.ExecuteScript('output("Execute Code 1");')
        $Driver.GetText('id=output') | Should -Be 'Execute Code 1'
    }
    
    It 'Execute Javascript and get result as [String]' {
        $Obj = $Driver.ExecuteScript('return document.title;')
        $Obj | Should -Be 'Test page for PSWebDriver'
        $Obj | Should -BeOfType [string]
    }
    
    It 'Execute Javascript and get result as [IWebElement]' {
        $Driver.ExecuteScript('return document.getElementById("normal_text");') | Should -BeOfType 'OpenQA.Selenium.IWebElement'
    }
    
    It 'Execute Javascript and get result as [int64]' {
        $Obj = $Driver.ExecuteScript('return 1;')
        $Obj | Should -Be 1
        $Obj | Should -BeOfType [int64]
    }
    
    It 'Execute Javascript and get result as [bool]' {
        $Obj = $Driver.ExecuteScript('return true;')
        $Obj | Should -Be $true
        $Obj | Should -BeOfType [bool]
    }
    
    It 'Execute Javascript and get result as [Array]' {
        $Obj = $Driver.ExecuteScript("return ['a', 'b', 123];")
        $Obj | Should -Be @('a', 'b', 123)
        $Obj.GetType().FullName | Should -Be 'System.Object[]'
    }
    
    It 'Execute Javascript with args' {
        $Driver.ExecuteScript('return arguments[0] + arguments[1];', [string[]]('Apple', 'Banana')) | Should -Be 'AppleBanana'
    }
}

Describe 'ScreenShot' {

    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
        $Driver.Start($script:TestData)
    }

    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }
    
    It 'Save Screen Shot with default image format' {
        $RandomFileName = Join-Path -Path $TestDrive -ChildPath ([System.IO.Path]::GetRandomFileName())
        $Driver.SaveScreenShot($RandomFileName)
        $RandomFileName | Should -Exist
    }
    
    It 'Save Screen Shot with specified image format' {
        $RandomFileName = Join-Path -Path $TestDrive -ChildPath ([System.IO.Path]::GetRandomFileName())
        $Driver.SaveScreenShot($RandomFileName, 'Tiff')
        $RandomFileName | Should -Exist
    }
    
    It 'Should throw exception when the invalid image format specified' {
        $RandomFileName = Join-Path -Path $TestDrive -ChildPath ([System.IO.Path]::GetRandomFileName())
        { $Driver.SaveScreenShot($RandomFileName, 'INVALID') } | Should -Throw
        $RandomFileName | Should -Not -Exist
    }
    
    It 'Save Screen Capture (Animation)' {
        $RandomFileName = Join-Path -Path $TestDrive -ChildPath ([System.IO.Path]::GetRandomFileName())
        $Driver.StartAnimationRecord()
        Start-Sleep -Seconds 3
        $Driver.StopAnimationRecord($RandomFileName)
        $RandomFileName | Should -Exist
    }
    
    It 'Animation capture interval should not less than 500ms' {
        { $Driver.StartAnimationRecord(499) } | Should -Throw
        { $Driver.StartAnimationRecord(500) } | Should -Not -Throw
        Start-Sleep -Seconds 1
        $Driver.StopAnimationRecord()
    }
}

Describe 'Miscellaneous Tests' {
    BeforeAll {
        # Initialization
        # Start Browser
        $Driver = New-PSWebDriver -Name $global:Browser
    }

    BeforeEach {
        $Driver.Quit()
    }
        
    AfterAll {
        # Teardown
        # Stop Browser
        $Driver.Quit()
    }

    Context 'BrowserOption: AcceptInsecureCertificates' {
        It 'Accept Certificate Errors' {
            if ($global:Browser -match 'Edge') {
                Set-ItResult -Skipped -Because 'Legacy Edge does not support "AcceptInsecureCertificates" option'
            }

            $TestURL = 'https://expired.badssl.com/'
            $Driver.BrowserOptions.AcceptInsecureCertificates = $true
            { $Driver.Start($TestURL) } | Should -Not -Throw
            $Driver.IsElementPresent('id=content') | Should -BeTrue
        }

        It 'DO NOT Accept Certificate Errors' {
            if ($global:Browser -match 'Firefox') {
                Set-ItResult -Skipped -Because 'This test freezes sometimes on Firefox (Investigating)'
            }

            $TestURL = 'https://expired.badssl.com/'
            $Driver.BrowserOptions.AcceptInsecureCertificates = $false
            try { $Driver.Start($TestURL) }catch { } # Whether an exception raises depend on the browser (Firefox will throw, but Chrome will not.)
            $Driver.IsElementPresent('id=content') | Should -BeFalse
        }
    }
}

# Remove global variables
Remove-Variable -Name Browser -Scope global -ErrorAction Ignore
