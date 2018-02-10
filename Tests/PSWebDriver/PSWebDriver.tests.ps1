#Requires -Version 5.0
#Requires -Modules @{ ModuleName="Pester"; ModuleVersion="4.1.0" }

<#
# Tests for PSWebDriver Class
#>
$moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$TestData = Join-Path $moduleRoot '\Tests\TestData\index.html'

# Import module
Get-Module 'PSWebDriver' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PSWebDriver.psd1') -Force

#Initialize
$Driver = [PSWebDriver]::New('Chrome')
$Driver.Start($TestData)

# Tests
Describe 'FindElement()' {
    Context 'FindElement([Selector]$Selector)' {
        It 'Return WebElement when element found' {
            $selector = [Selector]::New('btn1', [SelectorType]::Id)
            $Driver.FindElement($selector) | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            $selector = [Selector]::New('notexist', [SelectorType]::Id)
            {$Driver.FindElement($selector)} | Should -Throw 'no such element'
        }
    }

    Context 'FindElement([string]$SelectorExpression)' {
        It 'Return WebElement when element found' {
            $Driver.FindElement('id=btn1') | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            {$Driver.FindElement('id=notexist')} | Should -Throw 'no such element'
        }
    }

    Context 'FindElement([string]$SelectorExpression, [SelectorType]$Type)' {
        It 'Return WebElement when element found' {
            $Driver.FindElement('btn1', [SelectorType]::Id) | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
        }

        It 'Throw NoSuchElementException when element not found' {
            {$Driver.FindElement('notexist', [SelectorType]::Id)} | Should -Throw 'no such element'
        }
    }
}

Describe 'FindElements()' {
    It 'Return All WebElements when multiple elements found' {
        ($Driver.FindElements('name=hobby')).Count | Should -Be 3
    }

    It 'Return empty array when element not found' {
        ($Driver.FindElements('name=notexist')).Count | Should -Be 0
    }
}

Describe 'IsElementPresent()' {
    It 'Return $true when element found' {
        $Driver.IsElementPresent('id=btn1') | Should -Be $true
    }

    It 'Return $false when element not found' {
        $Driver.IsElementPresent('id=notexist') | Should -Be $false
    }
}

Describe 'GetTitle()' {
    It 'Return Page title' {
        $Driver.GetTitle() | Should -Be 'Test page for PSWebDriver'
    }
}

Describe 'Location' {
    BeforeEach {
        $Driver.Open($TestData)
    }

    Context 'Get Location property' {
        It 'Return current url' {
            $Driver.Location | Should -Be ([Uri]$TestData).AbsoluteUri
        }
    }

    Context 'Set Location property' {
        It 'Move to url' {
            $NewPage = Join-Path $moduleRoot '\Tests\TestData\newwindow.html'
            $Driver.Location = [Uri]$NewPage
            $Driver.Location | Should -Be ([Uri]$NewPage).AbsoluteUri
        }
    }

    Context 'GetLocation()' {
        It 'Return current url' {
            $Driver.GetLocation() | Should -Be ([Uri]$TestData).AbsoluteUri
        }
    }

    Context 'SetLocation()' {
        It 'Move to url' {
            $NewPage = Join-Path $moduleRoot '\Tests\TestData\newwindow.html'
            $Driver.SetLocation($NewPage)
            $Driver.GetLocation() | Should -Be ([Uri]$NewPage).AbsoluteUri
        }
    }

    AfterAll {
        $Driver.Open($TestData)
    }
}

Describe 'GetText()' {
    It 'Get text of the element [id="normal_text"]. Expect "This is normal text."' {
        $Driver.GetText('id=normal_text') | Should -Be 'This is normal text.'
    }
}

Describe 'IsVisible()' {
    It 'Return $true when visible element' {
        $Driver.IsVisible('name=first_name') | Should -Be $true
    }

    It 'Return $false when hidden element' {
        $Driver.IsVisible('name=hidden') | Should -Be $false
    }
}

Describe 'Click()' {
    It 'Click button element' {
        $Driver.Click('id=btnClearOutput')
        $Driver.Click('id=btn1')
        $Driver.GetText('id=output') | Should -Be 'Button1 Clicked!'
    }
}

Describe 'JavaScriptClick()' {
    BeforeEach {
        $Driver.Click('id=btnClearOutput')
    }

    It 'Click button element' {
        $Driver.GetText('id=output') | Should -Be ''
        $Driver.JavaScriptClick('id=btn1')
        $Driver.GetText('id=output') | Should -Be 'Button1 Clicked!'
    }
}

Describe 'DoubleClick()' {
    It 'Double click button element' {
        $Driver.Click('id=btnClearOutput')
        $Driver.DoubleClick('id=btn3')
        $Driver.GetText('id=output') | Should -Be 'Button3 Double Clicked!'
    }
}

Describe 'RightClick()' {
    It 'Right click button element' {
        $Driver.Click('id=btnClearOutput')
        $Driver.RightClick('id=btn1')
        $Driver.GetText('id=output') | Should -Be 'Button1 Right Clicked!'
    }
}

Describe 'SendKeys()' {
    BeforeEach {
        $Driver.FindElement('name=first_name').Clear()
    }

    It 'Input "ABC"' {
        $Driver.SendKeys('name=first_name', 'ABC')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be 'ABC'
    }

    It 'Input double byte chars "今日は"' {
        $Driver.SendKeys('name=first_name', '今日は')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be '今日は'
    }

    It 'Input Single special key "${KEY_N1}", Expect "1"' {
        $Driver.SendKeys('name=first_name', '${KEY_N1}')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be '1'
    }

    It 'Input missing special key code "${KEY_NOTHING}", Expect "${KEY_NOTHING}"' {
        $Driver.SendKeys('name=first_name', '${KEY_NOTHING}')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be '${KEY_NOTHING}'
    }

    It 'Input multiple special keys "${KEY_MULTIPLY}${KEY_SEPARATOR}${KEY_N7}${KEY_N7}", Expect "*,77"' {
        $Driver.SendKeys('name=first_name', '${KEY_MULTIPLY}${KEY_SEPARATOR}${KEY_N7}${KEY_N7}')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be '*,77'
    }

    It 'Input complex string with multiple special keys "H+${KEY_}llo${KEY_N7}${KEY_N7}@.comA${KEY_BKSP}", Expect "H+${KEY_}llo77@.com"' {
        $Driver.SendKeys('name=first_name', 'H+${KEY_}llo${KEY_N7}${KEY_N7}@.comA${KEY_BKSP}')
        $Driver.GetAttribute('name=first_name', 'value') | Should -Be 'H+${KEY_}llo77@.com'
    }
}

Describe 'ClearAndType()' {
    BeforeEach {
        $Driver.Click('id=reset')
    }

    It 'Input "ABC" to textbox that has text already' {
        $Driver.ClearAndType('name=last_name', 'ABC')
        $Driver.GetAttribute('name=last_name', 'value') | Should -Be 'ABC'
    }
}

Describe 'Select() & GetSelectedLabel()' {
    BeforeEach {
        $Driver.Click('id=reset')
    }

    It 'Select drop down' {
        $Driver.Select('name=blood', 'Type-O')
        $Driver.GetSelectedLabel('name=blood') | Should -Be 'Type-O'
    }

    It 'Throw exception when target element is not [select]' {
        {$Driver.Select('id=btn1', 'Type-AB')} | Should -Throw "Element should have been select but was input"
        {$Driver.GetSelectedLabel('id=btn1')} | Should -Throw "Element should have been select but was input"
    }
}

Describe 'SelectWindow()' {
    It 'Open and switch to new window, test page title' {
        $Driver.Click('id=btnNewWindow')
        $Driver.SelectWindow('New Window')
        $Driver.GetTitle() | Should -Be 'New Window'
    }

    It 'When target window not found, Stay current and throw NoSuchWindowException' {
        $Current = $Driver.GetTitle()
        {$Driver.SelectWindow('non existence window')} | Should -Throw 'no such window'
        $Driver.GetTitle() | Should -Be $Current
    }

    AfterAll {
        # Close and return to original window
        if ($Driver.GetTitle() -ne 'Test page for PSWebDriver') {
            if ($Driver.Driver.WindowHandles -gt 1) {
                $Driver.Close()
                $Driver.SelectWindow('Test page for PSWebDriver')
            }
        }
    }
}

Describe 'SelectFrame()' {
    BeforeAll {
        $FramePage = Join-Path $moduleRoot '\Tests\TestData\frame.html'
        $Driver.Open($FramePage)
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
        {$Driver.SelectFrame('notexist')} | Should -Throw 'No frame element found'
    }

    AfterAll {
        $Driver.Open($TestData)
    }
}

Describe 'GetHttpStatusCode()' {
    It 'Return Http Status 404' {
        #Emulate 404 using Mock is difficult...
        #TODO: Remove outside service dependency
        $Driver.GetHttpStatusCode('http://ozuma.sakura.ne.jp/httpstatus/404') | Should -Be 404
    }

    It 'Return Http Status 200' {
        Mock Invoke-WebRequest { return @{StatusCode = 200}}
        $Driver.GetHttpStatusCode('http://localhost/200') | Should -Be 200
    }

    It 'Throw Unexpected error when status code as not [int]' {
        Mock Invoke-WebRequest { return @{StatusCode = 'NOINT'}}
        {$Driver.GetHttpStatusCode('http://localhost/200')} | Should -Throw 'Unexpected Exception'
    }
}

Describe 'IsAlertPresent()' {
    BeforeEach {
        $Driver.CloseAlert()
    }

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

    AfterEach {
        $Driver.CloseAlert()
    }
}

Describe 'CloseAlert() & CloseAlertAndGetText()' {
    BeforeEach {
        $Driver.CloseAlert()
    }

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

    AfterEach {
        $Driver.CloseAlert()
    }
}

Describe 'ExecuteScript()' {
    BeforeAll {
        $Driver.Open($TestData)
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

# Finalize
if ($Driver) {
    # Close blowser
    $Driver.Quit()
    $Driver = $null
}