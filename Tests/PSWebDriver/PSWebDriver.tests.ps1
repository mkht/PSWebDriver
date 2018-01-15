$moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$TestData = Join-Path $moduleRoot '\Tests\TestData\index.html'

Get-Module 'PSWebDriver' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PSWebDriver.psd1') -Force

Describe 'Tests for PSWebDriver class' {
    $Driver = [PSWebDriver]::New('Chrome')
    $Driver.Start($TestData)

    Context 'FindElement()' {

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

    Context 'IsElementPresent()' {
        It 'Return $true when element found' {
            $Driver.IsElementPresent('id=btn1') | Should -Be $true
        }

        It 'Return $false when element not found' {
            $Driver.IsElementPresent('id=notexist') | Should -Be $false
        }
    }

    Context 'GetTitle()' {
        It 'Return Page title' {
            $Driver.GetTitle() | Should -Be 'Test page for PSWebDriver'
        }
    }

    Context 'GetAttribute()' {
        It 'Get class attribute of the element [id="normal_text"] Expect "normal_class"' {
            $Driver.GetAttribute('id=normal_text', 'class') | Should -Be 'normal_class'
        }
    }

    Context 'GetText()' {
        It 'Get text of the element [id="normal_text"]. Expect "This is normal text."' {
            $Driver.GetText('id=normal_text') | Should -Be 'This is normal text.'
        }
    }

    Context 'IsVisible()' {
        It 'Return $true when visible element' {
            $Driver.IsVisible('name=first_name') | Should -Be $true
        }

        It 'Return $false when hidden element' {
            $Driver.IsVisible('name=hidden') | Should -Be $false
        }
    }

    Context 'Click()' {
        It 'Click button element' {
            $Driver.Click('id=btnClearOutput')
            $Driver.Click('id=btn1')
            $Driver.GetText('id=output') | Should -Be 'Button1 Clicked!'
        }
    }

    Context 'SendKeys()' {
        BeforeEach {
            $Driver.Click('id=reset')
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

    Context 'ClearAndType()' {
        BeforeEach {
            $Driver.Click('id=reset')
        }

        It 'Input "ABC" to textbox that has text already' {
            $Driver.ClearAndType('name=last_name', 'ABC')
            $Driver.GetAttribute('name=last_name', 'value') | Should -Be 'ABC'
        }
    }

    Context 'Select() & GetSelectedLabel()' {
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

    Context 'SelectWindow()' {
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

    if ($Driver) {
        $Driver.Quit()
        $Driver = $null
    }
}