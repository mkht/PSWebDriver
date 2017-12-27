$moduleRoot = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent

Get-Module 'PSWebDriver' | Remove-Module -Force
Import-Module (Join-Path $moduleRoot './PSWebDriver.psd1') -Force

Describe 'Tests for PSWebDriver class' {
    $Driver = [PSWebDriver]::New('Chrome')
    $Driver.Start('https://www.google.co.jp/')

    Context 'FindElement' {

        Context 'FindElement with [String]' {
            It 'Return WebElement when element found' {
                $Driver.FindElement('id=lst-ib') | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
            }

            It 'Return $null when element not found' {
                $Driver.FindElement('id=notexist') | Should -BeNullOrEmpty
            }
        }

        Context 'FindElement with [Selector]' {
            It 'Return WebElement when element found' {
                $selector = [Selector]::New('lst-ib', [SelectorType]::Id)
                $Driver.FindElement($selector) | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
            }

            It 'Return $null when element not found' {
                $selector = [Selector]::New('notexist', [SelectorType]::Id)
                $Driver.FindElement($selector) | Should -BeNullOrEmpty
            }
        }

        Context 'FindElement with [string] & [SelectorType]' {
            It 'Return WebElement when element found' {
                $Driver.FindElement('lst-ib', [SelectorType]::Id) | Should -BeOfType 'OpenQA.Selenium.Remote.RemoteWebElement'
            }

            It 'Return $null when element not found' {
                $Driver.FindElement('notexist', [SelectorType]::Id) | Should -BeNullOrEmpty
            }
        }
    }

    Context 'IsElementPresent' {

        It 'Return $true when element found' {
            $Driver.IsElementPresent('id=lst-ib') | Should -Be $true
        }

        It 'Return $false when element not found' {
            $Driver.IsElementPresent('id=notexist') | Should -Be $false
        }
    }

    Context 'GetTitle' {
        It 'Return Page title' {
            $Driver.GetTitle() | Should -Be 'Google'
        }
    }

    Context 'Click' {
        It 'Click link' {
            $Driver.Click('link=Googleについて')
            $Driver.GetTitle() | Should -Match 'Google について'
        }
    }

    if ($Driver) {
        $Driver.Quit()
        $Driver = $null
    }
}