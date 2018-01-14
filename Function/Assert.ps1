function Assert {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [Object]
        $AssertValue,

        [Parameter(Mandatory)]
        [Object]
        $Expected,

        [Parameter()]
        [ValidateSet('Equal', 'RegExp', 'Like')]
        [string]
        $Matcher = 'Equal',

        [Parameter()]
        [string]
        $CustomMessage,

        [switch]
        $Not
    )

    Begin {
        $code = @'
    using System;
    public class AssertionFailedException : Exception
    {
        public AssertionFailedException() : base() {}
        public AssertionFailedException(string message) : base(message) {}
        public AssertionFailedException(string message, Exception innerException) : base(message, innerException) {}
    }
'@

        Add-Type -Language CSharp -TypeDefinition $code -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    }

    Process {
        $Result = $false
        $Message

        if ($Matcher -eq 'Equal') {
            $Result = [bool]($AssertValue -eq $Expected)
            if ($Not) {
                $Result = !$Result
                $Message = ("Expected: value was [{0}], but should not have been the same" -f [string]$Expected)
            }
            else {
                $Message = ("Expected [{0}] but was [{1}]" -f [string]$Expected, [string]$AssertValue)
            }
        }
        elseif ($Matcher -eq 'RegExp') {
            $Result = [bool]($AssertValue -match $Expected)
            if ($Not) {
                $Result = !$Result
                $Message = ("Expected: [{0}] to not match the expression [{1}]" -f [string]$AssertValue, [string]$Expected)
            }
            else {
                $Message = ("Expected: [{0}] to match the expression [{1}]" -f [string]$AssertValue, [string]$Expected)
            }
        }
        elseif ($Matcher -eq 'Like') {
            $Result = [bool]($AssertValue -like $Expected)
            if ($Not) {
                $Result = !$Result
                $Message = ("Expected: [{0}] not to be like the wildcard [{1}]" -f [string]$AssertValue, [string]$Expected)
            }
            else {
                $Message = ("Expected: [{0}] to be like the wildcard [{1}]" -f [string]$AssertValue, [string]$Expected)
            }
        }

        if ($CustomMessage) {
            $Message = $CustomMessage
        }

        if (!$Result) {
            throw (New-Object AssertionFailedException $Message)
        }
    }

}