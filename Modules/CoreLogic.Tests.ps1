$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$modulePath = Join-Path $here "CoreLogic.psm1"

Describe "Clean-WmiString" {
    BeforeAll {
        Import-Module $modulePath -Force
    }

    It "Should return null when input is null" {
        Clean-WmiString $null | Should -BeNullOrEmpty
    }

    It "Should return empty string when input is empty string" {
        Clean-WmiString "" | Should -Be ""
    }

    It "Should return the same string when it contains only valid characters" {
        $input = "Hello World 123! @#$%^&*()_+"
        Clean-WmiString $input | Should -Be $input
    }

    It "Should preserve valid whitespace characters (Tab, LF, CR)" {
        $input = "Line1`r`nLine2`tTabbed"
        Clean-WmiString $input | Should -Be $input
    }

    It "Should remove invalid control characters (e.g., Bell \x07)" {
        $input = "Good" + [char]0x07 + "Bad"
        Clean-WmiString $input | Should -Be "GoodBad"
    }

    It "Should remove null characters" {
        $input = "A" + [char]0x00 + "B"
        Clean-WmiString $input | Should -Be "AB"
    }

    It "Should preserve high Unicode characters in allowed ranges" {
        $input = "Unicode: " + [char]0x1234 + " " + [char]0x4321
        Clean-WmiString $input | Should -Be $input
    }

    It "Should remove Unicode surrogate characters (\uD800-\uDFFF)" {
        # D800 is the start of the high surrogate range
        $input = "Safe" + [char]0xD800 + "Unsafe"
        Clean-WmiString $input | Should -Be "SafeUnsafe"
    }

    It "Should remove non-characters \uFFFE and \uFFFF" {
        $input = "Start" + [char]0xFFFE + [char]0xFFFF + "End"
        Clean-WmiString $input | Should -Be "StartEnd"
    }
}
