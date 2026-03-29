# ============================================================================
# UIEvents.psm1 - UI Event Routing & Module Loader
# ============================================================================
# This module has been refactored! 
# It now acts as a router that imports specialized sub-modules to keep the codebase clean.

Import-Module "$PSScriptRoot\UI.Core.psm1" -Global -Force -DisableNameChecking
Import-Module "$PSScriptRoot\UI.UserActions.psm1" -Global -Force -DisableNameChecking
Import-Module "$PSScriptRoot\UI.ComputerActions.psm1" -Global -Force -DisableNameChecking
Import-Module "$PSScriptRoot\UI.Freshservice.psm1" -Global -Force -DisableNameChecking

function Register-MainUIEvents {
    param(
        [Parameter(Mandatory=$true)] $Window,
        [Parameter(Mandatory=$true)] $Config,
        [Parameter(Mandatory=$true)] $State
    )
    
    # Ensure a cross-module Actions dictionary exists in State
    if (-not $State.Actions) {
        $State | Add-Member -MemberType NoteProperty -Name "Actions" -Value @{} -Force
    }
    
    # 1. Register Core & Grid Events
    Register-CoreUIEvents -Window $Window -Config $Config -State $State
    
    # 2. Register AD User Context Menus & Overlays
    Register-UserUIEvents -Window $Window -Config $Config -State $State
    
    # 3. Register AD Computer Tools & Tabs
    Register-ComputerUIEvents -Window $Window -Config $Config -State $State
    
    # 4. Register ITIL / Freshservice Integrations
    Register-FreshserviceUIEvents -Window $Window -Config $Config -State $State

    # ------------------------------------------------------------------------
    # 5. CUSTOM EVENT BINDINGS: Remote Software Manager
    # ------------------------------------------------------------------------
    
}

Export-ModuleMember -Function Register-MainUIEvents