Repository onboarding instructions for GitHub Copilot coding agent  
Scope: Applies to this PowerShell repository. Trust these instructions. Only perform codebase searches when the information here is incomplete or proven incorrect.

GOALS
- Reduce PR rejections due to CI/validation failures or misbehavior.
- Minimize shell failures.
- Help the agent complete tasks quickly with fewer exploratory steps.

HIGH-LEVEL DETAILS
- Repository purpose: provide a set of PowerShell functions to manage a Microsoft 365 tenant.
- Project type: PowerShell module.
- Language/runtime: PowerShell 7.x (pwsh) required.

PROJECT LAYOUT
- PS365.psd1 : module manifest.
- PS365.psm1 : main module file; imports/exports functions from /Public and /Private.
- /Public : public functions exported by the module (Microsoft 365 tenant management).
- /Private : private/internal helper functions not exported.
- README.md : usage and installation.

ENVIRONMENT
- Supported OS: Windows, Linux, macOS with pwsh.
- Required environment variables: <list if applicable> (for Microsoft 365 authentication or API access).

POWERSHELL STYLE RULES
- Use full command names (Get-ChildItem instead of gci).
- Use camelCase for variables and PascalCase for function parameters.
- Use [System.Collections.Generic.List[PSCustomObject]]$array = @() for arrays; avoid [System.Collections.Generic.List[PSCustomObject]]::new().
- Use foreach loops, $null to discard output, and splatting for parameters.
- End scripts with return instead of exit.
- Prefer single quotes (') over double quotes (") unless interpolation is required.
- Use structured exception handling and explicit typing for arrays.
- Write clear English comments: block comments for sections, inline comments for short lines.
- Avoid over-commenting.
