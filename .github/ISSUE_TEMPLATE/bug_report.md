---
name: Bug report
about: Create a report to help us improve
title: "[BUG] "
labels: bug
assignees: ''

---

**Describe the bug**
A clear and concise description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:

**Expected behavior**
A clear and concise description of what you expected to happen.

**Screenshots**
If applicable, add screenshots to help explain your problem.

**Desktop (please complete the following information):**
Run the following command in PowerShell and paste the output below:
```powershell
Import-Module PS365 -ErrorAction SilentlyContinue
[PSCustomObject]@{
    OSVersion = [System.Runtime.InteropServices.RuntimeInformation]::OSDescription
    PowerShellVersion = $PSVersionTable.PSVersion.ToString()
    PS365ExtensionVersion = (Get-Module PS365 | Select-Object -First 1 -ExpandProperty Version | ForEach-Object { $_.ToString() })
} | Format-List
```

**Additional context**
Add any other context about the problem here.
