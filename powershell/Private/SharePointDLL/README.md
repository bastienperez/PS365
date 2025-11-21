# SharePoint Online Client Components SDK

This folder contains the official files from the **SharePoint Online Client Components SDK**.

## Included Files

| File Name | Last Modified | SHA256 Hash |
|---|---|---|
| Microsoft.SharePoint.Client.dll | 05/09/2017 23:54:26 | 06C11A4BC1D7CDEFF9CD1F23C70724E54639D6A88E7B4A5FAB4FEC89615A5D96 |
| Microsoft.SharePoint.Client.Runtime.dll | 05/09/2017 23:54:26 | A15614D20253852F0BCA360AFBD7104F41A0D615D979517EFAAC3A89F7F1C9F7 |

## Description

These DLLs are Microsoft's official client components for interacting with SharePoint Online via the CSOM (Client Side Object Model) APIs.

**No installation required** - the module contains these DLL files directly. 

These files were originally obtained by installing the SharePoint Online Client Components SDK and copying them from:
`%ProgramFiles%\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI`

- **Microsoft.SharePoint.Client.dll**: Contains the main classes for SharePoint interaction
- **Microsoft.SharePoint.Client.Runtime.dll**: Contains the runtime components required for SharePoint client functionality

## Integrity Verification

The SHA256 hashes above allow you to verify file integrity. To recalculate the hashes:

```powershell
Get-ChildItem -Path . -Filter *.dll | ForEach-Object {
    $h = Get-FileHash $_.FullName -Algorithm SHA256
    [PSCustomObject]@{
        Name         = $_.Name
        LastWriteTime= $_.LastWriteTime
        Hash         = $h.Hash
    }
} | Format-Table -AutoSize
``` 