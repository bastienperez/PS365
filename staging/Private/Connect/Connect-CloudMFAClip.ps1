function Connect-CloudMFAClip {
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory, Position = 0)]
        $CredFile
    )
    end {
        [System.Management.Automation.PSCredential]$Credential = Import-Clixml -Path $CredFile
        $Username = $Credential.Username
        $Password = $Credential.GetNetworkCredential().Password

        $null = Start-RSJob -ArgumentList $Username, $Password, $CredFile -ScriptBlock {
            param($Username, $Password, $CredFile)

            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.Application]::EnableVisualStyles()

            $PS365 = New-Object system.Windows.Forms.Form
            $PS365.ClientSize = '533,565'
            $PS365.text = 'PS365'
            $PS365.BackColor = '#333333'
            $PS365.TopMost = $false

            $Panel2 = New-Object system.Windows.Forms.Panel
            $Panel2.height = 172
            $Panel2.width = 516
            $Panel2.BackColor = '#007acc'
            $Panel2.location = New-Object System.Drawing.Point(7, 6)

            $Username = New-Object system.Windows.Forms.Button
            $Username.BackColor = '#252526'
            $Username.text = 'Copy Username'
            $Username.width = 486
            $Username.height = 72
            $Username.location = New-Object System.Drawing.Point(13, 11)
            $Username.Font = 'Microsoft Sans Serif,40'
            $Username.ForeColor = '#ffffff'
            $Username.Add_Click( { $Using:Username | CLIP })

            $Password = New-Object system.Windows.Forms.Button
            $Password.BackColor = '#252526'
            $Password.text = 'Copy Password'
            $Password.width = 486
            $Password.height = 72
            $Password.location = New-Object System.Drawing.Point(13, 90)
            $Password.Font = 'Microsoft Sans Serif,40'
            $Password.ForeColor = '#ffffff'
            $Password.Add_Click( { $Using:Password | CLIP })

            $Panel1 = New-Object system.Windows.Forms.Panel
            $Panel1.height = 364
            $Panel1.width = 516
            $Panel1.BackColor = '#007acc'
            $Panel1.location = New-Object System.Drawing.Point(8, 187)

            $Close = New-Object system.Windows.Forms.Button
            $Close.BackColor = '#252526'
            $Close.text = 'Close'
            $Close.width = 500
            $Close.height = 93
            $Close.location = New-Object System.Drawing.Point(8, 18)
            $Close.Font = 'Microsoft Sans Serif,24,style=Bold'
            $Close.ForeColor = '#ffffff'
            $Close.Add_Click( { $PS365.Close() })

            $PS365.controls.AddRange(@($Panel2, $Panel1))
            $Panel2.controls.AddRange(@($Username, $Password))
            $Panel1.controls.AddRange(@($Close))

            [void]$PS365.ShowDialog()
        }
    }
}