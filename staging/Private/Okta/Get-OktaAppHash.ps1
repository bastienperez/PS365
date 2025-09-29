function Get-OktaAppHash {

    $AppHash = @{}

    $App = Get-OktaApp

    foreach ($CurApp in $App) {

        $Id = $CurApp.Id
        $Accessibility = ($CurApp).Accessibility
        $Visibility = ($CurApp).Visibility
        $Credentials = ($CurApp).Credentials
        $Features = ($CurApp).Features

        $AppHash[$Id] = @{
            Name                 = $CurApp.Name
            Label                = $CurApp.Label
            Status               = $CurApp.Status
            Created              = $CurApp.Created
            LastUpdated          = $CurApp.LastUpdated
            Activated            = $CurApp.Activated
            UserNameTemplate     = $Credentials.UserNameTemplate.Template
            UserNameTemplateType = $Credentials.UserNameTemplate.Type
            CredentialScheme     = $Credentials.Scheme
            Features             = ($Features -join (';'))            
        }

    }

    $AppHash
}