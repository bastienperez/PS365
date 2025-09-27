function Get-AADMemDevice {
    [cmdletbinding(DefaultParameterSetName = 'PlaceHolder')]
    param (

        [Parameter(Mandatory, ParameterSetName = 'SearchString')]
        $SearchString,

        [Parameter(Mandatory, ParameterSetName = 'ID')]
        $Id,

        [Parameter(Mandatory, ParameterSetName = 'OS')]
        [ValidateSet('IPhone', 'iOS', 'AndroidForWork', 'Windows')]
        $OS,

        [Parameter(Mandatory, ParameterSetName = 'Compliant')]
        [switch]
        $CompliantOnly,

        [Parameter(Mandatory, ParameterSetName = 'NonCompliant')]
        [switch]
        $NonCompliantOnly
    )

    if ($SearchString) {
        Get-EIDMemDeviceData -SearchString $SearchString | Select-Object -ExpandProperty Value
    }
    elseif ($Id) {
        Get-EIDMemDeviceData -Id $ID
    }
    elseif ($OS) {
        Get-EIDMemDeviceData -OS $OS | Select-Object -ExpandProperty Value
    }
    elseif ($CompliantOnly) {
        Get-EIDMemDeviceData -CompliantOnly | Select-Object -ExpandProperty Value
    }
    elseif ($NonCompliantOnly) {
        Get-EIDMemDeviceData -NonCompliantOnly | Select-Object -ExpandProperty Value
    }
    else {
        Get-EIDMemDeviceData | Select-Object -ExpandProperty Value
    }
}
