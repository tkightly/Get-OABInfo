Param (
    [array]$ComputerName = $env:computername,
    [parameter(Mandatory=$true)]
    [string]$OABGuid
)
function Get-OABInfo () {
    New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
    $UserHives = Get-ChildItem -Path HKU:
    $OAB = New-Object -TypeName psobject

    $UserHives | ForEach-Object {
        if (Test-Path "HKU:\$_\Software\Microsoft\Exchange\Exchange Provider\OABs\$using:OABGuid") {
            $OAB = New-Object -TypeName psobject
            $data = Get-ItemProperty "HKU:\$_\Software\Microsoft\Exchange\Exchange Provider\OABs\$using:OABGuid" | Select-Object -ExpandProperty 'OAB Last Modified Time'
            $OAB | Add-Member -MemberType NoteProperty -Name LastModified -Value (Get-Date -Format "dd/MM/yyyy HH:mm:ss" -Date ([DateTime]::FromFileTime( (((((($data[7] * 256 + $data[6]) * 256 + $data[5]) * 256 + $data[4]) * 256 + $data[3]) * 256 + $data[2]) * 256 + $data[1]) * 256 + $data[0]) ))
            $OAB | Add-Member -MemberType NoteProperty -Name Name -Value (Get-ItemProperty "HKU:\$_\Software\Microsoft\Exchange\Exchange Provider\OABs\$using:OABGuid" | Select-Object -ExpandProperty 'OAB NameU')
            $OAB | Add-Member -MemberType NoteProperty -Name UserName -Value (Get-ItemProperty "HKU:\$_\Volatile Environment" | Select-Object -ExpandProperty 'USERNAME')
            $OAB
        }
    }
}
$DataToReturn = Invoke-Command -ComputerName $ComputerName -ScriptBlock ${Function:Get-OABInfo}
$DataToReturn