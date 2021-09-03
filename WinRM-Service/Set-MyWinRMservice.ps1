<#
.SYNOPSIS
Set-MyWinRMservice configura el servicio WinRM en un equipo remoto para atender peticiones.
.DESCRIPTION
Set-MyWinRMservice usa la clase system.management.managementclass con la que se crean las entradas
necesarias en el registro para configurar el servicio WinRM.
.PARAMETER ComputerName
Equipo/Equipos remotos en los que se configurará el servicio WinRM.
.EXAMPLE
Set-MyWinRMservice 
        ó
Set-MyWinRMservice -ComputerName NombreEquipo
        ó
Get-Content .\equipos.txt | Set-MyWinRMservice
.NOTES
Necesita reiniciar el servicio para que la configuración esté operativa, puede usar la función
Restart-MyWinRMservice en Restart-MyWinRMservice.ps1 
#>
Function Set-MyWinRMservice{

    [cmdletBinding()] 
    Param
    (
    [Parameter( 
    ValueFromPipeline=$True,    # ValueFromPipeline: especifica que el parámetro acepta la entrada de un objeto de canalización.
    ValueFromPipelineByPropertyName=$True,      # ValueFromPipelinePropertyName: Acepta la entrada de una propiedad de un objeto de canalización.
    Mandatory=$True,
    HelpMessage='Introduzca una IP o nombre de equipo, acepta multiples ComputerName')] 
    [String[]]$ComputerName
    )

    Begin 
    {
    Write-Debug 'Apertura del bloque Begin'
    $HKLM = 2147483650
    $Key = 'SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
    $DWORDName = 'AllowAutoConfig' 
    $DWORDvalue = '0x1'
    Write-Debug 'Cierre del bloque Begin con las variables ya definidas'
    }
    Process 
    {
            Foreach ($computer in $ComputerName)
            {
                    Write-Host
                    Write-Host
                    Write-Host "Iniciando la función en el equipo $computer" -ForegroundColor Yellow
                    Write-Host
                    Write-Debug "Apertura del bloque Process para el equipo $Computer"
                    Try
                    {
                    Write-Host 'Creando el constructor para manejar el registro remoto....'
                    Write-Host
                    $Reg = New-Object -TypeName System.Management.ManagementClass -ArgumentList \\$computer\Root\default:StdRegProv
                    }
                    Catch 
                    {
                    Write-Warning $_.exception.message
                    Write-Warning "La función aboratará operaciones en el equipo $Computer"
                    break
                    }
                    Try 
                    {
                    Write-Host 'Creando la clave HKLM en el registro remoto....'
                    Write-Host
                    if (($reg.CreateKey($HKLM, $key)).returnvalue -ne 0) {Throw 'Ocurrió un fallo al intentar crear la clave'}
                    }
                    Catch 
                    {
                    Write-Warning $_.exception.message
                    Write-Warning "La función aboratará operaciones en el equipo $Computer"
                    break
                    }
                    Try 
                    {
                    Write-Host 'Creando la configuración del valor DWORD....'
                    Write-Host
                    if (($reg.SetDWORDValue($HKLM, $Key, $DWORDName, $DWORDvalue)).ReturnValue -ne 0) {Throw 'Ocurrió un fallo al intentar crear el valor DWORD'}
                    }
                    Catch 
                    {
                    Write-Warning $_.exception.message
                    Write-Warning "La función aboratará operaciones en el equipo $Computer"
                    break
                    }
                    Write-Host "La operación se completó satisfactoriamente en el equipo $computer" -ForegroundColor Yellow
                    Write-Host
                    Write-Host
                    Write-Debug "Cierre del bloque Process para el equipo $Computer"
            }
    }
    End 
    {}
}
