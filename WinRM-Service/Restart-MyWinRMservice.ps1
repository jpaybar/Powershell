<#
.SYNOPSIS
Restart-MyWinRMservice reinicia el servicio WinRM en un equipo/s remoto.
.DESCRIPTION
Restart-MyWinRMservice hace uso del CmdLet Get-WmiObject y de la clase win32_service para reiniciar el servicio.
.PARAMETER ComputerName
Equipo/Equipos remotos en los que se configurará el servicio WinRM.
.EXAMPLE
Restart-MyWinRMservice 
        ó
Restart-MyWinRMservice -ComputerName NombreEquipo
        ó
Get-Content .\equipos.txt | Restart-MyWinRMservice
.NOTES
General notes
#>
Function Restart-MyWinRMservice{

    [CmdletBinding()]
    Param 
    (
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Introduzca una IP o nombre de equipo, acepta multiples ComputerName')] 
    [String[]]$ComputerName
    )
    Begin {}
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
                    Write-Host 'Parando el servicio WinRM....'
                    (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StopService() | Out-Null
                    Start-Sleep -Seconds 7
                    if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Stopped') {Throw 'Ocurrió un fallo al parar el servicio WinRM'}
                    }
                    Catch 
                    {
                    Write-Warning $_.exception.message
                    Write-Warning "La función aboratará operaciones en el equipo $Computer"
                    break
                    }
                    Try 
                    {
                    Write-Host 'Arrancando el servicio WinRM....'
                    Write-Host
                    (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StartService() | Out-Null
                    Start-Sleep -Seconds 7
                    if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Running') {Throw 'Ocurrió un fallo al arrancar el servicio WinRM'}
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
    End {}
}