<#
.SYNOPSIS
Instala software en un equipo remoto de forma silenciosa.

.DESCRIPTION
"Set-MyAppDeployment" instala software en un equipo remoto, copiando primero el ejecutable en "C:\" y realizando una instalación
transparente para el usuario sin necesidad de reiniciar el equipo.

.PARAMETER ComputerName
Nombre del equipo/equipos en los que se realizará la instalación de la aplicación.

.PARAMETER InstallerFilePath
Ruta donde se encuentra la aplicación (instalador/ejecutable).

.EXAMPLE
Set-MyAppDeployment -ComputerName NombrePCremoto -InstallerFilePath "C:\ruta\instalador.exe"

ó

Set-MyAppDeployment

.NOTES
Necesita que el servicio "winrm" se esté ejecutando en el equipo remoto, una cuenta con privilegios de administración
y acceso al recurso administrativo "c$".

NOTA: Para verificar el estado de "winrm" y arrancarlo si es necesario ejecute los siguientes comandos:

        Get-Service -ComputerName NombrePCremoto | Where-Object {$_ -like "winrm"}

        Get-Service -ComputerName NombrePCremoto | Where-Object {$_ -like "winrm"} | Start-Service
#>
function Set-MyAppDeployment
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
		#[ValidateScript({ Test-Connection -ComputerName $_ -Quiet -Count 1 })] #Si no hay respuesta a ping desde el equipo el script no se ejecuta.
		[string[]]$ComputerName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerFilePath
    )
    begin
	{
		$ErrorActionPreference = 'Stop' #Se establece el valor de la variable a "Stop" en vez de "Continue" para realizar la captura de excepciones en el Try/Catch/Finally
	}
    process
    {
        foreach ($Computer in $ComputerName) {
            try {
                #Chequeamos el estado del equipo/s remoto, si responde a ping y si el recurso compartido c$ está operativo.
                $ClientMsg = "El equipo remoto [$($Computer)] "
                if (-not (Test-Connection -ComputerName $Computer -Quiet -Count 1)) {
                    throw "$ClientMsg no responde a Ping"
                } elseif (-not (Test-Path "\\$Computer\c$")) {
                    throw "$ClientMsg no tiene disponible el recurso administrativo c$"
                } elseif ("WinRM" | Get-Service -ComputerName $Computer | Where-Object {$_.Status -like "Stopped"}) {
                    throw "$ClientMsg no está ejecutando el servicio 'WinRM'"
                } else {
                    clear-Host
                    write-Host
                    Write-Host "$ClientMsg está preparado para la instalación"
                }
    
                #Se copia el instalador al equipo remoto
                Write-Host
                Write-Host "Copiando instalador en equipo remoto..."
                Copy-Item -Path $InstallerFilePath -Destination "\\$Computer\c$"    
    
                #Invocamos el comando de instalación en el cliente remoto con "Invoke-command"
                Write-Host
                Write-Host "Iniciando la instalación de la aplicación..."
                Invoke-Command -ComputerName $Computer -ScriptBlock {                   #La variable "$using" referenciada con el comando "Invoke-Command" o "Start-Job", 
                    $installerFileName = $using:InstallerFilePath | Split-Path -Leaf    #permite usar una variable existente en la sesión actual de la consola "$InstallerFilePath"  
                                                                                        #enviandola a la sesión de la consola remota ("$installerFileName" será el nombre de la nueva variable)
                                                                                        #dentro del "scriptblock" que se ejecutará en el equipo remoto. El uso es el siguiente:
                                                                                        #Asignamos a la variable del entorno remoto el valor de la variable local con "$using:'Variable_local'"
                                                                                        # $installerFileName = $using:InstallerFilePath

                                                                                        #"Split-Path" devuelve la parte especificada de la ruta, el parametro "-Leaf" devuelve el último item
                                                                                        #de la ruta, en este caso el ejecutable.

                    #Se ejecuta el instalador de forma transparente para el usuario del equipo remoto.
                    Start-Process -NoNewWindow -Wait -FilePath "C:\$installerFileName" -ArgumentList '/silent /norestart'
                    
                    #Se elimina el directorio de instalación en el equipo remoto.
                    Remove-Item -Path "C:\$installerFileName" -Recurse
                }
            } 
            catch {
                Write-Host
                Write-Warning "Se ha producido un error ejecutando el script"
                Write-Host
                #Write-Error $_
                Write-Warning $_
                Write-Host
            }
            Finally{
                Write-Host
                Write-Output "Fin del script"
                Write-Host
            }
        }
    }
    end
    {}
}