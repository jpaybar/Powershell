Import-Module -Name .\Remote_PSRemoting_2.0.1.psm1
function MenuPrincipal
{
     param (
           [string]$tituloMenu = 'OPCIONES'
     )
     Clear-Host
     Write-Host "################################################################################################" -ForegroundColor Yellow
     Write-Host "#                                                                                              #" -ForegroundColor Yellow                                                                                                                #"
     Write-Host "# AUTOR:          Juan M. Payan - https://alinuxaday.wordpress.com                             #" -ForegroundColor Yellow
     Write-Host "# EMAIL:          st4rt.fr0m.scr4tch@gmail.com                                                 #" -ForegroundColor Yellow
     Write-Host "# SCRIPT:         MyRemoteAdmin.ps1                                                            #" -ForegroundColor Yellow
     Write-Host "# DESCRIPCIÓN:    Obtiene información de un equipo remoto usando PowerShell, CIM y WMI.        #" -ForegroundColor Yellow
     Write-Host "# VERSIÓN:        1.0                                                                          #" -ForegroundColor Yellow
     Write-Host "# SINOPSIS:       MyRemoteAdmin es un script de Powershell principal que llama a otros modulos #" -ForegroundColor Yellow
     Write-Host "#                 y scripts para usar diversas funciones. Obtiene informacion general del      #" -ForegroundColor Yellow
     Write-Host "#                 equipo, comprueba si esta disponible en la red e instala impresoras Tcp/ip   #" -ForegroundColor Yellow
     Write-Host "#                 en equipos remotos entre otras utilidades.                                   #" -ForegroundColor Yellow
     Write-Host "#                 https://www.powershellgallery.com/                                           #" -ForegroundColor Yellow
     Write-Host "#                                                                                              #" -ForegroundColor Yellow
     Write-Host "################################################################################################" -ForegroundColor Yellow

     Write-Host "=========================================== $tituloMenu ===========================================" -ForegroundColor Cyan
     Write-Host "=                                                                                              ="
     Write-Host "= 1: Pulse '1' para verificar respuesta del equipo 'Ping Test'.                                ="
     Write-Host "= 2: Pulse '2' para obtener informacion del PC remoto (OS version, serial, etc...).            ="
     Write-Host "= 3: Pulse '3' para enviar un mensaje a uno o varios equipos en la red.                        ="
     Write-Host "= 4: Pulse '4' para instalar una impresora TCP/IP en un equipo remoto.                         ="
     Write-Host "= 5: Pulse '5' para habilitar el servicio WinRM (PowerShell Remota).                           ="
     Write-Host "= 6: Pulse '6' para habilitar una regla que permita el servicio WinRM en el Firewall           =" 
     Write-Host "=              de Windows.                                                                     ="
     Write-Host "= Q: Pulse 'Q' para salir del Script.                                                          ="
     Write-Host "=                                                                                              ="
     Write-Host "================================================================================================" -ForegroundColor Cyan
}
do
{
     MenuPrincipal
     Write-Host
     $opcion = Read-Host "Por favor, eliga una opción"
     switch ($opcion)
     {
             '1' {
                Clear-Host
                .\Test-MyPcConnection.ps1
           } '2' {
                Clear-Host
                . .\Get-MyDeviceInfo.ps1
                $ComputerName= Read-Host "Introduzca la IP o el nombre del equipo"
                $ComputerName | Get-MyDeviceInfo                  
           } '3' {
                Clear-Host
                .\Send-MyMessage.ps1
           } '4' {
                Clear-Host
                .\Add-MyPrinter.ps1
           } '5' {
                Clear-Host
                Write-Host "Introduzca la IP o multiples IPs donde quiere activar el servicio WinRM." -ForegroundColor DarkGray
                Write-Host
                Set-WINRMListener
                Write-Host
                Write-Host "Vuelva a introducir los datos anteriores para reiniciar el servicio y levantarlo." -ForegroundColor DarkGray
                Write-Host  
                Restart-WinRM  
                Write-Host
                Write-Host "Ahora ya puede conectar al equipo remoto usando Powershell. Abra una consola y ejecute" -ForegroundColor DarkGray
                Write-Host "el siguiente comando, Ejemplo:" -ForegroundColor DarkGray
                Write-Host
                Write-Host "Enter-PSSession -ComputerName NombreEquipo" -ForegroundColor DarkCyan   
                Write-Host
                Write-Host "El resultado del comando debe de ser similar a lo siguiente:" -ForegroundColor DarkGray
                Write-Host
                Write-Host "[NombreEquipo]: PS C:\Users\NombreUsuario\Documents>" -ForegroundColor DarkCyan
                Write-Host
                Write-Host "NOTA: Introduzca el Nombre de Equipo y no la IP, para ello tendria que habilitar" -ForegroundColor Yellow
                Write-Host "TrustedHosts list en el equipo y asi poder usar direcciones IP en vez de nombres." -ForegroundColor Yellow 
                Write-Host                                           
           } '6' {
                Clear-Host
                Clear-Host
                Write-Host "Introduzca la IP o multiples IPs donde quiere configurar el Firewall de Windows para permitir el servicio WinRM." -ForegroundColor DarkGray
                Write-Host
                Set-WinRMFirewallRule
                Write-Host
                Write-Host "Vuelva a introducir los datos anteriores para reiniciar el Firewall y aplicar la regla." -ForegroundColor DarkGray
                Write-Host  
                Restart-WindowsFirewall  
                Write-Host
           } 'q' {
                return
           }
     }
     pause
}
until ($opcion -eq 'q')