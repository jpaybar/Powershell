<#
.SYNOPSIS
Get-MyWindowsSerialFromRegistry obtiene el número de serie de Windows, lo muestra por pantalla
y lo vuelca a un fichero TXT.
.DESCRIPTION
Get-MyWindowsSerialFromRegistry obtiene el número de serie de Windows,lo muestra por pantalla
y lo vuelca a un fichero TXT. Para ello, levanta el servicio "RemoteRegistry" en el equipo remoto, 
lee el valor de la clave "BackupProductKeyDefault" y la muestra junto con el nombre del equipo
y la versión del Sistema Operativo.
.PARAMETER file
Ruta al fichero que contiene los equipos de los cuales queremos obtener el S/N del sistema.
.EXAMPLE
. .\Get-MyWindowsSerialFromRegistry.ps1         Llamamos al script con la notación del punto para cargar la función.

Ejecutamos la función:

Get-MyWindowsSerialFromRegistry

        ó

Get-MyWindowsSerialFromRegistry -file .\equipos.txt
.NOTES
Necesita que se facilite un fichero con el nombre de máquina o IP de los equipos a examinar, 
las entradas de dicho fichero serán un equipo por fila. El fichero TXT generado a la salida
no tiene porque existir, se genera automaticamente.
#>
function Get-MyWindowsSerialFromRegistry {

        [CmdletBinding()]
        Param(
        [parameter(mandatory=$true, ValueFromPipeline)]
        [string]$file
        )
        
        begin {
                $ErrorActionPreference = "Stop"
                $key = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
                $valuename = 'BackupProductKeyDefault'
                $computers = Get-Content $file

                ## OTRA FORMA DE CREAR EL OBJETO Y OBTENER EL VALOR
                ##
                ## $reg = Get-WmiObject -List "StdRegprov" -ComputerName $computername #-Credential $Credential
                ## $HKLM = "2147483650"
                ## $key = 'SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform'
                ## $valuename = 'BackupProductKeyDefault'
                ## $data = ($reg.GetExpandedStringValue($HKLM, $key, $valuename)).svalue
        }       ## $data        OBTENEMOS VALOR EN "$data"

        
        process {
                Clear-host
                try {
                        foreach ($computer in $computers) {
                                if (Test-Connection -ComputerName $Computer -Count 1 -Quiet){
                                        $VersionWindows=$(Get-WmiObject Win32_OperatingSystem -ComputerName $Computer | Select-Object Caption).Caption
                                        Set-Service -ComputerName $computer -Name remoteregistry -StartupType Manual
                                        Set-Service -ComputerName $computer -Name remoteregistry -Status Running
                                        $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $computer)
                                        $regkey = $reg.opensubkey($key)
                                        $SerialNumber = $regkey.getvalue($valuename) 
                                        "$computer      " + "$VersionWindows    " + $SerialNumber >> seriales.txt
                                        Write-Host
                                        [pscustomobject]@{ 'Equipo' = $Computer; 'Sistema' = $VersionWindows; 'Serial' = $SerialNumber } 
                                        Set-Service -ComputerName $computer -Name remoteregistry -StartupType Disabled
                                        Get-Service -Computer $computer -Name remoteregistry | Stop-Service -Force
                                }
                                else {
                                        Write-Host
                                        Write-Warning "El equipo $computer no responde a Ping, está apagado o el Firewall bloquea las petciones ICMP echo."
                                        "$computer       no responde a Ping, está apagado o el Firewall bloquea las petciones ICMP echo." >> seriales.txt
                                        Write-Host
                                }
                        }   
                }
                catch {
                       Write-Warning $_.exception
                       Write-Warning "Se ha producido un error, la función no se ejecutará en $computer"
                }
                
        }
        
        end {
                
        }
}

#Get-MyWindowsSerialFromRegistry

<################################## CONSULTAR CLAVES DEL REGISTRO #####################################################################################################

Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\'

Get-Item 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\'  

Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\'

Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\'

Get-ItemPropertyValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\' -Name BackupProductKeyDefault | Out-File CLAVE_WINDOWS.txt

################################# MODIFICAR EL ESTADO DE LOS SERVICIOS #################################################################################################

Set-Service -ComputerName NombrePC -Name remoteregistry -StartupType Manual

Set-Service -ComputerName NombrePC -Name remoteregistry -StartupType Disabled

Set-Service -ComputerName NombrePC -Name remoteregistry -Status Running

Get-Service -Computer NombrePC -Name remoteregistry | Stop-Service -Force

########################################################################################################################################################################>