<# 
Script para enviar mensajes a equipos utilizando el comando msg.exe
El script solicita la siguiente informacion por pantalla.
Message :-  "Escribir el mensaje que se quiere enviar al equipo/s". 
Computer Name :- "Introducir el nombre(s) del equipo o la IP(s) al que se quiere enviar el mensaje"
Se pueden introducir multiples equipos separados por ","...Por ejemplo: PC1, PC2, PC3.
O introducir la ruta al fichero que contenga multiples equipos, un por fila.
De la siguiente forma:

PC1
PC2
PC3
 
Time :- "Introducir el tiempo en segundos que durara el mensaje en pantalla"
#>
 
# Declaracion de variables
Clear-Host
$Start_Time       =       Get-Date -Format T
$logFile          =       ‘Equipos_Sin_Ping.txt’
$Message          =       Read-Host -Prompt “Escriba el mensaje que quiere enviar” 
Write-Host
Write-Host "NOTA:" -ForegroundColor DarkCyan
Write-Host "=====" -ForegroundColor DarkCyan
Write-Host "Puede introducir el nombre o la IP del equipo, tambien multiples equipos separados por coma" -ForegroundColor DarkGray
Write-Host "Por ejemplo: PC1,PC2,PC3....etc o Introducir la ruta a un fichero que contenga los equipos" -ForegroundColor DarkGray
Write-Host "Por ejemplo C:\Equipos.txt , el formato debe ser un equipo por fila, de la siguiente forma:" -ForegroundColor DarkGray 
Write-Host
Write-Host "PC1" -ForegroundColor DarkGray
Write-Host "PC2" -ForegroundColor DarkGray
Write-Host "PC3" -ForegroundColor DarkGray
Write-Host
$ComputerName     =       Read-Host -Prompt “Escriba el nombre(s) o la IP(s) del equipo”  
Write-Host  
#$Time             =       Read-Host -Prompt “Introduzca el tiempo en segundos” 
Write-Host
$Session          =       “*”
$ComputerName     =       $ComputerName.Split(",").Trim() #$ComputerName -split ‘,’

Clear-Host
 
if ($ComputerName -match ':')
 
                      {
                      $Path = $ComputerName
                      $ComputerName = Get-Content $path
          }
                      $Total = $ComputerName.count 
                                foreach ($Computer in $ComputerName )
                                                {
                                                             if (Test-Connection -ComputerName $Computer -Count 1 -ErrorAction 0)
                                {
                                                                Write-Host “Enviando mensaje al equipo $Computer…….” -ForegroundColor DarkGray
                                #.\msg.exe $Session /Server:$Computer /Time:$Time $Message
                                .\msg.exe $Session /Server:$Computer $Message
                                                                Write-Host “Mensaje enviado correctamente a $Computer” -ForegroundColor Green
                                                                }
                                                                else
                                                                                {
 
                                                                Out-File -FilePath $logFile -InputObject $Computer -Append -Force
 
                                                                                                Write-Host “$Computer no responde a Ping…” -ForegroundColor Red
 
                                                                                }
 
                                                }
 
        $Not_Reachable_Count  = @(Get-Content $logFile).count
        $End_Time   =    Get-Date -Format T
        $Minute = (New-TimeSpan -Start $Start_Time -End $End_Time).Minutes
        $Second = (New-TimeSpan -Start $Start_Time -End $End_Time).Seconds
        Write-Host
        Write-Host "Inicio a las $Start_Time, Fin a las $End_Time, Realizado en $Minute Minutos $Second Segundos"
        Write-Host
        Write-Host “$Total equipos procesados, $Not_Reachable_Count equipos no respondieron. Resultado guardado en $logFile” -ForegroundColor white
        Write-Host
