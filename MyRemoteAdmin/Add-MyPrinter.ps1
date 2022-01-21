#######################################################################
####  INSTALA IMPRESORA TCP/IP EN EQUIPO REMOTO
#######################################################################
####   Este script instala una impresora en una computadora remota
####   por el puerto TCP/IP. Para ello hace uso de las "Printing
####   Admin Tools de Microsoft Windows" (scripts de Visual Basic) y
####   el uso de la libreria "PrintUI.dll" para la gestion de drivers
####   e impresoras.
#######################################################################
####  
#######################################################################
#######################################################################
####
#######################################################################
####  FUNCIONES  ######################################################


####################### FUNCION LISTAR IMPRESORAS #####################
FUNCTION ShowPrinterList($Computer, $Printer)
{
	Clear-Host
    Write-Host
    Write-Host
	Write-Host ("Obteniendo el listado de impresoras en " + $Computer + "...") -foregroundColor Red -BackgroundColor Cyan
	& cscript .\Printing_Admin_Scripts\prnmngr.vbs -l -s $Computer > C:\windows\temp\prnmngr.log
	$PrinterList = Get-Content C:\windows\temp\prnmngr.log | Where-Object {$_ -match "Nombre de impresora"} 
	Remove-Item C:\windows\temp\prnmngr.log
	if ($Null -eq $PrinterList)
	{
		Write-Host
		Write-Host
		Write-Host ("No hay impresoras instaladas en el equipo " + $Computer)
		Write-Host
    	Write-Host
		return
	}
	For ($i = 0;$i -lt $PrinterList.Count;$i++)
		{$PrinterList[$i] = $PrinterList[$i].Substring(0)}
	
	Write-Host
    Write-Host
	Write-Host ("Impresoras instaladas en " + $Computer) -foregroundColor $InfoMessage
    Write-Host "=====================================" -foregroundColor Cyan
	ForEach ($Item in $PrinterList)
	{
		if (($Printer) -and ($Item -match $Printer))
		{
			Write-Host $Item -noNewLine
			Write-Host " <-- Impresora instalada" -ForegroundColor $InfoMessage
		}
		else
			{Write-Host $Item}
	}
	Write-Host
    Write-Host
    Write-Host
}

################### FUNCION LISTAR DRIVERS #######################

FUNCTION ShowDriverList($Computer, $Driver)
{
	Clear-Host
    Write-Host
    Write-Host
	Write-Host ("Obteniendo el listado de drivers en " + $Computer + "...") -foregroundColor Red -BackgroundColor Cyan
	& cscript .\Printing_Admin_Scripts\prndrvr.vbs -l -s $Computer > C:\windows\temp\prndrvr.log
	$DriverList = Get-Content C:\windows\temp\prndrvr.log | Where-Object {$_ -match "Nombre de controlador"} 
	Remove-Item C:\windows\temp\prndrvr.log
	if ($Null -eq $DriverList)
	{
		Write-Host ("No hay drivers instalados en el equipo " + $Computer)
		return
	}
	For ($i = 0;$i -lt $DriverList.Count;$i++)
		{$DriverList[$i] = $DriverList[$i].Substring(0)}
	
	Write-Host
	Write-Host ("Drivers instalados en " + $Computer) -foregroundColor $InfoMessage
	Write-Host "==================================" -foregroundColor Cyan
	ForEach ($Item in $DriverList)
	{
		if (($Driver) -and ($Item -match $Driver))
		{
			Write-Host $Item -noNewLine
			Write-Host " <-- Driver instalado" -foregroundColor $InfoMessage
		}
		else
			{Write-Host $Item}
	}
	Write-Host
	Write-Host
    Write-Host
}

##################### FUNCION LISTAR PUERTOS ############################
FUNCTION ShowPortList($Computer, $Port)
{
	Clear-Host
    Write-Host
    Write-Host
	Write-Host ("Obteniendo el listado de puertos TCP/IP en " + $Computer + "...") -foregroundColor Red -BackgroundColor Cyan
	& cscript .\Printing_Admin_Scripts\prnport.vbs -l -s $Computer > C:\windows\temp\prnport.log
	$PortList = Get-Content C:\windows\temp\prnport.log | Where-Object {$_ -match "Nombre del puerto"} 
	Remove-Item C:\windows\temp\prnport.log
	if ($Null -eq $PortList)
	{
		Write-Host ("No hay puertos TCP/IP instalados en " + $Computer)
		return
	}
	
	Write-Host
	Write-Host ("Puertos TCP/IP instalados en " + $Computer) -foregroundColor $InfoMessage
	Write-Host "=========================================" -foregroundColor Cyan
	ForEach ($Item in $PortList)
	{
		if (($Port) -and ($Item -match $Port))
		{
			Write-Host $Item -noNewLine
			Write-Host " <-- Puerto creado" -foregroundColor $InfoMessage
		}
		else
			{Write-Host $Item}
	}
	Write-Host
	Write-Host
    Write-Host
}

##################### FUNCION MOSTRAR MENU MARCA ####################

function Show-PrinterMenu
{
     param (
           [string]$menuImpresoras = 'SELECCIONE LA MARCA DE LA IMPRESORA'
     )
     Clear-Host
     Write-Host "============$menuImpresoras==================" -ForegroundColor Gray
     Write-Host "                                                                 " -BackgroundColor DarkCyan
     Write-Host "1: Pulse '1' para instalar una impresora Canon Color.            " -BackgroundColor DarkRed -ForegroundColor Gray
     Write-Host "2: Pulse '2' para instalar una impresora Canon B/N.              " -BackgroundColor DarkRed -ForegroundColor Gray
     Write-Host "3: Pulse '3' para instalar una impresora HP.                     " -BackgroundColor DarkRed -ForegroundColor Gray
     Write-Host "4: Pulse '4' para instalar una impresora Brother.                " -BackgroundColor DarkRed -ForegroundColor Gray
     Write-Host "Q: Pulse 'Q' para salir.                                         " -BackgroundColor DarkRed -ForegroundColor Gray                                         
     Write-Host "                                                                 " -BackgroundColor DarkCyan
     Write-Host "=================================================================" -ForegroundColor Gray
}


Do{
    Show-PrinterMenu
    Write-Host
    Write-Host "Por favor, eliga una opcion del menu:" -ForegroundColor Gray
    $opcion = Read-Host
    Write-Host
    Write-Host
    switch ($opcion)
        {
            1 {
                $DriverName="Canon UFR II Color Class Driver"
                $DriverPath=".\Canon_Driver"
                $DriverFile="prncacl2.inf"
                $DriverFullPath = Join-Path $DriverPath $DriverFile
            } 
            2 {
                $DriverName="Canon UFR II B/W Class Driver"
                $DriverPath=".\Canon_Driver"
                $DriverFile="prncacl2.inf"
                $DriverFullPath = Join-Path $DriverPath $DriverFile
            } 
            3 {
                $DriverName="HP Universal Printing PCL 6"
                $DriverPath=".\HP_Driver"
                $DriverFile="hpcu250u.inf"
                $DriverFullPath = Join-Path $DriverPath $DriverFile
            } 
            4 {
                $DriverName="Brother PCL5e Driver"
                $DriverPath=".\Brother_Driver"
                $DriverFile="bhpcl5e.inf"
                $DriverFullPath = Join-Path $DriverPath $DriverFile
            }
            Q {return}
        }
}
until ($opcion -eq 'q' -or $opcion -eq '1' -or $opcion -eq '2' -or $opcion -eq '3' -or $opcion -eq '4')

#####################################################################


#####################################################################
####   INICIALIZAR VARIABLES PARA LOS MENSAJES
#####################################################################
$InfoMessage = $Host.PrivateData.WarningForeGroundColor
$ErrorMessage = $Host.PrivateData.ErrorForeGroundColor
#####################################################################

#####################################################################
####   SCRIPT PRINCIPAL 
#####################################################################
Clear-Host
Write-Host
$ComputerName = Read-Host "Introduzca el nombre o la IP del equipo donde quiere instalar la impresora"
if ($ComputerName -eq "")
	{return}
else
{	
	$Filter = "Address=" + [char]34 + $ComputerName + [char]34
	$PingStatus = Get-WmiObject -Query "SELECT * FROM Win32_PingStatus WHERE $Filter"
	if (($Null -eq $PingStatus) -or ($PingStatus.StatusCode -ne 0))
	{
		Write-Host
		Write-Host
		Write-Host ($ComputerName + " No responde a Ping, puede que el equipo este apagado.") -ForegroundColor $ErrorMessage
		Write-Host
		Write-Host "Revise las reglas del Firewall para peticiones ICMP echo." -ForegroundColor $ErrorMessage
		Write-Host
		Write-Host
		return
	}
}

ShowPrinterList -computer $ComputerName
Write-Host	
Write-Host
$PrinterName = Read-Host "Introduzca un nombre para la impresora"
	
$PortIP = Read-Host "Introduzca la direccion IP de la impresora"
$PortName = "IP_" + $PortIP
& cscript .\Printing_Admin_Scripts\prnport.vbs -a -r $PortName -h $PortIP -s $ComputerName -o RAW -n 9100
ShowPortList -computer $ComputerName -port $PortName
Write-Host
Write-Host
	
$CommandLine = '@Start /WAIT RunDLL32 PrintUI.dll PrintUIEntry /ia /c' + [char]92+[char]92 + $ComputerName
$CommandLine = $CommandLine + ' /m "' + $DriverName + '" /h "x64" /v "Type 3 – User Mode" /f "' + $DriverFullPath + '"'
Set-Content C:\Windows\Temp\prndrvr.bat $CommandLine
$CommandLine = '@Start /WAIT .\sc ' + [char]92+[char]92 + $ComputerName + " stop spooler"
Add-Content C:\Windows\Temp\prndrvr.bat $CommandLine
$CommandLine = '@Start /WAIT .\sc ' + [char]92+[char]92 + $ComputerName + " start spooler"
Add-Content C:\Windows\Temp\prndrvr.bat $CommandLine

& C:\Windows\Temp\prndrvr.bat
Remove-Item C:\Windows\Temp\prndrvr.bat
& cscript .\Printing_Admin_Scripts\prnmngr.vbs -a -s $ComputerName -m $DriverName -r $PortName -p $PrinterName
ShowPrinterList -computer $ComputerName -printer $PrinterName

return







