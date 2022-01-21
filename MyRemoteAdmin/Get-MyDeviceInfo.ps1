<#
.SYNOPSIS
Función que obtiene información de un equipo remoto por su nombre o IP.

.DESCRIPTION
Función que obtiene información de un equipo remoto por su nombre o IP. Devulve información del nonitor
o monitores instalados, fabricante, modelo y numero de serie. En cuanto al equipo devuelve el fabricante
modelo, numero de serie, nombre del equipo, dominio al que pertenece, versión del sistema operativo y
arquitectura (32/64 bits).

.PARAMETER ComputerName
Si no se introduce ningún parametro la función solicita tantos equipos como queramos
introducir por pantalla (Mandatory=$True).

.EXAMPLE
Get-MyDeviceInfo
Get-MyDeviceInfo equipo1, equipo2, equipo3, ................
Get-Content .\equipos.txt | Get-MyDeviceInfo
#>

Function Get-MyDeviceInfo
{
    [CmdletBinding()]
    Param
    (
        [Parameter( 
            ValueFromPipeline=$True, 
            ValueFromPipelineByPropertyName=$True,
            Mandatory=$True,
            HelpMessage='Introduzca el nombre del equipo o la IP, acepta multiples equipos para la consulta.')] 
        [String[]]$ComputerName
    )
 
    Begin {
        $pipelineInput = -not $PSBoundParameters.ContainsKey('ComputerName')
    }
 
    Process
    {
        Clear-Host
        Function DoWork([string]$ComputerName) {
            $ActiveMonitors = Get-WmiObject -Namespace root\wmi -Class wmiMonitorID -ComputerName $ComputerName
            $monitorInfo = @()
 
            foreach ($monitor in $ActiveMonitors)
            {
                $mon = $null
                $mon = New-Object PSObject -Property @{
                "Año de Fabricacion"=$monitor.YearOfManufacture
                "Semana de Fabricacion"=$monitor.WeekOfManufacture
                "Numero de Serie"=($monitor.SerialNumberID | % {[char]$_}) -join ''
                "Fabricante"=($monitor.ManufacturerName | % {[char]$_}) -join ''
                "Modelo"=($monitor.UserFriendlyName | % {[char]$_}) -join ''
                }
                $monitorInfo += $mon
            }
            $cmpserial=$(get-wmiobject win32_bios -ComputerName $ComputerName | select-object serialnumber).serialnumber
            $NombreEquipo=$(Get-WmiObject win32_computersystem -ComputerName $ComputerName | Select-Object Name).Name
            $Dominio=$(Get-WmiObject win32_computersystem -ComputerName $ComputerName | Select-Object Domain).Domain
            $VersionWindows=$(Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName | Select-Object Caption).Caption
            $Arquitectura=$(Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName | Select-Object OSArchitecture).OSArchitecture
            $fabricante=$(Get-WmiObject win32_computersystem -ComputerName $ComputerName | Select-Object Manufacturer).Manufacturer
            $Modelo=$(Get-WmiObject win32_computersystem -ComputerName $ComputerName | Select-Object model).model
            Write-Host
            Write-Host
            Write-Host "###########################################################" -ForegroundColor DarkRed
            Write-Host
            Write-Host "DATOS DEL MONITOR/ES" -foregroundColor Gray -BackgroundColor DarkCyan
            Write-Host "====================" -ForegroundColor DarkCyan
            Write-Output $monitorInfo 
            Write-Host
            Write-Host "DATOS DEL EQUIPO" -foregroundColor Gray -BackgroundColor DarkCyan
            Write-Host "================" -ForegroundColor DarkCyan
            Write-Host
            Write-Host
            Write-Host "Nombre del Equipo   : $NombreEquipo" 
            Write-Host "Nombre del Dominio  : $Dominio"
            Write-Host "Version SO          : $VersionWindows"
            Write-Host "Arquitectura bits   : $Arquitectura"
            Write-Host "Fabricante          : $fabricante"
            Write-Host "Modelo              : $Modelo"
            write-host "Numero de Serie     : $cmpserial"
            Write-Host
            Write-Host
            Write-Host "###########################################################" -ForegroundColor DarkRed
            
        }
 
        if ($pipelineInput) {
            DoWork($ComputerName)
        } else {
            foreach ($item in $ComputerName) {
                DoWork($item)
            }
        }
    }
}

