<#
.SYNOPSIS
Set-MyMultiEnter-PSSession, levanta el servicio "WinRM" y crea conexiones interactivas de Powershell 
en multiples equipos remotos.
.DESCRIPTION
Set-MyMultiEnter-PSSession, levanta el servicio "WinRM" y crea conexiones interactivas de Powershell 
en multiples equipos remotos. Si el sistema es "Windows 7" hace uso del ejecutable "PsExec.exe" 
de las herramientas PSTools y si es "Windows 10" arranca el servicio haciendo uso del CmdLet "start-service".
.PARAMETER ComputerName
Equipo/Equipos remotos en los que se abrirá una Powershell interactiva.
.EXAMPLE
Set-MyMultiEnter-PSSession 
        ó
Set-MyMultiEnter-PSSession -ComputerName NombreEquipo
        ó
Get-Content .\equipos.txt | Set-MyMultiEnter-PSSession
.NOTES
La función "Set-MyMultiEnter-PSSession" necesita el ejecutable "PsExec.exe" de las herramientas PSTools 
si el equipo que queremos administrar tiene instalado "Windows 7".
#>
function Set-MyMultiEnter-PSSession {
        
        [CmdletBinding()]
        param(
	[Parameter(Mandatory)]
	[ValidateNotNullOrEmpty()]
	[String[]]$ComputerName
        )
        
        begin {
                $ErrorActionPreference = "SilentlyContinue"
                $VersionWindows=$(Get-WmiObject Win32_OperatingSystem -ComputerName $computer | Select-Object Caption).Caption
        }
        process {
                Clear-Host
                foreach($computer in $ComputerName){
                        if(Test-Connection -ComputerName $computer -Count 1 -Quiet){
                                if ($VersionWindows -like "*7*") {
                                        .\PsExec.exe \\$computer -s -i powershell enable-PSRemoting -Force
                                }
                                else {
                                        Get-Service -ComputerName $computer -name winrm | Start-Service
                                }
                        
                                $cmds = "-Noexit","-command enter-pssession -computername $computer"
                                Start-Process powershell.exe -ArgumentList $cmds
                        }
                        else{ 
                                Write-Host
                                Write-Host "El equipo $computer no responde a Ping" -ForegroundColor Red
                                Write-Host
                        }   
                }     
        }
        end {
                
        }
}





