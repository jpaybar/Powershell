<# 
Script para probar la conectividad de un equipo, si el SO de destino es Windows 10, 
el Firewall puede estar filtrando las peticiones echo ICMP y no dar respuesta.
Se pueden introducir multiples equipos separados por ","...Por ejemplo: PC1, PC2, PC3.
#>
Clear-Host
$Lista = Read-Host "Introduzca la direccion IP o direcciones IP separadas por coma (Ejemplo - IP1,IP2,IP3)"

$IPs = $Lista.Split(",").Trim()

foreach ( $IP in $IPs )
    {
    if ( Test-Connection -ComputerName $IP -Count 1 -Quiet )
        {
        Write-Host
        Write-Host "$IP responde a Ping." -ForegroundColor Green
        }
    else
        {
        Write-Host
        Write-Host "$IP NO responde a Ping." -ForegroundColor red
        }
    }
Write-Host