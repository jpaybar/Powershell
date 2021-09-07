<#
.SYNOPSIS
	Repair-MyPcDomainTrust repara la relación de confianza entre un equipo y un dominio de AD.
.DESCRIPTION
	Repair-MyPcDomainTrust repara la relación de confianza entre un equipo y un dominio de AD.
	Solicita la ruta a 2 ficheros XML con las credenciales Locales y de Dominio desde los que
	se importaran usuario/contraseña.
.PARAMETER Computername
	Nombre del equipo que se quiere volver a unir al dominio.
.PARAMETER DomainName
	Nombre del dominio al que queremos unir de nuevo el equipo.
.PARAMETER UnjoinLocalCredentialXmlFilePath
	Repair-MyPcDomainTrust usa un fichero XML con las credenciales de un administrador Local/Dominio para todo
	el proceso de quitar y agregar la computadora al dominio, de esta forma no tenemos que introducir las credenciales
	del usuario Local/Dominio cada vez que se ejecute el script.
	Este parámetro solicita la ruta al fichero XML que contiene las credenciales del administrador local.
.PARAMETER DomainCredentialXmlFilePath
	Repair-MyPcDomainTrust usa un fichero XML con las credenciales de un administrador Local/Dominio para todo
	el proceso de quitar y agregar la computadora al dominio, de esta forma no tenemos que introducir las credenciales
	del usuario Local/Dominio cada vez que se ejecute el script.
	Este parámetro solicita la ruta al fichero XML que contiene las credenciales del administrador de Dominio.
.EXAMPLE
	.\Repair-MyPcDomainTrust.ps1
		ó
	Repair-MyPcDomainTrust -Computername NombrePC -DomainName sub.dominio.com -UnjoinLocalCredentialXmlFilePath .\LocalCred.xml -DomainCredentialXmlFilePath .\DomainCred.xml
.NOTES
	Credenciales:
	$Credenciales = Get-Credential  ## Introducimos usuario=(dominio\usuariodedominio) y contraseña
	$Credenciales | Export-Clixml -Path .\DomainCred.xml
#>
function Repair-MyPcDomainTrust {

	[CmdletBinding()]
	param (
		[Parameter(Mandatory,
				ValueFromPipeline,
				ValueFromPipelineByPropertyName)]
		[string[]]$Computername,
		[Parameter(Mandatory)]
		[string]$DomainName,
		[Parameter(Mandatory)]
		[string]$UnjoinLocalCredentialXmlFilePath,
		[Parameter(Mandatory)]
		[string]$DomainCredentialXmlFilePath
	)

	Begin {

		$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop	# Establecemos la variable para capturar las posibles excepciones.

		function Import-Credentials {
			<#
			.SYNOPSIS
			Esta función crea un objeto de tipo "PSCredential" importando un fichero XML.
			.PARAMETER XmlFilePath
			Ruta al fichero XML que contiene las credenciales.
			#>
			[CmdletBinding()]
			[OutputType([System.Management.Automation.PSCredential])]	# Especificamos el tipo de objeto que devolverá la función, en este caso se genera.
			param (														# el usuario y la contraseña al importar el fichero XML con las credenciales que después
				[Parameter(Mandatory)]									# se usan para crear el objeto de tipo "PSCredential".
				# Si no existe el fichero de credenciales, no ejecutamos la función.									
				[ValidateScript({ Test-Path -Path $_ -PathType Leaf })]		
				[string]$XmlFilePath
			)
			process {
				$Cred = Import-Clixml $XmlFilePath
				New-Object System.Management.Automation.PSCredential($Cred.username, $Cred.password)	# Creamos un nuevo objeto de credenciales con el fichero importado.
			}
		}
		
		function Test-Reboot ($Computername,$Credential) {
			while (Test-Connection -ComputerName $Computername -Count 1 -Quiet) {	# Controlamos la ejecución del script y lo paramos durante 5 segundos si el equipo
				Write-Host "Esperando a que el equipo $Computername se reinicie..." # responde a Ping.
				Start-Sleep -Seconds 5
			}
			Write-Host
			Write-Host "El equipo $Computername está fuera de linea. A la espera de respuesta Ping..."
			Write-Host
			while (!(Test-Connection -ComputerName $Computername -Count 1 -Quiet)) {	# Volvemos a parar el script mientras que el equipo no responda a Ping.
				Start-Sleep -Seconds 5
				Write-Host "Esperando que el equipo $Computername responda a Ping"
			}
			Write-Host
			Write-Host "El equipo $Computername vuelve a estar Online. Esperando que el Sistema inicie..."
			Write-Host
			$EapBefore = $ErrorActionPreference		# Volcamos el valor de la varible "$ErrorActionPreference" definida en el bloque Begin a "stop" para poder cambiar el 
			$ErrorActionPreference = [System.Management.Automation.ActionPreference]::SilentlyContinue # valor a la llamada de esta función  por "SilentlyContinue".
			while (!(Get-WmiObject -ComputerName $Computername -Class Win32_OperatingSystem -Credential $Credential)) {
				Start-Sleep -Seconds 5								# Volvemos a parar el script hasta que el sistema no inicie y nos permita hacer login con las credenciales
				Write-Host "Esperando a que el Sistema inicie..."	# (locales o de dominio), ya que la función se llama una vez que el equipo está fuera de dominio 
			}														# y cuando se vuelve a meter.
			$ErrorActionPreference = $EapBefore		# Devolvemos a la variable "$ErrorActionPreference" su valor original "Stop".
		}
	}
	Process {
		foreach ($Computer in $Computername) {
			try {
				if (Test-Connection -ComputerName $computer -Count 1 -Quiet) { 
					Write-Host
					Write-Host
					Write-Host "El equipo '$Computer' está Online" -ForegroundColor Yellow
					Write-Host
					## Se importan los ficheros XML con la función "Import-Credentials" definida en el bloque Begin y guardamos lo que nos devuelve la función
					## en sus correspondientes variables (local y de dominio), que pasaremos a los comandos "Remove-Computer" y "Add-Computer".
					$LocalCredential = (Import-Credentials -XmlFilePath $UnjoinLocalCredentialXmlFilePath)
					$DomainCredential = (Import-Credentials -XmlFilePath $DomainCredentialXmlFilePath)
					Write-Host "Sacando equipo de Dominio, forzando reinicio..."
					Write-Host
					Remove-Computer -ComputerName $Computer -LocalCredential $LocalCredential -UnjoinDomainCredential $DomainCredential -Workgroup Workgroup -Restart -Force
					Write-Host "El equipo ha salido del Dominio. Esperando reinicio." -ForegroundColor Yellow
					Write-Host
					Test-Reboot -Computername $Computer -Credential $LocalCredential				# Llamamos a la función "Test-Reboot" para parar el script después del
					Write-Host "El equipo $Computer se ha reiniciado. Uniendo equipo al Dominio."	# reinicio de sacar el equipo de dominio también se verifica que nos 
					Write-Host																		# permite login con credenciales locales.
					Add-Computer -ComputerName $Computer -DomainName $DomainName -Credential $DomainCredential -LocalCredential $LocalCredential -Restart -Force
					Write-Host "El equipo $Computer se ha unido al Dominio. Esperando reinicio final" -ForegroundColor Yellow
					Write-Host
					Test-Reboot -Computername $Computer -Credential $DomainCredential 				# Llamamos de nuevo a la función "Test-Reboot" pero esta vez después de agregar
					Write-Host "El equipo $Computer se unió al Dominio $DomainName correctamente" -ForegroundColor Yellow # el equipo a dominio y probar las credenciales de usuario
					Write-Host																							  # de dominio.
					[pscustomobject]@{ 'Equipo' = $Computer; 'Resultado' = $true }	# Creamos un objeto para mostrar que el resultado del script ha sido correcto.
				} else {
					throw "El equipo '$Computer' está apagado o el firewall no permite peticions ICMP echo"	
				}
			} catch {
				# Creamos un objeto para mostrar que el script no se ha ejecutado correctamente y mostramos el mensaje de error.
				[pscustomobject]@{ 'Equipo' = $Computer; 'Resultado' = $false; 'Error' = $_.Exception.Message }
			}
		}
	}
	End {}
}

Repair-MyPcDomainTrust