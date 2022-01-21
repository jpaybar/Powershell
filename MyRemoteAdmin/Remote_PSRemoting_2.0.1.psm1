Function Set-WINRMListener
{
<#
  .SYNOPSIS
    Set-WinRMListener will configure a remote computer's WinRM service to listen for WinRM requests
 
  .DESCRIPTION
    Set-WinRMListener uses the system.management.managementclass object to create 3 registry keys. These registry entries will configure the WinRM service when it restarts next. Use Restart-WinRM to restart the WinRM service after it has been successfully configured.
 
  .EXAMPLE
    Get-Content .\servers.txt | Set-WinRMListener
 
    Sets the WinRM service on the computers in the servers.txt file
 
  .EXAMPLE
    Set-WinRMListener -ComputerName (get-content .\servers.txt) -IPv4Range "10.10.1.0 255.255.255.0"
 
    Sets the WinRM service on the computers in the servers.txt file and restricts the IPv4 address to 10.10.1.0 /24
 
#>
[cmdletBinding()] # El CmdletBinding es un atributo de las funciones que hace que se comporte como los cmdlets compilados escritos en C#.
Param
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames

    # ValueFromPipeline: especifica que el parámetro acepta la entrada de un objeto de canalización.
    # ValueFromPipelinePropertyName: especifica que el parámetro acepta la entrada de una propiedad de un objeto de canalización.
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName,
    
    # Enter the IPv4 address range for the WinRM listener
    [Parameter(
    HelpMessage='Enter the IPv4 address range for the WinRM listener')]
    [String]$IPv4Range = '*',
    
    # Enter the IPv4 address range for the WinRM listener
    [Parameter(
    HelpMessage='Enter the IPv4 address range for the WinRM listener')]
    [String]$IPv6Range = '*'
  )
Begin # Los parametros dentro del bloque BEGIN son opcionales y solo se ejecutarán una vez por llamada
      # a la función. Este bloque se configura para inicializar objetos tales como variables, conexiones 
      # de base de datos o matrices que se utilizarán en toda la función. Cualquier variable que se cree 
      # en el bloque BEGIN será accesible en cualquier otra parte de la función. Nuevamente, el bloque 
      # BEGIN es opcional y no es obligatorio si solo desea utilizar los bloques PROCESS o END.
      #
      # En este caso declara las variables necesarias para modificar el registro.
  {
    Write-Debug 'Opening Begin block'
    $HKLM = 2147483650
    $Key = 'SOFTWARE\Policies\Microsoft\Windows\WinRM\Service'
    $DWORDName = 'AllowAutoConfig' 
    $DWORDvalue = '0x1'
    $String1Name = 'IPv4Filter'
    $String2Name = 'IPv6Filter'
    Write-Debug 'Finished begin block variables are built'
  }
Process # El bloque PROCESS se utiliza para especificar el código que se ejecutará continuamente 
        # en cada objeto que pueda pasarse a la función. Una función puede tener un bloque PROCESS
        # sin los otros bloques (BEGIN y END), y es obligatorio si se establece un parámetro 
        # para aceptar la entrada de datos (pipeline o parametros).

        # Cuando define un parámetro que acepta entrada de a la función de forma canalizada, 
        # [Parameter(ValueFromPipeline)] se obtiene un array de forma implícita:

        # Con la entrada canalizada mediante Pipeline, PowerShell llama al bloque PROCESS una vez para cada objeto de entrada. 
        # Por el contrario, si se pasa la entrada como un valor de parámetro solo ingresará al proceso una vez. 

        # Por lo tanto el bloque PROCESS se comporta de manera diferente dependiendo de cómo se pasa la entrada:

        # Los valores pasados ​​por parámetro solo se procesarán una vez, y el ciclo se ejecutará a través de todos los objetos pasados.
        # Al pasar por canalización Pipeline, el bucle foreach se ejecutará una vez, pero el bloque PROCESS se ejecutará una vez para 
        # cada elemento de la canalización.

        # En este caso un Foreach que recorrera el array $ComputerName
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to create remote registry handle'
            $Reg = New-Object -TypeName System.Management.ManagementClass -ArgumentList \\$computer\Root\default:StdRegProv
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to create Remote Key'
            if (($reg.CreateKey($HKLM, $key)).returnvalue -ne 0) {Throw 'Failed to create key'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attemping to set DWORD value'
            if (($reg.SetDWORDValue($HKLM, $Key, $DWORDName, $DWORDvalue)).ReturnValue -ne 0) {Throw 'Failed to set DWORD'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set first REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $String1Name, $IPv4Range)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set second REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $String2Name, $IPv6Range)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }
  }
End # Una vez que todos los objetos se han enviado a través de la canalización, se llama al bloque END.
    # Al igual que el bloque BEGIN, el bloque END se llama una vez por llamada a la función y es opcional. 
    # Es una buena práctica tener un bloque END incluso si se deja vacío.
  {}
}

Function Restart-WinRM
{
<#
  .SYNOPSIS
    Restarts the WinRM service
 
  .DESCRIPTION
    Uses the win32_service class to get and restart the WinRM service. This function was designed to be used after Set-WinRMListener to allow the new registry configuration to take hold.
 
  .EXAMPLE
    Restart-WinRM -computername TestVM
 
    Restarts the WinRM service on TestVM
 
  .EXAMPLE
    Get-Content .\servers.txt | Restart-WinRM
 
    Restarts the WinRM service on all the computers in the servers.txt file
 
#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to stop WinRM'
            (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StopService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Stopped') {Throw 'Failed to Stop WinRM'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to start WinRM'
            (Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).StartService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='WinRM'" -ComputerName $computer).state -notlike 'Running') {Throw 'Failed to Start WinRM'}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
End {}
}

Function Set-WinRMStartup
{
<#
  .SYNOPSIS
    Changes the startup type of the WinRM service to automatic
 
  .DESCRIPTION
    Uses the Win32_service class to change the startup type on the WinRM service to Automatic
 
  .EXAMPLE
    Set-WinRMStartup -Computername TestVM
 
    Sets the WinRM service startup type to Automatic on test VM
 
#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            if (((Get-WmiObject win32_service -Filter "name='WinRM'" -ComputerName $computer).ChangeStartMode('Automatic')).ReturnValue -ne 0) {Throw "Failed to change WinRM Startup type on $Computer"}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "Failed to change startmode on the WinRM service for $computer"
            break
          }
        Write-Verbose "Completed operations on $computer"
        Write-Debug "Finished process block for $Computer"
      }
  }
}

Function Set-WinRMFirewallRule
{
<#
  .SYNOPSIS
    Set-WinRMFirewallRule will configure a remote computer's firewall
 
  .DESCRIPTION
    Set-WinRMFirewallRule uses the system.management.managementclass object to create 2 registry keys. These registry entries will configure the Windows Firewall service when it nest restarts. Use Restart-WindowsFirewall to restart the Windows Firewall service after it has been successfully configured. Only use this function if you're using the windows firewall.
 
  .EXAMPLE
    Get-Content .\servers.txt | Set-WinRMFirewallRule
 
    Sets the rules for the windows firewall on the computers in the servers.txt file
 
  .EXAMPLE
    Set-WinRMFirewallRule -ComputerName (get-content .\servers.txt)
 
    Sets the rules for the windows firewall on the computers in the servers.txt file
 
#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin 
  {
    Write-Debug 'Opening Begin block'
    $HKLM = 2147483650
    $Key = 'SOFTWARE\Policies\Microsoft\WindowsFirewall\FirewallRules'
    $Rule1Value = 'v2.20|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Public|LPort=5985|RA4=LocalSubnet|RA6=LocalSubnet|App=System|Name=@FirewallAPI.dll,-30253|Desc=@FirewallAPI.dll,-30256|EmbedCtxt=@FirewallAPI.dll,-30267|'
    $Rule1Name = 'WINRM-HTTP-In-TCP-PUBLIC'
    $Rule2Value = 'v2.20|Action=Allow|Active=TRUE|Dir=In|Protocol=6|Profile=Domain|Profile=Private|LPort=5985|App=System|Name=@FirewallAPI.dll,-30253|Desc=@FirewallAPI.dll,-30256|EmbedCtxt=@FirewallAPI.dll,-30267|'
    $Rule2Name = 'WINRM-HTTP-In-TCP'
    Write-Debug 'Completed Begin block, Finshed creating variables'
  }
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to create remote registry handle'
            $Reg = New-Object -TypeName System.Management.ManagementClass -ArgumentList \\$computer\Root\default:StdRegProv
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to create Remote Key'
            if (($reg.CreateKey($HKLM, $key)).returnvalue -ne 0) {Throw 'Failed to create key'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set first REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $Rule1Name, $Rule1Value)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to set second REG_SZ Value'
            if (($reg.SetStringValue($HKLM, $Key, $Rule2Name, $Rule2Value)).ReturnValue -ne 0) {Throw 'Failed to set REG_SZ'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
}

Function Restart-WindowsFirewall
{
<#
  .SYNOPSIS
    Restarts the windows firewall service
 
  .DESCRIPTION
    Uses the Win32_Service class to restart the windows firewall service.
 
  .EXAMPLE
    Restart-WindowsFirewall -computername TestVM
 
    Restarts the Windows Firewall service on TestVM
 
  .EXAMPLE
    Get-Content .\servers.txt | Restart-WindowsFirewall
 
    Restarts the Windows Firewall service on all the computers in the servers.txt file
 
#>
[CmdletBinding()]
Param 
  (
    # Enter a ComputerName or IP Address, accepts multiple ComputerNames
    [Parameter( 
    ValueFromPipeline=$True, 
    ValueFromPipelineByPropertyName=$True,
    Mandatory=$True,
    HelpMessage='Enter a ComputerName or IP Address, accepts multiple ComputerNames')] 
    [String[]]$ComputerName
  )
Begin {}
Process 
  {
    Foreach ($computer in $ComputerName)
      {
        Write-Verbose "Beginning function on $computer"
        Write-Debug "Opening process block of $Computer"
        Try
          {
            Write-Verbose 'Attempting to stop MpsSvc'
            (Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).StopService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).state -notlike 'Stopped') {Throw 'Failed to Stop MpsSvc'}
          }
        Catch 
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Try 
          {
            Write-Verbose 'Attempting to start MpsSvc'
            (Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).StartService() | Out-Null
            Start-Sleep -Seconds 10
            if ((Get-WmiObject win32_service -Filter "Name='MpsSvc'" -ComputerName $computer).state -notlike 'Running') {Throw 'Failed to Start MpsSvc'}
          }
        Catch
          {
            Write-Warning $_.exception.message
            Write-Warning "The function will abort operations on $Computer"
            break
          }
        Write-Verbose "Successfully completed operation on $computer"
        Write-Debug "Finished process block for $Computer"
      }    
  }
End {}
}