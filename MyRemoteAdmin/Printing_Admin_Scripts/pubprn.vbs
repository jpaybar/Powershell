'----------------------------------------------------------------------
'    pubprn.vbs - publish printers from a non Windows 2000 server into Windows 2000 DS
'    
'
'     Arguments are:-
'        server - format server
'        DS container - format "LDAP:\\CN=...,DC=...."
'
'
'    Copyright (c) Microsoft Corporation 1997
'   All Rights Reserved
'----------------------------------------------------------------------

'--- Begin Error Strings ---

Dim L_PubprnUsage1_text
Dim L_PubprnUsage2_text
Dim L_PubprnUsage3_text      
Dim L_PubprnUsage4_text      
Dim L_PubprnUsage5_text      
Dim L_PubprnUsage6_text      

Dim L_GetObjectError1_text
Dim L_GetObjectError2_text

Dim L_PublishError1_text
Dim L_PublishError2_text     
Dim L_PublishError3_text
Dim L_PublishSuccess1_text


L_PubprnUsage1_text      =   "Uso: [cscript] pubprn.vbs servidor ""LDAP://OU=..,DC=..."""
L_PubprnUsage2_text      =   "       servidor es un nombre de servidor de Windows (p.e.: Servidor) o nombre de impresora UNC (\\Servidor\Impresora)"
L_PubprnUsage3_text      =   "       ""LDAP://CN=...,DC=..."" es la ruta DS del contenedor de destino"
L_PubprnUsage4_text      =   ""
L_PubprnUsage5_text      =   "Ejemplo 1: pubprn.vbs MiServidor ""LDAP://CN=MiContenedor,DC=MiDominio,DC=Compañía,DC=Com"""
L_PubprnUsage6_text      =   "Ejemplo 2: pubprn.vbs \\MiServidor\Impresora ""LDAP://CN=MiContenedor,DC=MiDominio,DC=Compañía,DC=Com"""

L_GetObjectError1_text   =   "Error: no se encuentra la "
L_GetObjectError2_text   =   " ruta de acceso."
L_GetObjectError3_text   =   "Error: No se puede tener acceso "

L_PublishError1_text     =   "Error: Pubprn no puede publicar impresoras desde "
L_PublishError2_text     =   " porque tiene instalado Windows 2000 u otro más reciente."
L_PublishError3_text     =   "Error al publicar la impresora "
L_PublishError4_text     =   "Error: "
L_PublishSuccess1_text   =   "Impresora publicada: "

'--- End Error Strings ---


set Args = Wscript.Arguments
if args.count < 2 then
    wscript.echo L_PubprnUsage1_text
    wscript.echo L_PubprnUsage2_text
    wscript.echo L_PubprnUsage3_text
    wscript.echo L_PubprnUsage4_text
    wscript.echo L_PubprnUsage5_text
    wscript.echo L_PubprnUsage6_text
    wscript.quit(1)
end if

ServerName= args(0)
Container = args(1)

if 1 <> InStr(1, Container, "LDAP://", vbTextCompare) then
    wscript.echo L_GetObjectError1_text & Container & L_GetObjectError2_text
    wscript.quit(1)
end if

on error resume next
Set PQContainer = GetObject(Container)

if err then
    wscript.echo L_GetObjectError1_text & Container & L_GetObjectError2_text
    wscript.quit(1)
end if
on error goto 0



if left(ServerName,1) = "\" then

    PublishPrinter ServerName, ServerName, Container

else

    on error resume next

    Set PrintServer = GetObject("WinNT://" & ServerName & ",computer")

    if err then
        wscript.echo L_GetObjectError3_text & ServerName & ": " & err.Description
        wscript.quit(1)
    end if

    on error goto 0


    For Each Printer In PrintServer
        if Printer.class = "PrintQueue" then PublishPrinter Printer.PrinterPath, ServerName, Container
    Next


end if




sub PublishPrinter(UNC, ServerName, Container)

    
    Set PQ = WScript.CreateObject("OlePrn.DSPrintQueue.1")

    PQ.UNCName = UNC
    PQ.Container = Container

    on error resume next

    PQ.Publish(2)

    if err then
        if err.number = -2147024772 then
            wscript.echo L_PublishError1_text & Chr(34) & ServerName & Chr(34) & L_PublishError2_text
            wscript.quit(1)
        else
            wscript.echo L_PublishError3_text & Chr(34) & UNC & Chr(34) & "."
            wscript.echo L_PublishError4_text & err.Description
        end if
    else
        wscript.echo L_PublishSuccess1_text & PQ.Path
    end if

    Set PQ = nothing

end sub
