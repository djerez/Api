Option Explicit

' Constantes.
Const ForReading = 1 'Open a file for reading only. You can't write to this file.
Const ForWriting = 2 'Open a file for writing. If a file with the same name exists, its previous contents are overwritten.
Const ForAppending = 8 'Open a file and write to the end of the file.
Const TristateUseDefault = -2 'Opens the file using the system default.
Const TristateTrue = -1 'Opens the file as Unicode.
Const TristateFalse =  0 'Opens the file as ASCII.

' Decaraciones.
Dim fso
Dim loArgs
Dim lotxtfile
Dim lcFile
Dim lcTxt
Dim q2
Dim lcSearch1
Dim lcReplace1
Dim lcSearch2
Dim lcReplace2
Dim lcFiles
Dim i
q2 = chr( 34 )
lcSearch1 = "</con:description><con:settings/><con:parameters/><con:method"
lcReplace1 = "</con:description><con:settings/><con:parameters>" _
                            & "<con:parameter>" _
                            & "<con:name>access_token</con:name>" _
                            & "<con:style>QUERY</con:style>" _
                            & "</con:parameter>" _
                            & "</con:parameters><con:method"
lcSearch2 = "basePath=" & q2 & "/" & q2
lcReplace2 = "basePath=" & q2 & q2
lcFiles = ""
Set loArgs = WScript.Arguments
Set fso = WScript.CreateObject( "Scripting.FileSystemObject" )
WScript.Echo "Inicio del proceso" & VbCrLf & "Archivos a procesar: " & loArgs.Count 

' Recorre todos los archivos que se pasaron como argumentos.
For i = 0 To loArgs.Count - 1
    ' Selecciona el archivo y lo lee completo.
    lcFile = loArgs( i )
    lcFiles = lcFiles & lcFile & VbCrLf
    Set lotxtfile = fso.OpenTextFile( lcFile, ForReading, True )
    lcTxt = lotxtfile.ReadAll
    lotxtfile.Close

    ' Realiza los remplazos.
    lcTxt = Replace( lcTxt, lcSearch1, lcReplace1 )
    lcTxt = Replace( lcTxt, lcSearch2, lcReplace2 )

    ' Graba el archivo actualizado.
    Set lotxtfile = fso.OpenTextFile( lcFile, ForWriting, True )
    lotxtfile.Write lcTxt
    lotxtfile.Close

    Set lotxtfile = Nothing

Next

' Muestra la lista de archivos procesados
WScript.Echo Trim( lcFiles )
Set fso = Nothing
Set loArgs = Nothing

WScript.quit