Attribute VB_Name = "Utils"
Public Const OU_SCOPE_DRIVE  As String = "https://www.googleapis.com/auth/drive"
Public Const OU_SCOPE_DRIVE_FILE As String = "https://www.googleapis.com/auth/drive.file"
Public Const OU_SCOPE_DRIVE_METADATA_READONLY As String = "https://www.googleapis.com/auth/drive.metadata.readonly"
Public Const OU_SCOPE_DRIVE_APPDATA As String = "https://www.googleapis.com/auth/drive.appdata"
Public Const OU_SCOPE_DRIVE_METADATA As String = "https://www.googleapis.com/auth/drive.metadata"
Public Const OU_SCOPE_SPREADSHEETS As String = "https://www.googleapis.com/auth/spreadsheets"
Public Const OU_SCOPE_DRIVE_READONLY As String = "https://www.googleapis.com/auth/drive.readonly"
Public Const OU_SCOPE_SPREADSHEETS_READONLY As String = "https://www.googleapis.com/auth/spreadsheets.readonly"
Public Const OU_SCOPE_PHOTOS_READONLY As String = "https://www.googleapis.com/auth/drive.photos.readonly"
Public Const OU_SCOPE_DRIVE_SCRIPTS As String = "https://www.googleapis.com/auth/drive.scripts"

Public Const HTTP_GET = "GET"
Public Const HTTP_POST = "POST"
Public Const HTTP_PUT = "PUT"
Public Const HTTP_DELETE = "DELETE"
Public Const HTTP_PATCH = "PATCH"

Public Enum CODE_STATUS_HTTP
    SUCCESSFUL = 200
    CREATED = 201
    NO_CONTENT = 204
    FOUND = 302
    RESUME_INCOMPLETE = 308
    FAILED = 400
    RATE_LIMIT = 403
    NOT_FOUND = 404
    SERVICE_UNAVAILABLE = 503
End Enum

Private Enum ERR_UTIL
    READ_FAILD = 6100
    WRITE_FAILD = 6101
    INTERPOLATE_FAILED = 6102
End Enum

Public Function URLEncode(ByVal str As String) As String
'    Esta función codifica una cadena de texto según las reglas de codificación
'    de URL, reemplazando caracteres especiales con sus equivalentes codificados.
'
'   Args:
'       str(String): Cadena de texto que se desea codificar.
'
'   Returns:
'       String: Cadena codificada.

    Dim chrSpecial As New Dictionary
    Dim key
    
    With chrSpecial
'    El primer Item debe ser el de porcentaje en este diccionario,
'    debido a que al los demas valores reemplazados incluyen el porcentaje
        .Add "%", "%25"
        .Add " ", "%20"
        .Add "=", "%3D"
        .Add ",", "%2C"
        .Add """", "%22"
        .Add "<", "%3C"
        .Add ">", "%3E"
        .Add "#", "%23"
        .Add "|", "%7C"
        .Add "/", "%2F"
        .Add ":", "%3A"
        .Add "_", "%5F"
'        .Add "'", "%27"
    End With
    
    For Each key In chrSpecial.Keys
        str = Replace(str, key, chrSpecial(key))
    Next key
    
    Set chrSpecial = Nothing
    
    URLEncode = str
    
End Function
Public Function generateString(Optional lenght = 8, Optional includeNumber = False) As String
    
    Dim dicUpper  As New Dictionary
    Dim dicLower As New Dictionary
    Dim dicNumbers As New Dictionary
    Dim randomString As String
    Dim character As Long
    Dim i As Integer
    
    'n?meros del 48 al 57
    'letras may?sculas 65 al 90
    'letras min?sculas 97 al 122
    
    For i = 48 To 57
        With dicNumbers
            .Add i, Empty
        End With
    Next i
    
    For i = 65 To 90
        With dicUpper
            .Add i, Empty
        End With
    Next i
    
    For i = 97 To 122
        With dicLower
            .Add i, Empty
        End With
    Next i
    
    
    Do While Len(randomString) <= lenght
        Randomize
        character = Int((122 - 48 + 1) * Rnd + 48)
        
         If (dicUpper.Exists(character) Or dicLower.Exists(character)) Or _
            (dicNumbers.Exists(character) And includeNumber) Then
            
            randomString = randomString + Chr(character)
            
        End If
    Loop
    
    Set dicUpper = Nothing
    Set dicLower = Nothing
    Set dicNumbers = Nothing
    
    generateString = randomString
    
End Function
Public Function existsFile(ByVal pathFile As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    
    existsFile = fso.FileExists(pathFile)
    Set fso = Nothing
    
End Function
Public Function existsDirectory(ByVal directory As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    
    existsDirectory = fso.FolderExists(directory)
    Set fso = Nothing
    
End Function
Public Function fstring(ByVal text As String, ParamArray values()) As String
    
    Dim i As Integer
    
    On Error GoTo Catch
    
    For i = LBound(values) To UBound(values)
        text = Replace(text, "{" & i & "}", values(i))
    Next i
    
    fstring = text
    
    Exit Function
    
Catch:
    Err.Raise ERR_UTIL.INTERPOLATE_FAILED, Description:="Failed to interpolate string."
    
End Function

Public Function readFile(ByVal pathFile As String) As String
    
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream
    Dim content As String
    
    On Error GoTo Catch
    
    If fso.FileExists(pathFile) Then
        Set t = fso.OpenTextFile(pathFile, ForReading)
        content = t.ReadAll
        t.Close
        readFile = content
        
        Set fso = Nothing
        Set t = Nothing
    End If
    
    Exit Function
    
Catch:
    Err.Raise Number:=ERR_UTIL.READ_FAILD, Description:="Failed to read file."
    
End Function

Public Function writeFile(ByVal content As String, ByVal pathTarget As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream

    On Error GoTo Cath
    
    Set t = fso.CreateTextFile(pathTarget, True)
    t.Write content
    t.Close
    
    writeFile = True
    
    Set fso = Nothing
    Set t = Nothing
    
    Exit Function
    
Cath:

    Err.Raise Number:=ERR_UTIL.WRITE_FAILD, Description:="Failed to write file."
    
End Function
Public Function sourceToBinary(ByVal pathFile As String) As Byte()
    
    Dim numFile As Byte
    Dim buffer() As Byte

    numFile = FreeFile
    
    Open pathFile For Binary Access Read As #numFile
        ReDim buffer(LOF(numFile))
        Get #numFile, , buffer
    Close #numFile
    
    sourceToBinary = buffer
    
End Function
Public Function binaryToBase64(ByRef arr() As Byte) As String
    
    Dim XML As MSXML2.DOMDocument60
    Dim DocElem As MSXML2.IXMLDOMElement
    
    Set XML = New MSXML2.DOMDocument60
    Set DocElem = XML.createElement("Base64Data")
    DocElem.DataType = "bin.base64"
    
    DocElem.nodeTypedValue = arr

    binaryToBase64 = DocElem.text

    Set XML = Nothing
    Set DocElem = Nothing

End Function
Public Function sourceToBase64(ByVal pathFile As String) As String

    Dim buffer() As Byte
    Dim base64 As String
        
    buffer = sourceToBinary(pathFile)
    base64 = binaryToBase64(buffer)
        
    sourceToBase64 = base64
    
    Erase buffer

End Function
Public Function createParteRelated(ByVal filePath As String, _
                                    ByVal boundary As String, _
                                    ByRef fileObject As Dictionary) As String
        
    Dim related As String
    Dim body As String
    Dim start_boundary As String
    Dim finish_boundary As String
    Dim fileName As Variant
    Dim base64 As String
    Dim strTmp As String
    Dim mimeType As String
    
    base64 = sourceToBase64(filePath)
    start_boundary = "--" + boundary
    finish_boundary = start_boundary + "--"
    
    fileName = Split(filePath, "\")
    fileName = fileName(UBound(fileName))

    fileObject.Add "name", fileName
    body = JsonConverter.ConvertToJson(fileObject, 2)
    mimeType = IIf(fileObject.Exists("mimeType"), fileObject("mimeType"), "")
    
    strTmp = "{0}{1}Content-Type: application/json; charset=UTF-8{1}{1}{2}{1}{0}{1}Content-Type: {3}{1}Content-Transfer-Encoding: base64{1}{1}{4}{1}{5}"
    related = fstring(strTmp, start_boundary, vbNewLine, body, mimeType, base64, finish_boundary)

    createParteRelated = related

End Function

