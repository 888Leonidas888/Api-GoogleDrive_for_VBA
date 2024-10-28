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

Public Const HT_POST As String = "POST"
Public Const HT_GET As String = "GET"
Public Const HT_PUT As String = "PUT"
Public Const HT_DELETE As String = "DELETE"
Public Const HT_PATCH As String = "PATCH"

Public Const GO_SUCCESSFUL As Integer = 200
Public Const GO_CREATED As Integer = 201
Public Const GO_NO_CONTENT As Integer = 204
Public Const GO_FOUND As Integer = 302
Public Const GO_RESUME_INCOMPLETE As Integer = 308
Public Const GO_FAILED As Integer = 400
Public Const GO_RATE_LIMIT As Integer = 403
Public Const GO_NOT_FOUND As Integer = 404
Public Const GO_SERVICE_UNAVAILABLE = 503

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

