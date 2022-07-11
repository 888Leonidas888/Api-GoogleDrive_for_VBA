VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleDriveService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'>>>Desarrollado por: Jhony Escriba Bautista                                                                                ***************
'>>>Correo : jhonny14_1@hotmail.com                                                                                            ***************
'>>>Año de lanzamiento 2022                                                                                                                  ***************
'>>>Version de Api Google Drive v3                                                                                                       ***************
'***************************************************************************

'***************************************************************************
'Activar las siguientes librerias:
'1.-Microsoft XML, v6.0
'2.-Microsoft Scripting Runtime
'3.-Microsoft Activex Data Objects 6.1 Library

'Instalar el siguiente módulo de GitHub
'1.-JSON Converter for VBA , ir al GitHub https://github.com/VBA-tools/VBA-JSON
'El módulo VBA-JSON  no es propio es creado por otro desarrollador,asi que lo créditos correspondientes
'para su creador.
'***************************************************************************

Private Const APP_NAME As String = "Service Google Drive"
Private Const version As String = "1.0"

Private YourApiKey As String
Private YourAccessToken As String
Private uriResumable As String
Private numberStatus As Integer
Private sizeByteFile As Long
Private fileBase64 As String
Private ReadBinaryFile As String
Private arrayResponse() As String
Private headerRange As String

Private Const gdGet As String = "GET"
Private Const gdPost As String = "POST"
Private Const gdPatch As String = "PATCH"
Private Const gdDelete As String = "DELETE"
Private Const gdPut As String = "PUT"

'EXPRESIONES LITERALES
Private Const gdAnEmptySpace As String = " "    'Un espacio
Private Const gdEqual As String = "="
Private Const gdComa As String = ","
Private Const gdDoubleQuote As String = """"    '(se coloca dobles dentro de dobles para que VBA lo entienda como 'UNA COMILLA DOBLE')
Private Const gdSmallerThan As String = "<"
Private Const gdGreaterThan As String = ">"
Private Const gdHashtags As String = "#"
Private Const gdPercentage As String = "%"
Private Const gdPipe As String = "|"
Private Const gdSlash As String = "/"
Private Const gdDoublePoint As String = ":"
Private Const gdUnderScore As String = "_"


'CODIFICACION DE URL
Private Const gdEqAnEmptySpace As String = "%20"
Private Const gdEqEqual As String = "%3D"
Private Const gdEqComa As String = "%2C"
Private Const gdEqDoubleQuote As String = "%22"
Private Const gdEqSmallerThan As String = "%3C"
Private Const gdEqGreaterThan As String = "%3E"
Private Const gdEqHashtags As String = "%23"
Private Const gdEqPercentage As String = "%25"
Private Const gdEqPipe As String = "%7C"
Private Const gdEqSlash As String = "%2F"
Private Const gdEqDoublePoint As String = "%3A"
Private Const gdEqUnderScore As String = "%5F"


Private http As New MSXML2.ServerXMLHTTP60
Private uri As String
Private boundary As String
Private Function CreateParteRelated(ByVal filePath As String, _
                                                                                        ByVal parents As String, _
                                                                                        ByVal mimeType As String) As String
        
    Dim related As String
    Dim strJson As String
    Dim start_boundary As String, finish_boundary As String
    Dim fileName
    Dim tranferEncoding As String
    Dim base64 As String

    tranferEncoding = "base64"
    base64 = EncodeBase64(filePath)
    start_boundary = "--" + boundary
    finish_boundary = start_boundary + "--"
    
    fileName = Split(filePath, "\")
    fileName = fileName(UBound(fileName))

    'Crear json con los metadatos
    strJson = "{" & vbNewLine & _
                            """name"":" + """" + fileName + """" + "," & vbNewLine & _
                            """mimeType"":" + """" + mimeType + """" + "," & vbNewLine & _
                            """parents"":" + """" + parents + """" + "," & vbNewLine & _
                            """description"":""Subido desde Api Google Drive para VBA""" & vbNewLine & _
                            "}"
    '---------------------------Fin

    'Part related----------------
    related = start_boundary + vbNewLine + _
                        "Content-Type: application/json; charset=UTF-8" + vbNewLine + vbNewLine + _
                        strJson + vbNewLine + _
                        start_boundary + vbNewLine + _
                        "Content-Type: " + mimeType + vbNewLine + _
                        "Content-Transfer-Encoding: " + tranferEncoding + vbNewLine + vbNewLine + _
                        base64 + vbNewLine + _
                        finish_boundary
    '-----------------------------Fin
        
    CreateParteRelated = related
    
End Function

Private Function ExistsFile(ByVal pathFile As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject
    Dim exists As Boolean
    
    ExistsFile = fso.FileExists(pathFile)
    
    Set fso = Nothing
    
End Function
Public Sub Filedownload3(ByVal fileId As String, Optional pathTarget)

    Dim itemJson As Dictionary
    Dim response As String
    Dim nameFile As String
    Dim arr() As Byte
    Dim numFile As Byte
    Dim strNewFile As String
    Dim fso As New Scripting.FileSystemObject
    
    On Error GoTo Cath
    
    If IsMissing(pathTarget) Then pathTarget = Environ("UserProfile")
    
    If fso.FolderExists(pathTarget) Then
        
        response = GetMeta(fileId, "name")

        Set itemJson = JsonConverter.ParseJson(response)

        nameFile = itemJson("name")
        strNewFile = pathTarget + "\" + nameFile
        
        numFile = FreeFile
        
        arr = FileDownload(fileId)
        
        Open strNewFile For Binary Access Write As #numFile
            Put #numFile, , arr
        Close #numFile
        
        Erase arr
        Set itemJson = Nothing
        MsgBox "Descarga completa en : " + vbNewLine + strNewFile, vbInformation, APP_NAME
    Else
        Err.Raise Number:=vbObjectError + 1, description:="el directorio de destino no existe"
    End If
    
    Set fso = Nothing
    
    Exit Sub
    
Cath:
    MsgBox "Lo siento no pude descargar el archivo", vbCritical, APP_NAME
    Debug.Print Err.Number; Err.description
    On Error GoTo 0
End Sub

Private Function RandomString(Optional ByVal lenght = 6) As String
        
    Rem esta función crea una cadena aleatoria(mayúsculas,minúsculas y números)
    'prederteminado largo de 6
    
      Dim password$, character$, n%, i%, j%, a%
      Dim character_excluded(0 To 12) As Byte
      'números del 48 al 57
      'letras mayúsculas 65 al 90
      'letras minúsculas 97 al 122
   
      Rem Application.Volatile False
      
      For i = 58 To 64
            character_excluded(n) = i
            n = n + 1
      Next i
      
      For i = 91 To 96
            character_excluded(n) = i
            n = n + 1
      Next i

      For a = 1 To lenght
otra_vez:
            Randomize
            character = Int((122 - 48 + 1) * Rnd + 48)
            
            For j = 0 To 12
                  If character = character_excluded(j) Then
                        GoTo otra_vez
                  End If
            Next j
            password = password & Chr(character)
      Next a
      
      RandomString = password
      
      Erase character_excluded
      
End Function

Private Function CodeUri(ByVal value As String) As String
    
    Dim listValue As Variant
    Dim listCode As Variant
    
    listValue = Array( _
                                            gdPercentage, _
                                            gdAnEmptySpace, _
                                            gdEqual, _
                                            gdComa, _
                                            gdDoubleQuote, _
                                            gdSmallerThan, _
                                            gdGreaterThan, _
                                            gdHashtags, _
                                            gdPipe, _
                                            gdSlash, _
                                            gdDoublePoint, _
                                            gdUnderScore _
                                            )
                                            
    listCode = Array( _
                                        gdEqPercentage, _
                                        gdEqAnEmptySpace, _
                                        gdEqEqual, _
                                        gdEqComa, _
                                        gdEqDoubleQuote, _
                                        gdEqSmallerThan, _
                                        gdEqGreaterThan, _
                                        gdEqHashtags, _
                                        gdEqPipe, _
                                        gdEqSlash, _
                                         gdEqDoublePoint, _
                                         gdEqUnderScore _
                                         )
    
    For i = LBound(listValue) To UBound(listValue)
        value = Replace(value, listValue(i), listCode(i))
    Next i
    
    CodeUri = value
    
End Function
Public Function FileDelete(ByVal fileId As String) As Boolean
    
    Rem si la solicitud ha tenido,devuelve un cuerpo vacio
    'validar por la propiedad serviceStatus que haya devuelto un codigo 2* y
    'la respuesta de esta funcion vacia
    
    uri = "https://www.googleapis.com/drive/v3/files/" & fileId & "?key=" & YourApiKey & " HTTP/1.1"
    
    With http
        .Open gdDelete, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        numberStatus = .Status
    End With
    
    If numberStatus = 204 Then
        FileDelete = True
    Else
        FileDelete = False
    End If
    
End Function

Public Function FileCopy(ByVal fileId As String, _
                                                            ByVal fields As String, _
                                                            ByVal parentTarget As String, _
                                                            Optional description) As String
    
    Dim strJson As String
    
    If IsMissing(description) Then description = "Copia para Google Drive desde VBA"
    
    strJson = "{""parents"":[""" & parentTarget & """],""description"":""" & description & """}"
    
    fields = CodeUri(fields)
    uri = "https://www.googleapis.com/drive/v3/files/" & fileId & "/copy?fields=" & fields & "&key=" & YourApiKey & " HTTP/1.1"
    
    With http
        .Open gdPost, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-Type", "application/json"
        .send (strJson)
        numberStatus = .Status
    End With
    
    FileCopy = http.responseText
    
End Function

Private Function FileReadContent(ByVal filePath As String) As String
    
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream
    Dim contents As String
    
    Set t = fso.OpenTextFile(filePath, ForReading)
    contents = t.ReadAll
    t.Close
    
    FileReadContent = contents
    
    Set fso = Nothing
    Set t = Nothing
    
End Function

Private Function FileReadBinary(ByVal filePath As String) As Byte()
    
    Dim numFile As Byte
    Dim buffer() As Byte

    numFile = FreeFile
    
    Open filePath For Binary Access Read As #numFile
        ReDim buffer(LOF(numFile))
        Get #numFile, , buffer
    Close #numFile
    
    FileReadBinary = buffer
    
End Function


Public Function FileUploadMultipart(ByVal pathFile As String, ByVal parent As String, Optional mimeType) As String

    Dim related As String
    boundary = RandomString(15)
    
    If IsMissing(mimeType) Then mimeType = "application/octet-stream"
    
    related = CreateParteRelated(pathFile, parent, mimeType)

    numberStatus = 0
    
    uri = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&key=" & YourApiKey & " HTTP/1.1"
    
    With http
        .Open gdPost, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-type", "multipart/related; boundary=" + boundary
        .setRequestHeader "Content-Lenght", FileLen(pathFile)
        .send (related)
        numberStatus = http.Status
    End With
    
    FileUploadMultipart = http.responseText
    boundary = Empty
    
End Function
Public Function FileUpLoadResumable(ByVal IsUploadChunk As Boolean, _
                                                                                            ByVal parent As String, _
                                                                                            ByVal pathFile As String, _
                                                                                            Optional mimeType) As String
    '@IsUploadChunk(bolean);true para subida en fragmentos y false lo subida total en una sola petición.NOTA:use false
    '@ipParents: id del folder donde se cargará el recurso.
    '@pathFile: ruta completa del recurso a subir.
    '@mimeType: opcional si no se conoce , revise el archivo README.txt para mas detalles.
    '@retorno: Si la carga ha sido exitosa el codigo es 200  ó 201 , la respuesta serám los metadatos del archivo subido.
    
    Dim bufer() As Byte
    Dim numFile As Byte
    Dim response As String
    
    On Error GoTo Cath
    
    Dim chunk, startByte, endByte, totalByte
    Dim strChunk As String

    If IsMissing(mimeType) Then mimeType = "application/octet-stream"
    
    totalByte = FileLen(pathFile)
    numFile = FreeFile

    Rem comenzamos con uri de sesión reanudable------------------->>>>
    uri = GetUriResumable(IsUploadChunk, pathFile, parent, mimeType)
    '-------------------------------------------------------------------------<<<<<
    
    If numberStatus <> 200 Then
        Exit Function
    End If
    
    If Not IsUploadChunk Then
                    
                    chunk = totalByte
                    startByte = 0
                    endByte = chunk - 1
        
                    numFile = FreeFile
                    Rem -------bufer
                    Open pathFile For Binary Access Read As #numFile
                        ReDim bufer(1 To chunk)
                        Get #numFile, startByte + 1, bufer
                    Close #numFile
                    '-------------fin bufer
                    
                    strChunk = "bytes " & startByte & "-" & endByte & "/" & totalByte
                    response = ContinueUpload(uri, mimeType, chunk, strChunk, bufer)
                    
                    Erase bufer
                    FileUpLoadResumable = response
    Else
                        
                    chunk = totalByte
                    startByte = 0
                    endByte = chunk - 1
            
                    Do
                    
                        numFile = FreeFile
                        Rem -------bufer
                        Open pathFile For Binary Access Read As #numFile
                            ReDim bufer(1 To chunk)
                            Get #numFile, startByte + 1, bufer
                        Close #numFile
                        '-------------fin bufer
                        
                        strChunk = "bytes " & startByte & "-" & endByte & "/" & totalByte
                        
                        Rem-----------------Empezamos la carga
                        response = ContinueUpload(uri, mimeType, chunk, strChunk, bufer)
                        Rem trabajar en una respuesta para el encabezado 'Range'
                        Debug.Print headerRange
                        '                -----------------------------continuar,si falta
                        
                    Loop Until numberStatus <> 308
                    
                    Erase bufer
                    headerRange = Empty
                    
                    FileUpLoadResumable = response
                          
        End If
    
    Exit Function
    
Cath:
    
   Err.Raise Number:=vbObjectError + 1, description:=Err.description
    On Error GoTo 0
End Function
Private Function ContinueUpload(ByVal uri As String, _
                                                                                ByVal mimeType As String, _
                                                                                ByVal chunk As String, _
                                                                                ByVal strChunk As String, _
                                                                                ByRef bufer() As Byte) As String
            
                numberStatus = 0
                
                With http
                    .Open gdPut, uri, False
                    .setRequestHeader "Authorization", "Bearer " & YourAccessToken
                    .setRequestHeader "Content-Range", strChunk
                    .setRequestHeader "Content-Type", mimeType
                    .setRequestHeader "Content-Length", chunk
                    .send (bufer)
                    numberStatus = .Status
                    headerRange = .getResponseHeader("Range")
                    ContinueUpload = .responseText
                End With
                
End Function

Public Property Get ResponseStatus() As Integer
    ResponseStatus = numberStatus
End Property

Public Function FileCreateFolder(ByVal folderName As String, ByVal parent As String) As String
    
    Dim strJson As String
    
    numberStatus = 0
    
    uri = "https://www.googleapis.com/drive/v3/files?fields=id%2Cname%2CwebViewLink&key=" & YourApiKey & " HTTP/1.1"

    strJson = "{""mimeType"":""application/vnd.google-apps.folder"",""name"":""" & folderName & """,""parents"":[""" & parent & """]}"
    
    With http
        .Open gdPost, uri
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-Type", "application/json"
        .send (strJson)
        
        numberStatus = .Status
        
    End With
    
    FileCreateFolder = http.responseText

End Function

Public Function FileCreateMetaData(ByVal strJson As String, ByVal fields As String) As String
    Rem con esta función tbm se pueden crear folderes usando el tipeMime correspondiente
    'revise el archivo README.txt que viene con el proyecto para mayores detalles y referencias.
    
    numberStatus = 0
    
    fields = CodeUri(fields)
    uri = "https://www.googleapis.com/drive/v3/files?fields=" & fields & "&key=" & YourApiKey & " HTTP/1.1"

    With http
        .Open gdPost, uri
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-Type", "application/json"
        .send (strJson)
        
        numberStatus = .Status
        
    End With
    
    FileCreateMetaData = http.responseText

End Function
Public Function FileUploadMedia(ByVal pathFile As String, Optional mimeType) As String

    Dim arr() As Byte
    
    numberStatus = 0
    
    uri = "https://www.googleapis.com/upload/drive/v3/files?uploadType=media&key=" & YourApiKey & " HTTP/1.1"
        
    arr = BinaryCode(pathFile)
    
    If IsMissing(mimeType) Then mimeType = "application/octet-stream"
    
    With http
        .Open gdPost, uri
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .setRequestHeader "Content-type", mimeType
        .setRequestHeader "Content-Length", FileLen(pathFile)
        .setRequestHeader "Content-Transfer-Encoding", "binary"
        .send (arr)
        numberStatus = http.Status
    End With
    
    Erase arr
    
    FileUploadMedia = http.responseText
    
End Function
Public Function FileUpdate(ByVal strJson As String, ByVal fileId As String) As String
 
    Rem  Sí se desea actulizar algun metadato pase una cadena en formato json.
    '@strJson : cadena de texto en formato json
    '@fileId : indique el id del archivo que se desea actualizar
    '@retorna : un json
    
    numberStatus = 0
    
    uri = "https://www.googleapis.com/drive/v3/files?" & fileId & "key=" & YourApiKey & " HTTP/1.1"
    
    With http
        .Open gdPatch, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send (strJson)
        numberStatus = http.Status
    End With
    
    FileUpdate = http.responseText

End Function
Public Function EmptyTrash() As Boolean
    Rem esta función vacia la papelera de forma irreversible;
    Rem si la solictud ha sido exitosa devolvera un codigo 204 de respuesta y un cuerpo vacio
    Rem evalúe por numberStatus
    
    numberStatus = 0
    
    uri = "https://www.googleapis.com/drive/v3/files/trash?key=" & YourApiKey & " HTTP/1.1"
    
    With http
        .Open gdDelete, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        numberStatus = .Status
    End With
    
    If numberStatus = 204 Then
        EmptyTrash = True
    Else
        EmptyTrash = False
    End If
    
End Function

Private Function FileExport(ByVal fileId As String, ByVal mimeType As String) As Byte()
    Rem Exporta un documento de Google Workspace al tipo MIME solicitado
    'y devuelve el contenido de bytes exportado. Tenga en cuenta que el contenido
    'exportado está limitado a 10 MB.
    
    Rem Si tiene éxito, este método devuelve el contenido del archivo como bytes.
    'Para obtener una lista de documentos de Google Workspace y formatos de
    'conversión de tipo MIME, consulte Documentos de Google Workspace y
    'tipos MIME de exportación correspondientes .
    'https://developers.google.com/drive/api/guides/ref-export-formats
    
    numberStatus = 0
    
    mimeType = CodeUri(mimeType)
    
    uri = "https://www.googleapis.com/drive/v3/files/" & fileId & "/export?mimeType=" & mimeType & "&key=" & YourApiKey & " HTTP/1.1"
        
    With http
        .Open gdGet, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        
        numberStatus = .Status
    End With
    
    FileExport = http.responseBody
            
End Function

Public Sub FileDownloadExport(ByVal fileId As String, _
                                                                            ByVal mimeType As String, _
                                                                            ByVal pathTarget As String, _
                                                                            ByVal nameFile As String)
    Rem use esta función paar exportar un documento del Workpace de google a un formato disponible,
    'para ver a que formatos se pueden exportar puede consultar con la función GetMeta o FileList
    'exportLinks = este campo indica los formatos disponibles para este mimeType application/vnd.google-apps.document
    'Para ver los mimeType correspondientes consulte el siguiente enlace:
    'https://developers.google.com/drive/api/guides/ref-export-formats
    'Apriori debe indicar el nombre + la extensión del archivo
    
    
    Dim arr() As Byte
    Dim numFile As Byte
    Dim strNewFile As String
    Dim fso As New Scripting.FileSystemObject
    
    If fso.FolderExists(pathTarget) Then
        
        strNewFile = pathTarget + "\" + nameFile
        numFile = FreeFile
        arr = FileExport(fileId, mimeType)
        
        Open strNewFile For Binary Access Write As #numFile
            Put #numFile, , arr
        Close #numFile
        
        Erase arr
        MsgBox "Descarga completa en : " + vbNewLine + strNewFile, vbInformation, APP_NAME
    Else
        Err.Raise Number:=vbObjectError + 1, description:="el directorio de destino no existe"
    End If
    
    Set fso = Nothing
    
End Sub

Public Function GenerateId(Optional count As Integer) As String()
    
    Rem Genera un  serie de Id  dependiendo el argumento
    '@count = acepta un entero de 1 hasta 1000,predeterminado es 10
    '@retorna un array de tipo string,validar sí tiene datos con IsArrayEmpty
    
    Dim arrStr() As String
    Dim index As Integer
    Dim itemJson As Dictionary
    
    numberStatus = 0
    
    If count = 0 Then
        count = 10
    ElseIf count < 0 Or count > 1000 Then
        Err.Raise Number:=vbObjectError + 1, _
        description:="Los valores permitidos son de 1 hasta 1000 " + _
                                        "para el argumento 'count'"
    End If

        uri = "https://www.googleapis.com/drive/v3/files/generateIds?count=" & count & "&key=" & YourApiKey & " HTTP/1.1"
            
        With http
            .Open gdGet, uri, False
            .setRequestHeader "Authorization", "Bearer " & YourAccessToken
            .setRequestHeader "Accept", "application/json"
            .send
            numberStatus = .Status
            Set itemJson = JsonConverter.ParseJson(.responseText)
        End With

            index = itemJson("ids").count - 1
            ReDim arrStr(0 To index)
            
            For i = 0 To index
                arrStr(i) = itemJson("ids")(i + 1)
            Next i
        
        GenerateId = arrStr

End Function

Public Function GetMeta(ByVal fileId As String, ByVal fields As String) As String
    
    numberStatus = 0
    
    fields = CodeUri(fields)
    uri = "https://www.googleapis.com/drive/v3/files/" & fileId & "?fields=" & fields & "&key=" & YourApiKey & " HTTP/1.1"
        
    With http
        .Open gdGet, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        
        numberStatus = .Status
    End With
    
    GetMeta = http.responseText
        
End Function

Private Function FileDownload(ByVal fileId As String) As Byte()
        
    numberStatus = 0
    
    uri = "https://www.googleapis.com/drive/v3/files/" & fileId & "?alt=media&key=" & YourApiKey & " HTTP/1.1"
        
    With http
        .Open gdGet, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        numberStatus = .Status
    End With
    
    FileDownload = http.responseBody

End Function


Public Sub FileDownload2(ByVal fileId As String)
    
    Dim itemJson As Dictionary
    Dim uriDownload As String
    Dim contents As String
    Dim webContentLink As String
    
    On Error GoTo Cath
    
    webContentLink = "webContentLink"
    contents = GetMeta(fileId, webContentLink)
    
    Set itemJson = JsonConverter.ParseJson(contents)
    
    uriDownload = itemJson("webContentLink")
    
    Shell ("cmd /c start chrome.exe " & """" & uriDownload & """"), vbHide
    
    Set itemJson = Nothing
    Exit Sub
    
Cath:
   MsgBox Err.Number & vbNewLine & Err.description, vbCritical, APP_NAME
   Set itemJson = Nothing
   On Error GoTo 0
End Sub

Public Function FileList(ByVal q As String, ByVal fields As String, ByVal pageSize As Integer) As String
    
    numberStatus = 0
    
    q = CodeUri(q)
    fields = CodeUri(fields)
    
    uri = "https://www.googleapis.com/drive/v3/files?corpora=user&pageSize=" & pageSize & "&q=" & q & "&fields=files(" & fields & ")&key=" & YourApiKey & "HTTP/1.1"
    
    With http
        .Open gdGet, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        
        numberStatus = .Status
    End With
    
    FileList = http.responseText

End Function
Private Function BinaryCode(ByVal path As String) As Byte()
    
    Const adTypeBinary = 1
    
    ' Lectura en binario
    Dim aStream As New ADODB.Stream
    
    With aStream
        .Type = adTypeBinary
        .Open
        .LoadFromFile (path)
        BinaryCode = .Read()
    End With
    
    Set aStream = Nothing
    
End Function

Private Function Base64Code(ByRef arr() As Byte) As String
    
    'conversión a base64
    'por qué base64?
    '* la codificación con ests juegos de caracteres garantiza la entrega del recurso al servidor
    '*Otro detalle importante es que el cuerpo de petición post  con multipart/related no acepta codificación en binario
    '*La conversióna esta base aumenta en poco más de 30% los bytes en viados.
    
    Dim XML As MSXML2.DOMDocument60
    Dim DocElem As MSXML2.IXMLDOMElement
    
    
    Set XML = New MSXML2.DOMDocument60
    Set DocElem = XML.createElement("Base64Data")
    DocElem.DataType = "bin.base64"
    
    DocElem.nodeTypedValue = arr

    Base64Code = DocElem.Text

    Set XML = Nothing
    Set DocElem = Nothing

End Function
Private Function EncodeBase64(ByVal path As String) As String

    Dim arr() As Byte
    Dim base64 As String
        
    arr = BinaryCode(path)
    base64 = Base64Code(arr)
        
    EncodeBase64 = base64
    
    Erase arr

End Function

Private Function GetUriResumable(ByVal IsUploadChunk As Boolean, _
                                                                                ByVal pathFile As String, _
                                                                                ByVal parents As String, _
                                                                                ByVal mimeType As String) As String
    
    Dim strJson As String
    Dim name As Variant
    
    numberStatus = 0
    
    Rem----------------------meta()
    name = Split(pathFile, "\")
    name = name(UBound(name))
    
'    If mimeType = Empty Then mimeType = "application/octet-stream"
    Rem--------------------fin--meta()

    
    Rem-----------------json
     strJson = "{" & vbNewLine & _
                            """name"":" + """" + name + """" + "," & vbNewLine & _
                            """mimeType"":" + """" + mimeType + """" + "," & vbNewLine & _
                            """parents"":[""" + parents + """]" & vbNewLine & _
                            "}"
    Rem------------------fin-json
    
    
    uri = "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&key=" & YourApiKey & " HTTP/1.1"

    With http
        .Open gdPost, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
'        .setRequestHeader "X-Upload-Content-Length", FileLen(meta(0))
        .setRequestHeader "X-Upload-Content-Type", mimeType
        .setRequestHeader "Content-Type", "application/json; charset=UTF-8"
        
        If Not IsUploadChunk Then
            .setRequestHeader "Content-Lenght", FileLen(pathFile)
        End If
        
        .send (strJson)
        numberStatus = .Status
    End With

    GetUriResumable = http.getResponseHeader("Location")
End Function

Private Function UploadOne(ByVal uri As String, ByVal path As String) As String
    Rem PASO 2.1: carga solictud única-veremos la primera forma de carga mediante una sola solicitud
    'Por qué ?
    'Sí, nuestra conexión a internet es buena y no necesitamos avisar al usario de la carga
    'puedes usar este método.
    
    Dim base64 As String
    Dim arr() As Byte
    
    numberStatus = 0
     arr = BinaryCode(path)
     
    With http
        .Open gdPut, uri
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "X-Upload-Content-Length", FileLen(path)
        .send (arr)
        numberStatus = http.Status
    End With

    UploadOne = http.responseText
    Erase arr

End Function

Public Sub ConnectionService(ByVal oFlowOauth As FlowOauth)
    
    YourAccessToken = oFlowOauth.GetTokenAccess()
    YourApiKey = oFlowOauth.GetApiKey()
    
End Sub


