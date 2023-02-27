VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleDriveService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'start 948 lines
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
Private WithEvents http As HTTPRequestService
Attribute http.VB_VarHelpID = -1

Private Const SERVICE_END_POINT As String = "https://www.googleapis.com/drive/v3/files"
Private Const SERVICE_END_POINT_UPLOAD As String = "https://www.googleapis.com/upload/drive/v3/files"
Private Const APP_NAME As String = "Service Google Drive"
Private Const VERSION As String = "2.0.0.0"

Private Type tGds
    yourApiKey As String
    yourAccessToken As String
    responseText As Variant
    status As Integer
    getHeader As String
End Type

Private This As tGds

Private Function HTTPGoogleDrive(ByVal methods As String, _
                                ByVal url As String, _
                                Optional body = Empty, _
                                Optional dicHeaders As Dictionary, _
                                Optional response As String, _
                                Optional getHeader As String) As Variant

    Dim dicHeaders1 As New Dictionary
    Dim k As Variant
    Dim v As String
    
    Set http = New HTTPRequestService
    
    With dicHeaders1
        .Add "Authorization", "Bearer " & This.yourAccessToken
        .Add "Accept", "application/json"
        .Add "Content-Type", "application/json"
    End With
    
    If Not dicHeaders Is Nothing Then
        For Each k In dicHeaders.Keys
            If dicHeaders1.exists(k) Then
                dicHeaders1.Remove k
                dicHeaders1.Add k, dicHeaders(k)
            End If
        Next k
    End If
    
    If getHeader <> Empty Then http.header = getHeader
    
    This.responseText = http.httpRequest(methods, url, body, dicHeaders1, response)
    
    If getHeader <> Empty Then This.getHeader = http.getHeader()
    
    HTTPGoogleDrive = This.responseText
    
    Set k = Nothing
    Set dicHeaders1 = Nothing
    Set http = Nothing
    
End Function

Private Sub HTTP_forRequestHTTP(ByVal status As Integer)
    This.status = status
End Sub
Public Property Get Operation() As Integer
    Operation = This.status
End Property

Public Sub ConnectionService(ByRef oFlowOauth As FlowOauth)
    This.yourAccessToken = oFlowOauth.GetTokenAccess()
    This.yourApiKey = oFlowOauth.GetApiKey()
End Sub

Public Function DetailsError() As String
       
    Dim dicErrors As Dictionary
    Dim dicError As Dictionary
    Dim dicItem As Dictionary
    Dim messageError As String
    
    Select Case This.responseText
        Case Is = Empty
            Exit Function
        Case Is = "Not Found"
            DetailsError = This.responseText
        Case Is >= InStr(1, This.responseText, "{", vbTextCompare)
            Set dicErrors = JsonConverter.ParseJson(This.responseText)
            
            If dicErrors.exists("error") Then
                Set dicError = dicErrors("error")
                
                If dicError.exists("errors") Then
                    Set dicItem = dicError("errors")(1)
                    
                    For i = 0 To dicItem.count - 1
                        messageError = messageError & dicItem.Keys(i) & " : " & dicItem.Items(i) & vbCrLf
                    Next i
                
                    DetailsError = messageError
                End If
            End If
    End Select

    Set dicErrors = Nothing
    Set dicError = Nothing
    Set dicItem = Nothing
    
End Function

Public Function Delete(ByVal fileId As String) As Boolean
    
    Dim url As String

    url = formatString(SERVICE_END_POINT & "/{0}?key={1} HTTP/1.1", fileId, This.yourApiKey)
    
    This.responseText = HTTPGoogleDrive(HT_DELETE, url)
    
    If Me.Operation = GO_NO_CONTENT Then Delete = True

End Function

Public Function CreateResource() As fileResource
    
    Dim fr As New fileResource
    Dim dicResource As Dictionary
    
    Set dicResource = JsonConverter.ParseJson(This.responseText)
    
    If dicResource.exists("kind") Then
        With fr
            .kind = dicResource("kind")
            .mimeType = dicResource("mimeType")
            .id = dicResource("id")
            .name = dicResource("name")
        End With
    End If

    Set CreateResource = fr
    
End Function

Public Function Copy(ByVal fileId As String, _
                        Optional parentTarget As String, _
                        Optional description) As Boolean
        
    Dim body As String, url As String
    
    url = formatString(SERVICE_END_POINT & "/{0}/copy?key={1} HTTP/1.1", fileId, This.yourApiKey)
    
    If IsMissing(description) Then description = "Copy GoogleDrive for VBA"
    
    body = "{'parents':['{0}'],'description':'{1}'}"
    
    If parentTarget = Empty Then body = "{'parents':[{0}],'description':'{1}'}"
    
    body = formatString(Replace(body, "'", """"), parentTarget, description)
                                
    This.responseText = HTTPGoogleDrive(HT_POST, url, body)
                                
    If Me.Operation = GO_SUCCESSFUL Then Copy = True

End Function

Public Function NewFolder(Optional name As String, Optional parents As String) As Boolean

    Const MIMETYPE_FOLDER As String = "application/vnd.google-apps.folder"
    Dim url As String
    Dim body As String
    
    url = formatString(SERVICE_END_POINT + "?key={0} HTTP/1.1", This.yourApiKey)
    
    body = "{'mimeType':'{0}','name':'{1}','parents':['{2}']}"
    
    If parents = Empty Then body = "{'mimeType':'{0}','name':'{1}','parents':[{2}]}"
    
    body = formatString(Replace(body, "'", """"), MIMETYPE_FOLDER, name, parents)
    
    This.responseText = HTTPGoogleDrive(HT_POST, url, body)
    
    If Me.Operation = GO_SUCCESSFUL Then NewFolder = True

End Function

Public Function Update(ByVal fileId As String, Optional json As String) As Boolean
 
    Dim url As String
    
    url = formatString(SERVICE_END_POINT + "/{0}?key={1} HTTP/1.1", fileId, This.yourApiKey)
    
    This.responseText = HTTPGoogleDrive(HT_PATCH, url, json)

    If Me.Operation = GO_SUCCESSFUL Then Update = True
    
End Function

Public Function EmptyTrash() As Boolean
    Rem esta función vacia la papelera de forma irreversible;
    Rem si la solictud ha sido exitosa devolvera un codigo 204 de respuesta y un cuerpo vacio

    Dim url As String
    
    url = formatString(SERVICE_END_POINT + "/trash?key={0} HTTP/1.1", This.yourApiKey)
    
    This.responseText = HTTPGoogleDrive(HT_DELETE, url)
    
    If Me.Operation = GO_NO_CONTENT Then EmptyTrash = True
    
End Function
Public Function GenerateId(Optional count As Integer) As Collection
    
    Rem Genera un  serie de Id  dependiendo el argumento
    '@count = acepta un entero de 1 hasta 1000,predeterminado es 10
    'retorna un obj collection
    
    Dim url As String

    If count >= 1 And count <= 1000 Then
        url = formatString(SERVICE_END_POINT + "/generateIds?count={0}&key={1} HTTP/1.1", count, This.yourApiKey)
    Else
        url = formatString(SERVICE_END_POINT + "/generateIds?key={0} HTTP/1.1", This.yourApiKey)
    End If
    
    This.responseText = HTTPGoogleDrive(HT_GET, url)
    
    If Me.Operation = GO_SUCCESSFUL Then
    
        Dim idCollection As New Collection
        Dim idsDic As Dictionary
           
        Set idsDic = JsonConverter.ParseJson(This.responseText)
        
        If idsDic.exists("ids") Then
            For i = 1 To idsDic("ids").count
                idCollection.Add idsDic("ids")(i)
            Next i
        End If
        
        Set GenerateId = idCollection
    Else
        Set GenerateId = Nothing
    End If
    
End Function
Public Function GetFields(ByVal fileId As String, Optional fields As String) As String
    'antes GetMeta
    
    Dim url As String
    
    If fields <> Empty Then
        fields = codificationUrl(fields)
        url = formatString(SERVICE_END_POINT + "/{0}?fields={1}&key={2} HTTP/1.1", fileId, fields, This.yourApiKey)
    Else
        url = formatString(SERVICE_END_POINT + "/{0}?key={1} HTTP/1.1", fileId, This.yourApiKey)
    End If
    
    This.responseText = HTTPGoogleDrive(HT_GET, url)
    
    GetFields = This.responseText
    
End Function
Public Function List(ByVal q As String, ByVal fields As String, ByVal pageSize As Integer) As String
    
    Dim url As String
    
    q = codificationUrl(q)
    fields = codificationUrl(fields)
    
    url = formatString(SERVICE_END_POINT + "?pageSize={0}&q={1}&fields={2}&key={3} HTTP/1.1", pageSize, q, fields, This.yourApiKey)
    
    This.responseText = HTTPGoogleDrive(HT_GET, url)
    
    List = This.responseText

End Function

Public Function DownloadExport(ByVal fileId As String, _
                            ByVal mimeType As String, _
                            ByVal pathTarget As String, _
                            ByVal nameFile As String) As Boolean
    Rem use esta función paar exportar un documento del Workpace de google a un formato disponible,
    'para ver a que formatos se pueden exportar puede consultar con la función GetMeta o FileList
    'exportLinks = este campo indica los formatos disponibles para este mimeType application/vnd.google-apps.document
    'Para ver los mimeType correspondientes consulte el siguiente enlace:
    'https://developers.google.com/drive/api/guides/ref-export-formats
    'Apriori debe indicar el nombre + la extensión del archivo
    
    Dim buffer() As Byte
    Dim numFile As Byte
    Dim fullPath As String
 
    On Error GoTo Cath
    
    fullPath = pathTarget + "\" + nameFile
    numFile = FreeFile
    buffer = google_fileExport(fileId, mimeType)
    
    Open fullPath For Binary Access Write As #numFile
        Put #numFile, , buffer
    Close #numFile
    
    Erase buffer
    
    DownloadExport = True
    Exit Function
    
Cath:
    DownloadExport = False
    
End Function

Public Function Download2(ByVal fileId As String) As Boolean
    
    Const WEB_CONTENT_LINK As String = "webContentLink"
    Dim dicResponse As Dictionary
    Dim content As String
    Dim webContentLink As String
    
    On Error GoTo Cath
    
    content = Me.GetFields(fileId, WEB_CONTENT_LINK)
    Set dicResponse = JsonConverter.ParseJson(content)
    
    If dicResponse.exists(WEB_CONTENT_LINK) Then
        webContentLink = dicResponse(WEB_CONTENT_LINK)
        Shell ("cmd /c start chrome.exe " & """" & webContentLink & """"), vbHide
    Else
        MsgBox "No existe etiqueta de descarga", vbCritical, APP_NAME
    End If
    
    Set dicResponse = Nothing
    
    Download2 = True
    
    Exit Function
    
Cath:
    Download2 = False

End Function
Public Function Download3(ByVal fileId As String, ByVal pathTarget As String) As Boolean
    
    Dim nameFile As String
    Dim response As String
    Dim responseDic As Dictionary
    Dim fullPath As String
    Dim buffer() As Byte
    Dim numFile As Byte
    
    On Error GoTo Cath
    
    response = Me.GetFields(fileId, "name")
    
    Set responseDic = JsonConverter.ParseJson(response)
    nameFile = responseDic("name")
    
    fullPath = pathTarget + "\" + nameFile
    buffer = google_fileDownload(fileId)
   
    numFile = FreeFile
    
    Open fullPath For Binary Access Write As #numFile
        Put #numFile, , buffer
    Close #numFile
    
    Download3 = True
    
    Exit Function
    
Cath:
    Download3 = False

End Function
Public Function UpLoadResumableSingle(ByVal parent As String, _
                                    ByVal pathFile As String, _
                                    Optional mimeType = "application/octet-stream") As String
 
    Dim url  As String
    Dim header As New Dictionary
    Dim buffer() As Byte
    Dim numFile As Byte
    
    header.Add "Content-Lenght", FileLen(pathFile)
    
    url = google_getUriResumable(pathFile, parent, mimeType)
    buffer = google_convertFileToBinary(pathFile)
    
    If Me.Operation = GO_SUCCESSFUL Then
        This.responseText = HTTPGoogleDrive(HT_POST, url, buffer, header)
        UpLoadResumableSingle = This.responseText
    End If
    
End Function
Public Function UploadMultipart(ByVal pathFile As String, _
                                    ByVal parent As String, _
                                    Optional mimeType = "application/octet-stream") As String
    'FUNCION
        'Use esta función cuando se requiera subir archivos <=5mb y la conexión sea estable
        'La carga se hará en una sola petición.Carga tanto el recurso como los metadatos.
        
        'El nombre de 'multiparte' hace mención al body, se envia tanto los metadatos como el recurso
        'convertido en base64.
    
    'PARAM
        'PATHFILE(string):debe indicarse la ruta del recurso que desea subirse <=5MB.
        'PARENT(string):el directorio donde se subirá el recurso.
        'MIMETYPE(string):Debe indicarse el tipo de mime por defecto será 'application/octet-stream'.
    'RETURN(string-json): entrega un recurso de archivo que puede posteriormente ser convertido a un objeto FileResource
    
    Dim url As String
    Dim related As String
    Dim boundary As String
    Dim headers As New Dictionary
    
    boundary = generateString(15)
    headers.Add "Content-Type", "multipart/related; boundary=" + boundary
    headers.Add "Content-Lenght", FileLen(pathFile)
   
    related = google_createParteRelated(pathFile, parent, mimeType, boundary)

    url = formatString(SERVICE_END_POINT_UPLOAD + "?uploadType=multipart&key={0} HTTP/1.1", This.yourApiKey)
    
    This.responseText = HTTPGoogleDrive(HT_POST, url, related, headers)
    
    UploadMultipart = This.responseText
    Set headers = Nothing
    
End Function
Public Function UploadMedia(ByVal pathFile As String, Optional mimeType = "application/octet-stream") As String
    
     'FUNCION
        'Use esta función cuando se requiera subir archivos <=5mb y la conexión sea estable
        'La carga se hará en una sola petición y no incluye metadatos, estos pueden subirse usando un función
        'Update.
        
        'El contenido es enviado en BYTES al servidor.
    
    'PARAM
        'PATHFILE(string):debe indicarse la ruta del recurso que desea subirse <=5MB.
        'MIMETYPE(string):Debe indicarse el tipo de mime por defecto será 'application/octet-stream'.
    'RETURN(string-json): entrega un recurso de archivo que puede posteriormente ser convertido a un objeto FileResource
    
    Dim url As String
    Dim buffer() As Byte
    Dim headers As New Dictionary
    Dim sizeFile As Double
    
    sizeFile = FileLen(pathFile)
    
    url = formatString(SERVICE_END_POINT_UPLOAD + "?uploadType=media&key={0} HTTP/1.1", This.yourApiKey)
      
    buffer = google_convertFileToBinary(pathFile)
    
    With headers
        .Add "Content-Type", mimeType
        .Add "Content-Length", sizeFile
        .Add "Content-Transfer-Encoding", "binary"
    End With
    
    This.responseText = HTTPGoogleDrive(HT_POST, url, buffer, headers)
    
    Erase buffer
    
    UploadMedia = This.responseText
    
End Function

'@function private
Private Function google_convertFileToBinary(ByVal pathFile As String) As Byte()
    
    Const adTypeBinary = 1
    
    ' Lectura en binario
    Dim aStream As New ADODB.Stream
    
    With aStream
        .Type = adTypeBinary
        .Open
        .LoadFromFile (pathFile)
        google_convertFileToBinary = .Read()
    End With
    
    Set aStream = Nothing
    
End Function
Private Function google_convertFileToBinary2(ByVal pathFile As String) As Byte()
    
    Dim numFile As Byte
    Dim buffer() As Byte

    numFile = FreeFile
    
    Open pathFile For Binary Access Read As #numFile
        ReDim buffer(LOF(numFile))
        Get #numFile, , buffer
    Close #numFile
    
    google_convertFileToBinary2 = buffer
    
End Function
Private Function google_convertToBase64(ByRef arr() As Byte) As String
    
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

    google_convertToBase64 = DocElem.text

    Set XML = Nothing
    Set DocElem = Nothing

End Function
Private Function google_enCodeBase64(ByVal pathFile As String) As String

    Dim buffer() As Byte
    Dim base64 As String
        
    buffer = google_convertFileToBinary(pathFile)
    base64 = google_convertToBase64(buffer)
        
    google_enCodeBase64 = base64
    
    Erase buffer

End Function

Private Function google_fileDownload(ByVal fileId As String) As Byte()
        
    Dim url As String
    Dim buffer() As Byte
    
    url = formatString(SERVICE_END_POINT + "/{0}?alt=media&key={1} HTTP/1.1", fileId, This.yourApiKey)
    
    buffer = HTTPGoogleDrive(HT_GET, url, response:="body")
    
    google_fileDownload = buffer

End Function
Private Function google_fileExport(ByVal fileId As String, ByVal mimeType As String) As Byte()
    Rem Exporta un documento de Google Workspace al tipo MIME solicitado
    'y devuelve el contenido de bytes exportado. Tenga en cuenta que el contenido
    'exportado está limitado a 10 MB.
    
    Rem Si tiene éxito, este método devuelve el contenido del archivo como bytes.
    'Para obtener una lista de documentos de Google Workspace y formatos de
    'conversión de tipo MIME, consulte Documentos de Google Workspace y
    'tipos MIME de exportación correspondientes .
    'https://developers.google.com/drive/api/guides/ref-export-formats
    
    Dim buffer() As Byte
    
    mimeType = codificationUrl(mimeType)
    url = formatString(SERVICE_END_POINT + "/{0}/export?mimeType={1}&key={2} HTTP/1.1", fileId, mimeType, This.yourApiKey)
    buffer = HTTPGoogleDrive(HT_GET, url, response:="body")
    google_fileExport = buffer
            
End Function
Private Function google_createParteRelated(ByVal filePath As String, _
                                            ByVal parents As String, _
                                            ByVal mimeType As String, _
                                            ByVal boundary As String) As String
        
    Dim related As String
    Dim body As String
    Dim start_boundary As String, finish_boundary As String
    Dim fileName
    Dim base64 As String
    
    base64 = google_enCodeBase64(filePath)
    start_boundary = "--" + boundary
    finish_boundary = start_boundary + "--"
    
    fileName = Split(filePath, "\")
    fileName = fileName(UBound(fileName))

    body = formatString(Replace("{'name':'{0}','mimeType':'{1}','parents':['{2}'],'description':'{3}'}", "'", """"), fileName, mimeType, parents, "Uploaded from VBA")

    related = formatString( _
            "{0}{1}Content-Type: application/json; charset=UTF-8{1}{1}{2}{1}{0}{1}Content-Type: {3}{1}Content-Transfer-Encoding: base64{1}{1}{4}{1}{5}", _
            start_boundary, vbNewLine, body, mimeType, base64, finish_boundary)

    google_createParteRelated = related

End Function
Private Function google_getUriResumable(ByVal pathFile As String, _
                                        ByVal parent As String, _
                                        ByVal mimeType As String) As String
                                
    Dim url As String
    Dim headers As New Dictionary
    Dim body As String
    Dim sizeFile As Double
    Dim nameFile As Variant
    
    nameFile = Split(pathFile, "\")
    nameFile = nameFile(UBound(nameFile))
    sizeFile = FileLen(pathFile)
    
    url = formatString(SERVICE_END_POINT_UPLOAD + "?uploadType=resumable&key={0} HTTP/1.1", This.yourApiKey)
    
    body = formatString( _
            Replace("{'name':'{0}','parents':['{1}'],'mimeType':'{2}'}", "'", """"), _
            nameFile, parent, mimeType)
    
    With headers
        .Add "X-Upload-Content-Type", mimeType
        .Add "Content-Type", "application/json; charset=UTF-8"
        .Add "Content-Length", sizeFile
    End With
    
    HTTPGoogleDrive HT_POST, url, body, headers, getHeader:="Location"
    
    google_getUriResumable = This.getHeader
    
End Function

