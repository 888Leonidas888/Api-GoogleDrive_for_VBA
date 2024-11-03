VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GoogleDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const SERVICE_END_POINT = "https://www.googleapis.com/drive/v3/files"
Private Const SERVICE_END_POINT_UPLOAD = "https://www.googleapis.com/upload/drive/v3/files"

Private Enum ERR_DRIVE
    COPY_FAILED = 6000
    CREATE_FAILED = 6001
    DELETE_FAILED = 6002
    EMPTY_TRASH_FAILED = 6003
    EXPORT_FAILED = 6004
    GENERATED_IDS_FAILED = 6005
    METADATA_FAILED = 6006
    LIST_FAILED = 6007
    LIST_LABELS_FAILED = 6008
    DOWNLOAD_WEB_CONTENT_LINK_FAILED = 6009
    DOWNLOAD_BODY_CONTENT_FAILED = 6010
    UPLOAD_MEDIA_FAILED = 6011
    UPLOAD_MULTIPART_FAILED = 6012
    UPLOAD_RESUMEABLE_FAILED = 6013
    UPDATE_FAILED = 6014
    CREATE_RESOURCE_FAILED = 6100
    DIRECTORY_NOT_FOUN_FAILD = 6101
    FILE_NOT_EXISTS_FAILD = 6102
End Enum

Private Type data
    apiKey As String
    accessToken As String
    status As Integer
End Type

Private This As data

Public Property Get operation() As Integer
    operation = This.status
End Property

Public Sub connectionService(ByRef oFlowOauth As FlowOauth)
    With oFlowOauth
        This.accessToken = .GetTokenAccess()
        This.apiKey = .GetApiKey()
    End With
End Sub

Public Function copy(ByVal fileID As String, ByVal fileObject As String, _
                    Optional queryParameters As Dictionary) As String

    Dim pathParameters As String
    Dim url As String
    Dim headers As New Dictionary
    
    On erro GoTo Catch
    
    headers.Add "Content-Type", "application/json"
    pathParameters = Utils.fstring("/{0}/copy", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    
    copy = request(HTTP_POST, url, fileObject, headers).responseText
    
    Exit Function
    
Catch:
    Err.Raise COPY_FAILED, Description:="Failed to copy file."
End Function

Public Function uploadMedia(Optional pathFile As String, Optional mimeType = "application/octet-stream") As String
    
    Dim url As String
    Dim headers As Dictionary
    Dim queryParameters As New Dictionary
    Dim fileObject As Dictionary
    Dim buffer() As Byte
    Dim endPointUpload As Boolean
    Dim body As String
    
    If Trim(pathFile) <> Empty Then
        If existsFile(pathFile) Then
            Set headers = New Dictionary
            
            With headers
                .Add "Content-Type", mimeType
                .Add "Content-Length", FileLen(pathFile)
                .Add "Content-Transfer-Encoding", "binary"
            End With
            
            endPointUpload = True
            
            buffer = sourceToBinary(pathFile)
        Else
            Err.Raise FILE_NOT_EXISTS_FAILD, , "File not exists."
        End If
    Else
        Set fileObject = New Dictionary
        With fileObject
            .Add "mimeType", "application/vnd.google-apps.folder"
            .Add "parents", "root"
        End With

        body = JsonConverter.ConvertToJson(fileObject)
    End If
    
    On erro GoTo Catch
    
    queryParameters.Add "uploadType", "media"
    
    url = createQueryParameters(queryParameters:=queryParameters, endPointUpload:=endPointUpload)
    
    uploadMedia = request(HTTP_POST, url, IIf(body <> Empty, body, buffer), headers).responseText
    
    Exit Function
    
Catch:
    Err.Raise UPLOAD_MEDIA_FAILED, Description:="Upload media failed."
End Function
Public Function uploadMultipart(ByVal pathFile As String, ByRef fileObject As Dictionary) As String
    
    Dim url As String
    Dim body As String
    Dim boundary As String
    Dim headers As New Dictionary
    Dim queryParameters As New Dictionary
    
    If Not existsFile(pathFile) Then Err.Raise FILE_NOT_EXISTS_FAILD, , "File not exists."
    
    On Error GoTo Catch
    
    boundary = generateString(15)
    
    With headers
        .Add "Content-Type", "multipart/related; boundary=" + boundary
        .Add "Content-Lenght", FileLen(pathFile)
    End With
    
    body = createParteRelated(pathFile, boundary, fileObject)
    queryParameters.Add "uploadType", "multipart"
    url = createQueryParameters(queryParameters:=queryParameters, endPointUpload:=True)
    
    uploadMultipart = request(HTTP_POST, url, body, headers).responseText
    
    Exit Function
    
Catch:
    Err.Raise UPLOAD_MULTIPART_FAILED, , "Upload multipart failed."
End Function
Public Function delete(ByVal fileID As String) As String

    Dim pathParemeters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParemeters = "/" & fileID
    url = createQueryParameters(pathParemeters)
    delete = request(HTTP_DELETE, url).responseText
    
    Exit Function
Catch:
    Err.Raise DELETE_FAILED, , "Failed to delete faile."
End Function

Public Function emptyTrash(Optional queryParameters As Dictionary) As String

    Dim pathParameters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParameters = "/trash"
    url = createQueryParameters(pathParameters, queryParameters)
    emptyTrash = request(HTTP_DELETE, url).responseText
    
    Exit Function
Catch:
    Err.Raise EMPTY_TRASH_FAILED, , "Failed to clear trash."
End Function
Public Function export(ByVal fileID As String, ByRef queryParameters As Dictionary, ByVal pathFile As String) As Boolean
    
    Dim pathParameters As String
    Dim url As String
    Dim buffer() As Byte
    
    On Error GoTo Catch
    
    pathParameters = fstring("/{0}/export", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    buffer = request(HTTP_GET, url).responseBody

    If createResource(buffer, pathFile) Then export = True
    
    Exit Function
Catch:
    Err.Raise EXPORT_FAILED, , "Failed to export resource."
End Function
Public Function generateIds(Optional queryParameters As Dictionary) As Collection

    Dim pathParameters As String
    Dim url As String
    Dim json As String
    Dim ids As Collection
    
    On Error GoTo Catch
    
    pathParameters = "/generateIds"
    url = createQueryParameters(pathParameters, queryParameters)
    
    json = request(HTTP_GET, url).responseText
    
    If Me.operation = SUCCESSFUL Then: Set ids = extractIds(json)
    
    Set generateIds = ids
    
    Exit Function
Catch:
    Err.Raise GENERATED_IDS_FAILED, , "Failed to generate ids."
End Function
Public Function getMetadata(ByVal fileID As String, Optional queryParameters As Dictionary) As String

    Dim pathParameters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParameters = "/" + fileID
    url = createQueryParameters(pathParameters, queryParameters)
    getMetadata = request(HTTP_GET, url).responseText
    
    Exit Function
Catch:
    Err.Raise METADATA_FAILED, , "Failed to get metadata."
End Function
Public Function list(Optional queryParameters As Dictionary) As String
    Dim url As String
    
    On Error GoTo Catch
    
    url = createQueryParameters(queryParameters:=queryParameters)
    list = request(HTTP_GET, url).responseText
    
    Exit Function
Catch:
    Err.Raise LIST_FAILED, , "Failed to list files."
End Function
Public Function listLabels(ByVal fileID As String, Optional queryParameters As Dictionary) As String

    Dim pathParameters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParameters = Utils.fstring("/{0}/listLabels", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    
    listLabels = request(HTTP_GET, url).responseText
    
    Exit Function
Catch:
    Err.Raise LIST_LABELS_FAILED, , "Failed to get list labels."
End Function
Public Function update(ByVal fileID As String, Optional fileObject As Dictionary, Optional queryParameters As Dictionary) As String
    
    Dim pathParameters As String
    Dim url As String
    Dim body As String
    
    On Error GoTo Catch
    
    body = JsonConverter.ConvertToJson(fileObject)
    
    pathParameters = fstring("/{0}", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    
    update = request(HTTP_PATCH, url, body).responseText
    
    Exit Function
Catch:
    Err.Raise UPDATE_FAILED, , "Failed to update resource."
End Function
Public Function downloadContentLink(ByVal fileID As String) As Boolean

        Const WEB_CONTENT_LINK = "webContentLink"
        Dim fields As New Dictionary
        Dim json As String
        Dim fileObject As Dictionary
        Dim webContentLink As String
        Dim hasLink As Boolean
            
        On Error GoTo Catch
        
        fields.Add "fields", WEB_CONTENT_LINK
        json = Me.getMetadata(fileID, fields)
        
        Set fileObject = JsonConverter.ParseJson(json)
        
        If fileObject.Exists(WEB_CONTENT_LINK) Then
            webContentLink = fileObject(WEB_CONTENT_LINK)
            Shell (fstring("cmd /c start chrome.exe ""{0}""", webContentLink)), vbHide
            hasLink = True
        End If
        
        downloadContentLink = hasLink
        
        Exit Function
Catch:
    Err.Raise DOWNLOAD_WEB_CONTENT_LINK_FAILED, , "Failed get link to download."
End Function
Public Function download(ByVal fileID As String, ByVal directory As String) As Boolean

    Dim pathParameters As String
    Dim queryParameters As New Dictionary
    Dim nameFile As String
    Dim fullPath As String
    Dim url As String
    Dim buffer() As Byte
    
    If Not existsDirectory(directory) Then Err.Raise DIRECTORY_NOT_FOUN_FAILD, , "Directory does not exist."
    
    On Error GoTo Catch
    
    nameFile = getNameFileForId(fileID)
    
    fullPath = directory + "\" + nameFile
    
    queryParameters.Add "alt", "media"
    pathParameters = fstring("/{0}", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    buffer = request(HTTP_GET, url).responseBody
    download = createResource(buffer, fullPath)
  
    Exit Function
Catch:
    Err.Raise DOWNLOAD_BODY_CONTENT_FAILED, , "Failed to download resource."
End Function
Private Function extractIds(ByVal json As String) As Collection
    
    Dim ids As Collection
    
    Set objJson = JsonConverter.ParseJson(json)
    If objJson.Exists("ids") Then
        Set ids = New Collection
        For i = 1 To objJson("ids").Count
            ids.Add objJson("ids")(i)
        Next i
    End If

    Set extractIds = ids
    
End Function
Private Function createResource(ByRef buffer() As Byte, ByVal pathFile As String) As Boolean

    Dim numFile As Byte

    On Error GoTo Catch

    numFile = FreeFile

    Open pathFile For Binary Access Write As #numFile
        Put #numFile, , buffer
    Close #numFile

    Erase buffer

    createResource = True

    Exit Function
Catch:
    Err.Raise CREATE_RESOURCE_FAILED, , "Failed to create resource."
End Function
Private Function createQueryParameters(Optional pathParameters As String, Optional queryParameters As Dictionary, _
                                        Optional endPointUpload = False) As String

    Dim queryString As String
    Dim endPoint As String
    
    queryString = "?"
    
    If Not queryParameters Is Nothing Then
        For Each key In queryParameters.Keys
            queryString = queryString + key + "=" + URLEncode(queryParameters(key)) + "&"
        Next key
    End If
    
    Select Case endPointUpload
        Case True
            endPoint = SERVICE_END_POINT_UPLOAD
        Case Else
            endPoint = SERVICE_END_POINT
    End Select
    
    endPoint = endPoint + pathParameters
    queryString = queryString + "key=" + This.apiKey + " HTTP/1.1"
    createQueryParameters = endPoint + queryString

End Function
Private Function getNameFileForId(ByVal fileID As String) As String
    
    Dim nameFile As String
    Dim json As String
    Dim fileObject As Dictionary
    Dim queryParameters As New Dictionary
    
    queryParameters.Add "fields", "name"
    json = Me.getMetadata(fileID, queryParameters)
    
    Set fileObject = JsonConverter.ParseJson(json)
    nameFile = IIf(fileObject.Exists("name"), fileObject("name"), "unknown")
    
    getNameFileForId = nameFile
    
    Set fileObject = Nothing
    Set queryParameters = Nothing
    
End Function
Private Function request(ByVal method As String, ByVal url As String, _
                        Optional body = Empty, Optional headers As Dictionary) As MSXML2.ServerXMLHTTP60
    
    Dim http As New MSXML2.ServerXMLHTTP60
    
    On Error GoTo Catch
    
    With http
        .Open method, url
        
        .setRequestHeader "Authorization", "Bearer " + This.accessToken
        .setRequestHeader "Accept", "application/json"
        
        If Not headers Is Nothing Then
            For Each key In headers.Keys
                .setRequestHeader key, headers(key)
            Next key
        End If
        
        .send body
        
        This.status = .status
        Set request = http
    End With
    
    Exit Function
Catch:
    Err.Raise 6001, Description:="Failed conexion with the server."
End Function
