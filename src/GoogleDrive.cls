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

Private Const HTTP_GET = "GET"
Private Const HTTP_POST = "POST"
Private Const HTTP_DELETE = "DELETE"
Private Const HTTP_PATCH = "PATCH"

Private Enum CODE_STATE_RESPONSE_HTTP
    GO_SUCCESSFUL = 200
    GO_CREATED = 201
    GO_NO_CONTENT = 204
    GO_FOUND = 302
    GO_RESUME_INCOMPLETE = 308
    GO_FAILED = 400
    GO_RATE_LIMIT = 403
    GO_NOT_FOUND = 404
    GO_SERVICE_UNAVAILABLE = 503
End Enum

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
    CREATE_RESOURCE_FAILED = 6100
End Enum

Private Type data
    apiKey As String
    accessToken As String
    responseText As String
    status As Integer
End Type

Private This As data

Public Property Get Operation() As Integer
    Operation = This.status
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
    pathParameters = utils.fstring("/{0}/copy", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    
    copy = request(HTTP_POST, url, fileObject, headers)
    
    Exit Function
    
Catch:
    Err.Raise COPY_FAILED, Description:="Failed to copy file."
    
End Function
Public Function create(ByVal body As String) As String

End Function
Public Function delete(ByVal fileID As String) As String

    Dim pathParemeters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParemeters = "/" & fileID
    url = createQueryParameters(pathParemeters)
    delete = request(HTTP_DELETE, url)
    
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
    emptyTrash = request(HTTP_DELETE, url)
    
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
    buffer = request2(HTTP_GET, url)

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
    
    json = request(HTTP_GET, url)
    
    If Me.Operation = GO_SUCCESSFUL Then: Set ids = extractIds(json)
    
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
    getMetadata = request(HTTP_GET, url)
    
    Exit Function
Catch:
    Err.Raise METADATA_FAILED, , "Failed to get metadata."
End Function
Public Function list(Optional queryParameters As Dictionary) As String
    Dim url As String
    
    On Error GoTo Catch
    
    url = createQueryParameters(queryParameters:=queryParameters)
    list = request(HTTP_GET, url)
    
    Exit Function
Catch:
    Err.Raise LIST_FAILED, , "Failed to list files."
End Function
Public Function listLabels(ByVal fileID As String, Optional queryParameters As Dictionary) As String

    Dim pathParameters As String
    Dim url As String
    
    On Error GoTo Catch
    
    pathParameters = utils.fstring("/{0}/listLabels", fileID)
    url = createQueryParameters(pathParameters, queryParameters)
    
    listLabels = request(HTTP_GET, url)
    
    Exit Function
Catch:
    Err.Raise LIST_LABELS_FAILED, , "Failed to get list labels."
End Function
Public Function update()

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

Private Function createQueryParameters(Optional pathParameters As String, Optional queryParameters As Dictionary) As String

    Dim queryString As String
    Dim endPoint As String
    
    queryString = "?"
    
    If Not queryParameters Is Nothing Then
        For Each key In queryParameters.Keys
            queryString = queryString + key + "=" + URLEncode(queryParameters(key)) + "&"
        Next key
    End If
    
    endPoint = SERVICE_END_POINT + pathParameters
    queryString = queryString + "key=" + This.apiKey + " HTTP/1.1"
    createQueryParameters = endPoint + queryString

End Function
Private Function request(ByVal method As String, ByVal url As String, _
                        Optional body = Empty, Optional headers As Dictionary) As String
    
    Dim req As New MSXML2.ServerXMLHTTP60
    
    On Error GoTo Catch
    
    With req
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
        request = .responseText
    End With
    
    Exit Function
Catch:
    Err.Raise 6001, Description:="Failed conexion with the server."
End Function
Private Function request2(ByVal method As String, ByVal url As String, _
                        Optional body = Empty, Optional headers As Dictionary) As Byte()
    
    Dim req As New MSXML2.ServerXMLHTTP60
    
    On Error GoTo Catch
    
    With req
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
        request2 = .responseBody
    End With
    
    Exit Function
Catch:
    Err.Raise 6001, Description:="Failed conexion with the server."
End Function

