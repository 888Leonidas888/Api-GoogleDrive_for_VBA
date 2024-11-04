VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlowOauth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum ERR_OAUTH
    HOLA = 6200
    FILE_API_KEY_NOT_FOUND = 6201
    FILE_CLIENT_NOT_FOUND = 6202
    FAILED_READ_CREDENTIALS = 6203
End Enum

Private Type access
    apiKey As String
    accessToken As String
    client_id As String
    redirect_uri As String
    client_secret As String
    refresh_token As String
    status As Integer
End Type

Private Const OAUTH2_END_POINT = "https://accounts.google.com/o/oauth2/v2/auth"
Private Const TOKEN_END_POINT = "https://oauth2.googleapis.com/token"
Private Const REVOKE_TOKEN_END_POINT = "https://oauth2.googleapis.com/revoke"

Private this As access

Public Function getTokenAccess() As String
    getTokenAccess = this.accessToken
End Function

Public Function getApiKey() As String
    getApiKey = this.apiKey
End Function

Public Property Get operation() As Variant
    operation = this.status
End Property

Public Sub InitializeFlow(ByVal credentialsClient As String, _
                                ByVal credentialsToken As String, _
                                ByVal credentialsApikey As String, _
                                ByVal scope As String)
                                
        Dim url As String
        Dim ExcelObj As New Excel.Application
        Dim code As String
        Dim json As String
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; "flow started"
        
        Rem lectura de api_key;client_id;redirect_uris;client_secret
        Call readCredentialsGoogleConsole(credentialsClient, credentialsApikey)
        
        If Utils.existsFile(credentialsToken) Then
            Call verifyingToken(credentialsToken)
        Else
            Debug.Print APP_NAME; " "; Now(); " >>> "; "FILE NOT FOUND " & credentialsToken
            
        Rem #1.- Autenticar y autorizar a al APP;redirección del usuario a la pantalla de consentimiento
            url = urlRedirect(scope)
            Shell fstring("cmd /c start chrome.exe ""{0}""", url), vbHide

        Rem #2.- Obtener código para cambiar
            ExcelObj.Wait Now() + TimeValue("00:00:05")
            code = ExcelObj.InputBox( _
            "Ingrese el valor del parámetro ""code"" que aparece en la URL de su navegador", _
            APP_NAME, Type:=2)
            Set ExcelObj = Nothing
            
            If code = Empty Then Exit Sub
            
        Rem #3.- Cambiar el código para un token de acceso
            json = changeCodeForToken(code)
        
        Rem #4.- Guardar el token de acceso para su posterior uso.
            If Me.operation = SUCCESSFUL Then
                Dim newTokenObject As Dictionary
                Dim currentDate As Double
                
                'agregamos la fecha y la hora menos 5 segundos en la que se escribe el token.
                Set newTokenObject = JsonConverter.ParseJson(json)
                currentDate = CDbl(Date) + CDbl(DateAdd("s", 3599 - 300, Time))
                newTokenObject.Add "current_date", currentDate
                this.accessToken = IIf(newTokenObject.Exists("access_token"), _
                                            newTokenObject("access_token"), "token not found")
                                            
                json = JsonConverter.ConvertToJson(newTokenObject, 2)
                '----------------------------------------------------
                
                writeFile json, credentialsToken
                
                Debug.Print APP_NAME; " "; Now(); " >>> "; "new token generated"
            Else
                Debug.Print APP_NAME; " "; Now(); " >>> "; "error "; json
            End If
        End If
        
End Sub

Public Function RevokeToken(ByVal credentialsToken As String) As Boolean
    Dim body As String
    Dim headers As New Dictionary
    Dim dicToken As Dictionary
    Dim content As String
    
    If Not Utils.existsFile(credentialsToken) Then Exit Function
    
    content = Utils.readFile(credentialsToken)
    Set dicToken = JsonConverter.ParseJson(content)
    this.accessToken = dicToken("access_token")
    
    body = "token=" & this.accessToken
    headers.Add "Content-type", "application/x-www-form-urlencoded"
    request HTTP_POST, REVOKE_TOKEN_END_POINT, body, headers
    
    If Me.operation = SUCCESSFUL Then
        Debug.Print APP_NAME; " "; Now(); " >>> "; "revoked token"
        RevokeToken = True
    Else
        Debug.Print APP_NAME; " "; Now(); " >>> "; "failed token revoke"
        RevokeToken = False
    End If
    Set dicToken = Nothing
    Set headers = Nothing
End Function

Private Sub readCredentialsGoogleConsole(ByVal credentialsClient As String, ByVal credentialsApikey As String)
        
    Dim apiKeyObject As Dictionary
    Dim clientObject As Dictionary
    Dim content As String
    Dim app As String
    
    On Error GoTo Catch
    
    If Utils.existsFile(credentialsApikey) Then
        content = Utils.readFile(credentialsApikey)
        Set apiKeyObject = JsonConverter.ParseJson(content)
        this.apiKey = IIf(apiKeyObject.Exists("your_api_key"), apiKeyObject("your_api_key"), "api_key_not_found")
    Else
        Err.Raise FILE_API_KEY_NOT_FOUND, , "File with api key not found."
    End If
          
    If Utils.existsFile(credentialsClient) Then
        content = Utils.readFile(credentialsClient)
        Set clientObject = JsonConverter.ParseJson(content)
        
        With clientObject
            If .Exists("web") Then app = "web"
            If .Exists("installed") Then app = "installed"
        End With
        
        With this
            .client_id = clientObject(app)("client_id")
            .redirect_uri = clientObject(app)("redirect_uris")(1)
            .redirect_uri = Utils.URLEncode(.redirect_uri)
            .client_secret = clientObject(app)("client_secret")
        End With
    Else
        Err.Raise FILE_CLIENT_NOT_FOUND, , "File client not found."
    End If
    
    Set clientObject = Nothing
    Set apiKeyObject = Nothing
    
    Exit Sub
Catch:
    Err.Raise FAILED_READ_CREDENTIALS, , "Failed to read credentials."
End Sub
Private Function urlRedirect(ByVal scope As String) As String
        
        Dim response_type As String
        Dim access_type As String
        Dim include_granted_scopes As String
        Dim prompt As String
        Dim pathParameters As String
 
        response_type = "code"
        scope = Utils.URLEncode(scope)
        access_type = "offline"
        include_granted_scopes = "true"
        prompt = "consent"
        
        pathParameters = Utils.fstring( _
        "?scope={0}&access_type={1}&include_granted_scopes={2}&response_type={3}&redirect_uri={4}&client_id={5}&prompt={6}", _
                        scope, _
                        access_type, _
                        include_granted_scopes, _
                        response_type, _
                        this.redirect_uri, _
                        this.client_id, prompt)
        
        urlRedirect = OAUTH2_END_POINT & pathParameters
            
End Function

Private Function changeCodeForToken(ByVal code As String) As String
    Rem esta funcion reemplaza a ChangeTokenUpDate
        
    Dim body As String
    Dim headers As New Dictionary
    
    body = fstring( _
            "code={0}&client_id={1}&client_secret={2}&redirect_uri={3}&grant_type={4}", _
            code, _
            this.client_id, _
            this.client_secret, _
            this.redirect_uri, _
            "authorization_code")
            
    headers.Add "Content-Type", "application/x-www-form-urlencoded"
    
    changeCodeForToken = request(HTTP_POST, TOKEN_END_POINT, body, headers).responseText
    
    Set headers = Nothing
    
End Function

Private Sub verifyingToken(ByVal credentialsToken As String)
    
    Dim tokenObject As Dictionary
    Dim error_ As String, error_description_ As String
    Dim rsp As Integer
    Dim content As String
    
    content = Utils.readFile(credentialsToken)
    Set tokenObject = JsonConverter.ParseJson(content)
    
    Debug.Print APP_NAME; " "; Now(); " >>> "; "token verifying"
    
    If tokenObject.Exists("error") Or tokenObject.Exists("error_description") Then
        error_ = tokenObject("error")
        error_description_ = tokenObject("error_description")
        
        rsp = MsgBox(error_ & vbCrLf & error_description_ & vbCrLf & vbCrLf & _
            "Presione ""SI"" para eliminar el archivo con el token,esto generará un error en la respuesta de la API", _
            vbYesNo + vbDefaultButton1 + vbCritical, APP_NAME)
        
        If rsp = vbYes Then Kill credentialsToken
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; error_; "  "; error_description_
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Rem "A DONDE ME VOY"
        Set tokenObject = Nothing
        Exit Sub
        
    ElseIf tokenObject.Exists("access_token") And tokenObject.Exists("token_type") Then
        
        Dim registeredDate As Double
        Dim refresh_token As String
        
        registeredDate = tokenObject("current_date")
        
        If CDbl(Now()) > CDbl(registeredDate) Then
            this.refresh_token = tokenObject("refresh_token")
            
            Debug.Print APP_NAME; " "; Now(); " >>> "; "requests update token"
            Call updateTokenAccess(credentialsToken)
        Else

            this.accessToken = IIf(tokenObject.Exists("access_token"), _
                                            tokenObject("access_token"), "token not found")
            Debug.Print APP_NAME; " "; Now(); " >>> "; "using token_access"
        End If
        
    End If
    
    Set tokenObject = Nothing
    
End Sub

Private Sub updateTokenAccess(Optional credentialsToken As String)
    
    Dim body As String
    Dim headers As New Dictionary
    Dim json As String
    
    With this
        body = fstring( _
            "client_id={0}&client_secret={1}&refresh_token={2}&grant_type=refresh_token", _
            .client_id, .client_secret, .refresh_token)
    End With
    
    headers.Add "Content-Type", "application/x-www-form-urlencoded"
    json = request(HTTP_POST, TOKEN_END_POINT, body, headers).responseText
    
    If Me.operation = SUCCESSFUL Then
        Dim currentDate As Double
        Dim tokenObject As Dictionary
        
        Debug.Print APP_NAME; " "; Now(); " >>> "; "successfully updated token"
        
        'agregamas la fecha y la hora menos 5 segundos en la que se escribe el token.
        currentDate = CDbl(Date) + CDbl(DateAdd("s", 3599 - 300, Time))
      
        Set tokenObject = JsonConverter.ParseJson(json)
        
        With tokenObject
            .Add "current_date", currentDate
            .Add "refresh_token", this.refresh_token
            'agregamos el refresh_token al contenido
        End With
        
        'leemos el token_access
        this.accessToken = IIf(tokenObject.Exists("access_token"), _
                                            tokenObject("access_token"), _
                                            "token not found")
                                            
        json = JsonConverter.ConvertToJson(tokenObject, 2)
        Set tokenObject = Nothing
    Else
        Debug.Print APP_NAME; " "; Now(); " >>> "; "failed update token"
        
    End If
    'guardamos respuesta exitosa o fallida
    Utils.writeFile json, credentialsToken
    Set headers = Nothing
    
End Sub
Private Function request(ByVal method As String, ByVal url As String, _
                        Optional body = Empty, Optional headers As Dictionary) As MSXML2.ServerXMLHTTP60
    
    Dim http As New MSXML2.ServerXMLHTTP60
    
    On Error GoTo Catch
    
    With http
        .Open method, url
        
        If Not headers Is Nothing Then
            For Each key In headers.Keys
                .setRequestHeader key, headers(key)
            Next key
        End If
        
        .send body
        
        this.status = .status
        Set request = http
    End With
    
    Exit Function
Catch:
    Err.Raise 6001, Description:="Failed conexion with the server."
End Function
