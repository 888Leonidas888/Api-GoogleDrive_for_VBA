VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FlowOauth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'**************************************************************************
'>>>Desarrollado por: Jhony Escriba Bautista                                                                                ***************
'>>>Correo : jhonny14_1@hotmail.com                                                                                            ***************
'>>>Año de lanzamiento 2022                                                                                                                   ***************
'>>>Version de Oauth 2.0                                                                                                                              ***************
'***************************************************************************

Private Const APP_NAME As String = "Flujo de Oauth 2.0"
Private Const version As String = "1.0"

Private YourApiKey As String
Private YourAccessToken As String

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
Private numberStatus As Integer
Private uri As String

Private Function ExistsFile(ByVal pathFile As String) As Boolean
    
    Dim fso As New Scripting.FileSystemObject

    ExistsFile = fso.FileExists(pathFile)
    
    Set fso = Nothing
    
End Function
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

Private Sub ReadCredentials(ByVal credentials_token As String, ByVal credentials_api_key As String)
    
    Dim itemJson As Dictionary
    Dim contents  As String
    Dim arrCredentialsJson As Variant
    
    arrCredentialsJson = Array(credentials_token, credentials_api_key)
    
    For i = 0 To 1
    
        contents = FileReadContent(arrCredentialsJson(i))
        Set itemJson = JsonConverter.ParseJson(contents)

        Select Case i
            Case Is = 0
                YourAccessToken = itemJson("access_token")
            Case Is = 1
                YourApiKey = itemJson("your_api_key")
        End Select
        
    Next i
    
    Set itemJson = Nothing
    
End Sub

Private Function ValidCredentials() As Boolean
    
    Dim valid As Boolean
    
    uri = "https://www.googleapis.com/drive/v3/files?corpora=user&pageSize=1&q=name%20contains%20'.txt'&fields=files(name)&key=" & YourApiKey & "HTTP/1.1"
    
    With http
        .Open gdGet, uri, False
        .setRequestHeader "Authorization", "Bearer " & YourAccessToken
        .setRequestHeader "Accept", "application/json"
        .send
        
        Select Case .Status
            Case Is = 200
                valid = True
            Case Is = 401
                valid = False
            Case Else
                valid = False
        End Select
    End With
    
    ValidCredentials = valid

End Function
Private Function ConsentScreen(ByVal credentials_path As String, ByVal scope As String) As String

        Dim url As String
        Dim parameters As String
        Dim client_id As String, redirect_uri As String
        Dim response_type As String, include_granted_scopes As String
        Dim access_type As String, state As String, flowName As String

        Dim itemJson As Dictionary
        Dim contents  As String
        
        On Error GoTo Catch
        
        Rem lectura de credenciales y asignación de valores_______________>>>
        contents = FileReadContent(credentials_path)
        Set itemJson = JsonConverter.ParseJson(contents)
        
        url = "https://accounts.google.com/o/oauth2/v2/auth/oauthchooseaccount"
        client_id = itemJson("web")("client_id")
        redirect_uri = itemJson("web")("redirect_uris")(1)
        response_type = "code"
        access_type = "offline"
        state = RandomString(30)
        include_granted_scopes = "true"
        flowName = "GeneralconsentScreen"
        Rem ----------------------------------------------------------------------------------------<<<
        
        Rem Codificar uri ________________________>>>
        redirect_uri = CodeUri(redirect_uri)
        scope = CodeUri(scope)
        Rem ___________________________________<<<

                
        parameters = "?client_id=" + client_id + _
                                        "&redirect_uri=" + redirect_uri + _
                                        "&response_type=" + response_type + _
                                        "&scope=" + scope + _
                                        "&include_granted_scopes=" + include_granted_scopes + _
                                        "&access_type=" + access_type + _
                                        "&state=" + state + _
                                        "&flowName=" + flowName

        ConsentScreen = url + parameters
        
        Set itemJson = Nothing
        Exit Function
    
Catch:
    MsgBox Err.description, vbCritical, APP_NAME, Err.Number
    Debug.Print Err.description
    Set itemJson = Nothing
    On Error GoTo 0
End Function
Private Function ChangeTokenUpDate(ByVal code As String, ByVal credentials As String)
    
    Dim url As String
    Dim parameters As String
    Dim client_id As String, redirect_uri As String
    Dim client_secret As String, grant_type As String
    
    Dim itemJson As Dictionary
    Dim contents  As String
    
    numberStatus = 0
    
     Rem lectura de credenciales y asignación de valores_______________>>>
    contents = FileReadContent(credentials)
    Set itemJson = JsonConverter.ParseJson(contents)
'       Stop
    client_id = itemJson("web")("client_id")
    client_secret = itemJson("web")("client_secret")
    redirect_uri = itemJson("web")("redirect_uris")(1)
    code = code
    grant_type = "authorization_code"
    Rem ----------------------------------------------------------------------------------------<<<
    
    Rem codificación uri___________________________________>>>
        redirect_uri = CodeUri(redirect_uri)
    Rem -----------------------------------------------------------------------------<<<
    
    parameters = "code=" + code + _
                                "&client_id=" + client_id + _
                                "&client_secret=" + client_secret + _
                                "&redirect_uri=" + redirect_uri + _
                                "&grant_type=" + grant_type


    url = "https://oauth2.googleapis.com/token"

    With http
        .Open gdPost, url
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send (parameters)
        numberStatus = .Status
    End With
    
    ChangeTokenUpDate = http.responseText
    
End Function

Private Function RefreshToken(ByVal credentials_path As String, ByVal credentials_token As String) As String

    Dim url As String
    Dim parameters As String
    Dim client_id As String, refresh_token As String
    Dim client_secret As String, grant_type As String
    
    Dim itemJson As Dictionary
    Dim contents  As String
    
    numberStatus = 0
    
     Rem lectura de credenciales y asignación de valores_______________>>>
     '-----------credenciales
    contents = FileReadContent(credentials_path)
    Set itemJson = JsonConverter.ParseJson(contents)

    client_id = itemJson("web")("client_id")
    client_secret = itemJson("web")("client_secret")
    grant_type = "refresh_token"
    
    '-----------refresh token--------------------------
    contents = FileReadContent(credentials_token)
    Set itemJson = JsonConverter.ParseJson(contents)

    refresh_token = itemJson("refresh_token")
    Rem ----------------------------------------------------------------------------------------<<<
    

    parameters = "client_id=" + client_id + _
                                "&client_secret=" + client_secret + _
                                "&refresh_token=" + refresh_token + _
                                "&grant_type=" + grant_type

    url = "https://oauth2.googleapis.com/token"

    With http
        .Open gdPost, url
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send (parameters)
    End With
    
    numberStatus = http.Status

    RefreshToken = http.responseText
    Set itemJson = Nothing
    
End Function

Public Function RevokeToken(ByVal credentials_token As String) As Boolean

    Dim url As String
    Dim token_access As String
    'pruede ser token de acceso(token_access) o el token de actualización(token_refresh)
    
    Dim itemJson As Dictionary
    Dim contents  As String
    
    numberStatus = 0
    
     Rem lectura de credenciales y asignación de valores_______________>>>
     '-----------credenciales
    contents = FileReadContent(credentials_token)
    Set itemJson = JsonConverter.ParseJson(contents)

    token_access = itemJson("refresh_token")
    url = "https://oauth2.googleapis.com/revoke?token=" + token_access
    '---------------------------------------------------------------------------------<<<
    
    
    With http
        .Open gdPost, url, False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .send
        
        If .Status = 200 Then
            RevokeToken = True
        Else
            RevokeToken = False
        End If
        
    End With
    
    Set itemJson = Nothing
    
End Function

Public Function InicializeFlow( _
                                                                    ByVal credentials_token As String, _
                                                                    ByVal credentials_client As String, _
                                                                    ByVal credentials_api_key As String, _
                                                                    ByVal scope As String) _
                                                                    As Boolean
    
    Rem esta función leera tu Api key y token(sí es que existe este último,caso contario creara uno)
    
    Rem los sgtes parametros:
    'credentials_token = Archivo con extensión .json que debes proporcionar, si no existe este archivo,el algortimo lo crea
    'credentials_client = Archivo con extensión .json que se obtiene de la plataforma Google Cloud Platform,
    '----------------------de este archivo nos sirve algunos valores los cuales se usaran para redireccionar al usuario
    '----------------------a la pantalla de consetimiento y poder obtener un token de acceso.
    'credentials_api_key = Este 'valor' se obtiene de Google Cloud Platform (credenciales-->crear credenciales-->>clave de api)
    '----------------------una vez obtenida la clave, crear un archivo con extensión .json (clave - valor) ;la clave será 'YourApiKey'
    '----------------------el valor será la 'clave de api' .NOTA:Importante que se haga de esta forma
    'scope = Son los alcances que se otorgarán a la aplicación.
    
    Dim urlRedirect As String
    Dim pid As Double
    Dim code As String
    Dim contentJson As String
    Dim fileJsonToken As String
    Dim fso As New Scripting.FileSystemObject
    Dim t  As TextStream
    
    Dim message As String
    Dim itemJson As Dictionary
        
    If ExistsFile(credentials_token) Then
        
        ReadCredentials credentials_token, credentials_api_key
        
        Rem probaremos la validez de las credenciales
        If Not (ValidCredentials()) Then
            Rem Sí las credenciales son inválidas refresca el token de acceso y sobrescribe el archivo con la respuesta del servidor
            contentJson = RefreshToken(credentials_client, credentials_token)
'            Debug.Print contentJson
            UpdateFileTokenAccess credentials_token, contentJson
        
            Rem ahora vuleve a leer las credenciales.
            'Por fin Listo para usar...............
            ReadCredentials credentials_token, credentials_api_key
        End If
        
    Else
    
        urlRedirect = ConsentScreen(credentials_client, scope)
        pid = Shell("cmd /c start chrome.exe " & """" & urlRedirect & """", vbHide)
        
        Application.Wait now + TimeValue("00:00:05")
        
        code = Application.InputBox( _
                                                                            "Ingrese el valor del ""code"" que aparece en su navegador >>>" _
                                                                            , APP_NAME, _
                                                                            "code", _
                                                                            Type:=2 _
                                                                            )
        
        contentJson = ChangeTokenUpDate(code, credentials_client)
        
        If numberStatus = 200 Then
            
                Rem Escribir codificación en  fichero
                'el valor true sobre escribe el archivo sí ya existe
                Set t = fso.CreateTextFile(credentials_token, True)
                t.Write (contentJson)
                t.Close
                                   
                  ReadCredentials credentials_token, credentials_api_key
                                   
        Else
            MsgBox "Ocurrio un error,no se pudo crear el " + credentials_token, vbCritical, APP_NAME
        End If
'        Debug.Print contentJson
    End If
      
        InicializeFlow = True
        
        Set fso = Nothing
        Set t = Nothing
        Set itemJson = Nothing

End Function
Private Sub UpdateFileTokenAccess(ByVal credentials_token As String, ByVal contentJson As String)
    
    Dim itemJson As Dictionary
    Dim contents As String
    Dim refresh_token As String
    Dim fso As New Scripting.FileSystemObject
    Dim t As TextStream
        
    contents = FileReadContent(credentials_token)
    Set itemJson = JsonConverter.ParseJson(contents)
    
    refresh_token = itemJson("refresh_token")

    contents = FileReadContent(credentials_token)
    
    contentJson = Replace(contentJson, "{", "{" + vbCrLf + Chr(9) + """refresh_token"":""" + refresh_token + """,")
    
    Set t = fso.CreateTextFile(credentials_token, True)
    t.Write (contentJson)
    t.Close
    
    Set itemJson = Nothing
    Set fso = Nothing
    Set t = Nothing
    
End Sub

Public Function GetTokenAccess() As String
    GetTokenAccess = YourAccessToken
End Function

Public Function GetApiKey() As String
    GetApiKey = YourApiKey
End Function
    






