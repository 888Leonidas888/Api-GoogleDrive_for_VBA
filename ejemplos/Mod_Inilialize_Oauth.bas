Attribute VB_Name = "Mod_Inilialize_Oauth"
Private Const api_key As String = "C:\Users\JHONY\Desktop\proyecto api google drive\mis credenciales\mi_api_key_616.json"
Private Const client As String = "C:\Users\JHONY\Desktop\proyecto api google drive\mis credenciales\cliente_id.json"
Private Const access_token As String = "C:\Users\JHONY\Desktop\proyecto api google drive\mis credenciales\mi_token.json"
Private Const scope As String = "https://www.googleapis.com/auth/drive"

Sub comenzar_flujo()
    
    Dim oauth As New FlowOauth
    
    oauth.InicializeFlow access_token, client, api_key, scope
    
End Sub
Sub ejemplo()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim response As String
    Dim q As String
    Dim fields As String
    Dim pageSize As String
    
    q = "parents = '1oXIsoy11Zp7--xjQVyqj9mqCQ-b8LuEy'"
    fields = "name,id,mimeType"
    pageSize = 10
    
    
    oauth.InicializeFlow access_token, client, api_key, scope
    
    
    With google
        .ConnectionService oauth
        response = .FileList(q, fields, pageSize)
    End With
    
    Debug.Print response
    
    Set oauth = Nothing
    Set google = Nothing
    
End Sub


