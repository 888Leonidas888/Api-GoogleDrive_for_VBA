Attribute VB_Name = "main"
Sub delete_file()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim fileId As String
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    fileId = "19IXkHGi6fhchk-jQ6Qcb2TsLWDo90yuU"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Debug.Print .Delete(fileId)
        Rem el método Operation devolverá una constante
        'GO_NO_CONTENT
        
        If .Operation = GO_NO_CONTENT Then
            Debug.Print "file delete"
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub

Sub copy_file()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim fileId As String
    Dim parents As String
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    fileId = "1BwacJSDshnTOW05POsWLtqpSJnaotaEV"
    parents = "1Pc4IfuV9UGVCDFD9elxGcaHWIOzwmvx5"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Debug.Print .Copy(fileId, parents)
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        
        If .Operation = GO_SUCCESSFUL Then
            Dim fr As fileResource
            
            Set fr = .CreateResource()
            
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
            
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub

Sub deleteTrash()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Rem una llamada exitosa debolvera un True
        Debug.Print .EmptyTrash()
        
        Rem el método Operation devolverá una constante
        'GO_NO_CONTENT
        If .Operation = GO_NO_CONTENT Then
           Debug.Print "papelera limpia"
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub

Sub download_for_link()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim fileId As String

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    fileId = "1U5qZ5Jsqtycbw34gLvABdomg7HoBujSR"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Debug.Print .Download2(fileId)
    End With
    
End Sub

Sub download_3()
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    
    Dim fileId As String
    Dim pathTarget As String
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    fileId = "1P4hmXIRWlqa1Cp0YIHCPJ5RsZRCoMscJ"
    pathTarget = ThisWorkbook.path + "\multimedia"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Debug.Print .Download3(fileId, pathTarget)
    End With
End Sub

Sub generteId()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    
    Dim c As Collection

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Set c = .GenerateId()
                
        If Not c Is Nothing Then
            For i = 1 To c.count
                Debug.Print c.item(i)
            Next i
        Else
            Debug.Print "No se genero Ids"
        End If
    End With
    
End Sub
Sub GetFields()
        
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim fileId As String
    Dim response As String
    Dim fields As String

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    fileId = "1P4hmXIRWlqa1Cp0YIHCPJ5RsZRCoMscJ"
    fields = "*"

    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        
        Rem puedes recuperar el texto para crear un objeto
        response = .GetFields(fileId, fields)
        
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print response
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub

Sub List()
    
    'Vea los siguientes enlaces para aprender hacer consultas
    
    'https://developers.google.com/drive/api/v3/reference/files/list
    'https://developers.google.com/drive/api/guides/search-files?hl=es-419#specific
    'https://developers.google.com/drive/api/guides/search-files?hl=es-419#examples
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim response As String
    Dim q As String
    Dim fields As String
    Dim pageSize As Integer

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    
'    q = "name contains 'vba'and trashed = false"
'    q = "modifiedTime > '2023-02-24T12:00:00'"
'    q = "mimeType = 'video/mp4' and trashed = false"
    q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
'    fields = "nextPageToken,kind,incompleteSearch,files(name,id)"
'    fields = "files(id,capabilities/canAddChildren)"
    fields = "files(name,id,mimeType)"
    pageSize = 10
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        
        Rem puede s recuperar el texto para crear un objeto
        response = .List(q, fields, pageSize)
        
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print response
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub

Sub NewFolder()
    
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim parents As String
    Dim name As String

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    parents = "1Pc4IfuV9UGVCDFD9elxGcaHWIOzwmvx5"
    name = "other folder"
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Rem tanto el name como el parents son opcionales
        Rem en caso que no se le envie el parent creara en 'Mi Unidad'
        Debug.Print .NewFolder(name, parents)
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        
        If .Operation = GO_SUCCESSFUL Then
            Dim fr As fileResource
            
            Set fr = .CreateResource()
            
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
            
        Else
            Debug.Print .DetailsError
        End If
    End With
    
End Sub
Sub update_file()
    
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    Dim fileId As String
    Dim json As String

    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fileId = "195zrZ9lQW7o2QZ-sLDS5aZOo57sUkZ2L"
    json = Replace("{'name':'actualizado para VBA'}", "'", """")
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        Debug.Print .Update(fileId, json)
        
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        If .Operation = GO_SUCCESSFUL Then
            Dim fr As fileResource
            
            Set fr = .CreateResource()
            
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
            
        Else
            Debug.Print .DetailsError
        End If
    End With
End Sub

Sub upload_media()
    Rem metodo recomendado para archivos de <=5mb, sin metadatos
    
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    
    Dim response As String
    Dim pathFile As String
    Dim parent As String
    
    pathFile = ThisWorkbook.path + "\multimedia\Dua Lipa.mp3"
    parent = "195zrZ9lQW7o2QZ-sLDS5aZOo57sUkZ2L"
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        .UploadMedia pathFile
        
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        If .Operation = GO_SUCCESSFUL Then
             Dim fr As fileResource
            
            Set fr = .CreateResource()
            
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
        Else
            Debug.Print .DetailsError
        End If
    End With
End Sub

Sub upload_multipart()
    Rem metodo recomendado para archivos de <=5mb
    
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    
    Dim response As String
    Dim pathFile As String
    Dim parent As String
    
    pathFile = ThisWorkbook.path + "\multimedia\edificiones.jpg"
    parent = "195zrZ9lQW7o2QZ-sLDS5aZOo57sUkZ2L"
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        .UploadMultipart pathFile, parent
        
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL
        If .Operation = GO_SUCCESSFUL Then
             Dim fr As fileResource
            
            Set fr = .CreateResource()
            
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
        Else
            Debug.Print .DetailsError
        End If
    End With
End Sub

Sub upload_resumable()
    
    Dim credentialsToken As String
    Dim credentialsApikey As String
    
    Dim fo As New FlowOauth
    Dim drive As New GoogleDriveService
    
    Dim response As String
    Dim pathFile As String
    Dim parent As String
    
    pathFile = ThisWorkbook.path + "\multimedia\Dua Lipa.mp3"
    parent = "195zrZ9lQW7o2QZ-sLDS5aZOo57sUkZ2L"
    
    credentialsClient = ThisWorkbook.path + "\credentials\client.json"
    credentialsToken = ThisWorkbook.path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    With drive
        .ConnectionService fo
        .UpLoadResumableSingle parent, pathFile
        Debug.Print Time; " finish"
        Rem el método Operation devolverá una constante
        'GO_SUCCESSFUL or GO_CREATED
        If .Operation = GO_SUCCESSFUL Or .Operation = GO_CREATED Then
            Dim fr As fileResource
            Set fr = .CreateResource()
            With fr
                Debug.Print "Id --> "; .id
                Debug.Print "Name --> "; .name
                Debug.Print "Kind --> "; .kind
                Debug.Print "MimeType --> "; .mimeType
            End With
        Else
            Debug.Print .DetailsError
        End If
    End With
    
    
End Sub

