Attribute VB_Name = "demo"
Function initOauthFlow() As FlowOauth
    
     
    Dim credentialsClient As String
    Dim credentialsToken As String
    Dim credentialsApikey As String
    Dim fo As New FlowOauth
 
    credentialsClient = ThisWorkbook.Path + "\credentials\clientweb.json"
    credentialsToken = ThisWorkbook.Path + "\credentials\token.json"
    credentialsApikey = ThisWorkbook.Path + "\credentials\apikey.json"
    
    fo.InitializeFlow credentialsClient, credentialsToken, credentialsApikey, OU_SCOPE_DRIVE
    
    Set initOauthFlow = fo

End Function
Sub copy_file()
    
    Dim parents As New Collection
    Dim drive As New GoogleDrive
    Dim fileObject As New Dictionary
    Dim file As String
    Dim fileID As String
    
    On Error GoTo Catch
    
    fileID = "10N8q9lkLIP0np-1X_pYJUBP-lXgaV7qo"
    parents.Add "10SgXNVOUO35QLdg1Fsd_Jb5mugagfm5u"
    
    With fileObject
        .Add "parents", parents
        .Add "description", "prueba método copy"
    End With
    
    file = JsonConverter.ConvertToJson(fileObject)
    
    With drive
        .connectionService initOauthFlow
        result = .copy(fileID, file)
        
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print result
        End If
    End With
    
    Exit Sub
    
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
End Sub
Sub list_file()
    
    Dim drive As New GoogleDrive
    Dim queryParameters As New Dictionary
    Dim result As String
    
    With queryParameters
        'parametros de consulta
        .Add "q", "name contains 'handbook '"
        'campos a devolver
        .Add "fields", "files(name,id,parents, mimeType,webContentLink)"
    End With
    
    With drive
        .connectionService initOauthFlow
        result = .list(queryParameters)
        
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print result
        End If
    End With
    
End Sub
Sub delete_file()
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    
    fileID = "1hsDZop6aAI25fpwRp7O6gWbX1fgDRhcd"
    
    With drive
        .connectionService initOauthFlow
        .delete (fileID)
        
        If .Operation = GO_NO_CONTENT Then
            Debug.Print "status "; .Operation
        End If
    End With
    
End Sub
Sub empty_trash()
    
    Dim drive As New GoogleDrive

    With drive
        .connectionService initOauthFlow
        .emptyTrash
        
        If .Operation = GO_NO_CONTENT Then
            Debug.Print "status "; .Operation
        End If
    End With
    
End Sub
Sub getMetada_file()
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    Dim queryParameters As New Dictionary
    Dim result As String
    
    fileID = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
'    queryParameters.Add "fields", "id,name, parents,exportLinks"
    queryParameters.Add "fields", "*"
    
    With drive
        .connectionService initOauthFlow
        result = .getMetadata(fileID, queryParameters)

        If .Operation = GO_SUCCESSFUL Then
            Debug.Print result
        End If
    End With
       
End Sub

Sub listLabels_file()
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    Dim queryParameters As New Dictionary
    Dim result As String
    
    fileID = "1d6v2DeKHNSA68RXjCDX-lcqpi7Spz5Nx"
    queryParameters.Add "maxResults", 2
    
    With drive
        .connectionService initOauthFlow
        result = .listLabels(fileID, queryParameters)
        
        If .Operation = GO_SUCCESSFUL Then
            Debug.Print result
        End If
    End With
    
End Sub
Sub generate_labels()
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    Dim queryParameters As New Dictionary
    Dim ids As Collection
    
    With queryParameters
        .Add "count", 2
        .Add "space", "drive"
        .Add "type", "files"
    End With
    
    With drive
        .connectionService initOauthFlow
        Set ids = .generateIds(queryParameters)
        
        If Not ids Is Nothing Then
            For Each ID In ids
                Debug.Print ID
            Next ID
        End If
    End With
    
End Sub
Sub export_file()
    
    'Este se limite a solo 10mb de descarga.
    'Para saber a que tipo de mimeType es exportable, use getMetada con el campo
    'exportLinks(en la documentación propone usar exportFormats).
    'Para saber sobre los formatos a exportar visite:
    'https://developers.google.com/drive/api/guides/ref-export-formats
    
    'Otra forma mas sencilla de exportar es obtener los valores mimeType
    'disponible del archivo, este regresa los enlaces de descarga con los formatos
    'disponibles.
    
    'Sí el archivo existe, este se sobre escribirá sin preguntar.
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    Dim pathFile As String
    Dim mimeType As New Dictionary
    
    fileID = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    mimeType.Add "mimeType", "text/csv"
    pathFile = ThisWorkbook.Path & "\multimedia\test.txt"
    
    With drive
        .connectionService initOauthFlow
        Debug.Print .export(fileID, mimeType, pathFile)
    End With
    
End Sub
Sub download_webContentLink()
    
    Dim drive As New GoogleDrive
    Dim fileID As String
    
    fileID = "1FC3AXegBhMeDWtjE-cPnVWlZAENLkOjXTueMWye7L4w"
    
    With drive
        .connectionService initOauthFlow
        Debug.Print .downloadContentLink(fileID)
    End With
    
End Sub
