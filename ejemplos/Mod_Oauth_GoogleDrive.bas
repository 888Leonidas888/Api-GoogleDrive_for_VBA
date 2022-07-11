Attribute VB_Name = "Mod_Oauth_GoogleDrive"
Option Private Module
Option Explicit

Private Const access_token As String = "D:\Mis documentos personales\Proyectos en VBA\Cliente VBA para GoogleDrive\credenciales\token.json"
Private Const api_key As String = "D:\Mis documentos personales\Proyectos en VBA\Cliente VBA para GoogleDrive\credenciales\apikey.json"
Private Const client As String = "D:\Mis documentos personales\Proyectos en VBA\Cliente VBA para GoogleDrive\credenciales\clientWeb.json"
Private Const scope As String = "https://www.googleapis.com/auth/drive"

Sub comenzar_flujo()
    
    
    
End Sub

Rem funciones de limpieza
'@FileDelete
'@EmptyTrash
Sub Eliminar_archivos()
        
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim IsDelete As Boolean
    Dim file_id As String
 
    file_id = "1WgywrlRmUjcFiTTxFwYgd2U2R5q1ZVoH"

    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        IsDelete = .FileDelete(file_id)
    End With
    
    Debug.Print IsDelete
    
    Set oauth = Nothing
    Set google = Nothing
    
End Sub

Sub limpiar_papelera()
        
        Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim response As Variant

    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        response = .EmptyTrash()
    End With
    
    Debug.Print response
    
    Set oauth = Nothing
    Set google = Nothing
    
End Sub

Rem funciones de consulta
'@FileList
'@GetMeta

Sub listar_archivos()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim response As String
    Dim q As String
    Dim fields As String
    Dim pageSize As Integer
    
'    q = "name contains 'Google' and trashed = false"
'    q = "mimeType = 'application/vnd.google-apps.folder'"
'    q = "trashed = true"
    q = "name contains '.xlsm'"
    fields = "name,size,modifiedTime,owners"
    pageSize = 10
            
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        response = .FileList(q, fields, pageSize)
    End With
    
    Debug.Print response
    
    Set oauth = Nothing
    Set google = Nothing
    
End Sub
Sub listar_archivos_mas_array()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim response As Dictionary
    Dim q As String
    Dim fields As String
    Dim pageSize As Integer
    Dim arrStr() As String
    Dim i As Integer
    
'    q = "name contains 'Google' and trashed = false"
'    q = "mimeType = 'application/vnd.google-apps.folder'"
'    q = "trashed = true"
'    q = "name contains '.mp3'"
'    q = Empty
    q = "parents = '1Pc4IfuV9UGVCDFD9elxGcaHWIOzwmvx5'"
    fields = "name,size,owners,modifiedTime"
    pageSize = 30
            
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        Set response = JsonConverter.ParseJson(.FileList(q, fields, pageSize))
    End With
    
'    Debug.Print response
    
    arrStr = GetFields2(response)
    
    If IsArrayEmpty(arrStr) Then
        For i = LBound(arrStr, 1) To UBound(arrStr, 1)
            Debug.Print "NOMBRE : "; arrStr(i, 0)
            Debug.Print "PROPIETARIO :"; arrStr(i, 1)
            Debug.Print "ULTIMA MODIFICACION : "; arrStr(i, 2)
            Debug.Print "TAMAÑO : "; arrStr(i, 3)
            Debug.Print vbCrLf
        Next i
    End If
    
    Set oauth = Nothing
    Set google = Nothing
    Set response = Nothing
    
End Sub

Sub obtener_metadatos()
            
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim response As String
    Dim file_id As String
    Dim fields As String

    file_id = "17go5thKOA_6NeCxX1P2rvuG_IlRl8Xsl"
    fields = "*"
    
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        response = .GetMeta(file_id, fields)
    End With
    
    Debug.Print response
    
    Set oauth = Nothing
    Set google = Nothing

End Sub

Rem funciones de descarga
'@FileDownload2
'@FileDownload3
'@FileDownloadExport
Sub descargar_archivo()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    Dim file_id As String

    file_id = "1BIIIEAlUb5WV7n_xlFyYz3fSFMrTeaEc"

    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        .FileDownload2 file_id
    End With
    
    
    Set oauth = Nothing
    Set google = Nothing
    
    MsgBox "ok  se descargo el archivo"
    
End Sub
Sub descargar_archivo_metodo_dos()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    Dim pathTarget As String
    Dim file_id As String
        
    file_id = "1BIIIEAlUb5WV7n_xlFyYz3fSFMrTeaEc"
    pathTarget = "D:\Mis documentos personales\Proyectos en VBA\Cliente VBA para GoogleDrive"
    
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        .Filedownload3 file_id, pathTarget
    End With
        
    Set oauth = Nothing
    Set google = Nothing
        
End Sub
Sub descargar_archivo_metodo_tres()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    Dim pathTarget As String
    Dim file_id As String
    Dim nameFile As String
    Dim mimeType As String
        
    file_id = "1jank_7hmokao_1UpEF99S0cT_81JbVM8Djj--8x6JTM"
    pathTarget = "D:\Mis documentos personales\Proyectos en VBA\Cliente VBA para GoogleDrive"
    mimeType = "application/pdf"
    nameFile = "mi_pdf.pdf"
    
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        .FileDownloadExport file_id, mimeType, pathTarget, nameFile
    End With
    
    Set oauth = Nothing
    Set google = Nothing
        
End Sub

Rem Funciones de carga
'@CreateFolder
'@FilesUploadResumable

Sub crear_folder()
    
    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    Dim itemJson As Dictionary
    
    Dim response As String
    Dim parents As String
    Dim folderName As String
    Dim id_folder As String
    
    parents = "0AMeDt0PwE9FLUk9PVA"
    folderName = "mi_folder_creado_para_VBA"
    
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        response = .FileCreateFolder(folderName, parents)
    End With
    
    Debug.Print response
    
    Set itemJson = JsonConverter.ParseJson(response)
    id_folder = itemJson("id")
    
    Debug.Print id_folder
    
    Set oauth = Nothing
    Set google = Nothing
    MsgBox "carpeta creada"
End Sub

Sub subir_archivo_reanudable()
            

    Dim oauth As New FlowOauth
    Dim google As New GoogleDriveService
    
    Dim parent As String
    Dim pathFile As String
    
    parent = "1oXIsoy11Zp7--xjQVyqj9mqCQ-b8LuEy"
    pathFile = "D:\Mis documentos personales\Mis videos\api google drive para vba\Demo api google drive (crear folder,subir,contenido y listar) para VBA + UserForm.mp4"
    
    With oauth
        .InicializeFlow access_token, client, api_key, scope
    End With
    
    With google
        .ConnectionService oauth
        .FileUpLoadResumable False, parent, pathFile
    End With
    MsgBox "Ok se cargo el archivo"
    
    Set oauth = Nothing
    Set google = Nothing

        
End Sub


