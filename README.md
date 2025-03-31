# Api GoogleSheets for VBA

![Static Badge](https://img.shields.io/badge/VBA-4c9c4c)
![Static Badge](https://img.shields.io/badge/Api%20Google%20Drive-v3-cccccc)

![Microsoft Visual Basic for Applications](./img/google_drive_vba.png)

Carga, descarga, lista, mueve, crea, copia, elimina recursos de tu Google Drive usando la [Google Drive API](https://console.cloud.google.com/marketplace/product/google/drive.googleapis.com?q=search&referrer=search&project=total-messenger-353018) con VBA.

Para este proyecto se ha usado la **versión 3 de Google Drive**.

## Hay mucho por desarrollar

Eh tratado de cubrir las operaciones básicas, sin embargo aún queda mucho por cubrir, si quieres colaborar con el desarrollo de este repositorio haz un fork y sube una PR para discutirlo.

## Tabla de contenido

1. [Instalación](#instalación)
2. [Activar referencias](#activar-referencias)
3. [Configuración de entorno en Google](#configuración-de-entorno-en-google)
4. [Guardar credenciales de acceso](#guardar-credenciales-de-acceso)
5. [Probar FlowOauth y generar el token de acceso](#probar-flowoauth-y-generar-el-token-de-acceso)
6. [Ejemplo de uso](#ejemplo-de-uso)
7. [Recursos adicionales](#Recursos-adicionales)

## Instalación

Puedes hacerlo de cualquiera de las 2 formas:

Sí cuentas con [git](https://git-scm.com/) :

```sh
git clone https://github.com/888Leonidas888/Api-GoogleSheets_for_VBA.git
```

O también puedes descargar este repositorio de forma manual, haz lo siguiente:

1. Presiona en el botón de color verde ![Static Badge](https://img.shields.io/badge/<>%20Code-4c9c4c) en la parte superior derecha.
2. Selecciona la opción **Download ZIP** para comenzar la descarga.
3. Por último descomprime el repositorio que acabas de descargar.

## Activar referencias

Antes de hacer uso, debes asegurarte de tener activadas las siguientes referencias, una vez abierto el archivo te saltará una advertencia pidiendo que actives las macros, acepta para continuar, una vez habilitada las macros presiona `Alt` + `F11` para ir al VBE, en la barra de menú seleciona **Herramientas** -> **Referencias** y procede activar las siguientes referencias:

1. Visual Basic For Applications
2. Microsoft Access 16.0 Object Library
3. OLE Automation
4. Microsoft Office 16.0 Access database engine Object Library
5. Microsoft Excel 16.0 Object Library
6. Microsoft Scripting Runtime
7. Microsoft XML, v6.0
8. Microsoft ActiveX Data Objects 6.1 Library

> [!NOTE]
> Aparte de las referencias mencionadas líneas arriba, también se debe contar con el siguiente módulo [JsonConverter.bas v2.3.1](https://github.com/VBA-tools/VBA-JSON/tree/master), este módulo es imprescindible para poder manipular las respuestas que se reciban por parte de la API, no te preocupes por importarlo, los archivos ya estan equipados con dicho módulo.

## Configuración de entorno en Google

Posiblemente este sea uno de los pasos mas tediosos a seguir pero tomese su tiempo para leerlo detenidamente, pronto agregaré un videotutorial de como hacerlo, pero por ahora siga los pasos en los enlaces o visite [Desarrolla en Google Workspace](https://developers.google.com/workspace/guides/get-started?hl=es_419).

1. [Crea un proyecto de Google Cloud](https://developers.google.com/workspace/guides/create-project?hl=es-419)
2. [Habilita las APIs que deseas usar](https://developers.google.com/workspace/guides/enable-apis?hl=es-419)
3. [Obtén información sobre cómo funcionan la autenticación y autorización](https://developers.google.com/workspace/guides/auth-overview?hl=es-419)
4. [Configura el consentimiento de OAuth](https://developers.google.com/workspace/guides/configure-oauth-consent?hl=es-419)
5. [Crea credenciales de acceso](https://developers.google.com/workspace/guides/create-credentials?hl=es-419)

## Guardar credenciales de acceso

[Las credenciales de acceso](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#api-key) obtenidas debes guardarlas en el directorio **credentials** (no es obligatorio) con extensión **json**, al comienzo solo tendrás 2 archivos; el primero para la [Clave API](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#api-key) y el segundo [ID de cliente de OAuth](https://developers.google.com/workspace/guides/create-credentials?hl=es-419#oauth-client-id)

- **Clave de API:** Guardalo de la siguiente manera, esto es obligatorio, de lo contrario la instancia de `FlowOauth` no podrá encontrar este valor, nombra al archivo como mejor convengas:

```json
{
  "your_api_key": "AIzaSiAsOpGUEW5oS_A6cPkMFLonxGy2uhtgv2j4"
}
```

- **ID de cliente de OAuth:** Solo descargamos y guardamos el archivo, nombra al archivo como mejor convengas, el contenido será algo como esto:

```json
{
  "web": {
    "client_id": "293831635874-8dfdmnbctsmfhsgfhg874.apps.googleusercontent.com",
    "project_id": "elegant-tangent-388222",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_secret": "GOCSPX-...",
    "redirect_uris": ["http://localhost:5500/"],
    "javascript_origins": ["http://localhost:5500"]
  }
}
```

- **Token de acceso:** En este apartado mencioné que solo serián dos archivos, hay un tercer archivo en formato json, este archivo será generado por la instancia de `FlowOauth` cuando lo invoques desde **VBA** al intentar acceder a tu **Google Drive**. Solo debes asegurarte de pasarlo en el argumento `credentialsToken` la ruta de dicho archivo a la instancia de `FlowOauth`. Nombra al archivo como mejor convengas.

> [!NOTE]
> En ningún caso será necesario crear el archivo con el **token de acceso** de forma manual, la instancia de `FlowOauth` se encargará de crearlo si no lo encuentra o actualizarlo según corresponda.

## Probar FlowOauth y generar el token de acceso

Abre tu ventana inmediato en la barra de menú: **Ver** -> **Ventana inmediato** o control `Ctrl` + `G`, ejecuta el sgte código, esto generará el archivo con el **token de acceso**.

- La primera vez que ejecutes este código,sigue estos paso:

  1. Selecciona tu cuenta google.
  2. Luego se mostrará una ventana **Google no ha verificado esta aplicación**; selecciona la opción de **continuar**.
  3. Se te mostrará un ventana indicandote los permisos que estas otorgando para acceder a tus **Google Drive**, selecciona la opción de **continuar**.
  4. La siguiente vista será un **No se puede encontrar esta página (localhost)**, debes ir a la barra de direcciones y copiar el valor de `code`(la parte que indica **code=**`copiar_valor`**&scope**), habrás notado que hay cuadro de diálogo **inputbox** esperando que pegues ese valor, después de aceptar se habrá generado el token en la ruta que le has indicado.

```vb
Function initOauthFlow() As FlowOauth
    ' Use esta función para no tener que redundar en su código.

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
```

Las siguientes líneas son las que verás en tu **ventana inmediato** excepto la última línea hasta que terminé de ejecutar el proceso:

```
Flujo de Oauth 2.0 19/03/2024 15:57:41  >>> flow started
Flujo de Oauth 2.0 19/03/2024 15:57:41  >>> FILE NOT FOUND C:\Users\JHONY\Desktop\Api-GoogleDrive_for_VBA\credentials\token.json
Flujo de Oauth 2.0 19/03/2024 16:04:22  >>> new token generated
```

## Ejemplo de uso

### Listar archivos

```vb
Sub list_file()

    'Vea los siguientes enlaces para aprender hacer consultas

    'https://developers.google.com/drive/api/v3/reference/files/list
    'https://developers.google.com/drive/api/guides/search-files?hl=es-419#specific
    'https://developers.google.com/drive/api/guides/search-files?hl=es-419#examples

    Dim drive As New GoogleDrive
    Dim queryParameters As New Dictionary
    Dim result As String

    With queryParameters
        'parametros de consulta
        .Add "q", "name contains 'handbook '"
        'campos a devolver
        .Add "fields", "files(name,id,parents,mimeType,webContentLink)"
    End With

    With drive
        .connectionService initOauthFlow
        result = .list(queryParameters)

        If .operation = SUCCESSFUL Then
            Debug.Print result
        End If
    End With
End Sub

```

### Cargar video

Con este método puede cargar archivos hasta 5GB.

```vb
Sub upload_resumable()

    Dim drive As New GoogleDrive
    Dim filePath As String
    Dim fileObject As New Dictionary
    Dim parents As New Collection

    On Error GoTo Catch

    filePath = ThisWorkbook.Path & "\multimedia\2025-02-15_21h05_25.mp4"
    parents.Add "root"

    With fileObject
        .Add "parents", parents
        .Add "mimeType", "video/mp4"
        .Add "description", "video upload test 30-3"
        .Add "name", "2025-02-15_21h05_25.mp4"
    End With

    With drive
        .connectionService initOauthFlow
        Debug.Print .uploadResumable(filePath, fileObject)
    End With

    Exit Sub
Catch:
    Debug.Print Err.Number
    Debug.Print Err.Description
End Sub
```

## Recursos adicionales

Los siguientes enlaces estan relacionados a las consultas para listar.

- [Method: files.list](https://developers.google.com/drive/api/v3/reference/files/list)
- [Buscar carpetas o archivos específicos en la sección Mi unidad del usuario actual](https://developers.google.com/drive/api/guides/search-files?hl=es-419#specific)
- [Ejemplos de cadenas de consulta](https://developers.google.com/drive/api/guides/search-files?hl=es-419#examples)
