VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTTPRequestService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event forRequestHTTP(ByVal status As Integer)

Public header As String

Private Type tState
    AllHeaders As String
    getHeader As String
    responseText As Variant
End Type

Private This As tState

Public Function httpRequest(ByVal methods As String, _
                            ByVal url As String, _
                            Optional body = Empty, _
                            Optional dicHeaders As Dictionary, _
                            Optional response As String) As Variant
    
    Dim http As New MSXML2.ServerXMLHTTP60
    
    With http
        .setTimeouts 0, 100, 60000, 60000
        .Open methods, url, False
        
        If Not dicHeaders Is Nothing Then
            For Each key In dicHeaders.Keys
                .setRequestHeader key, dicHeaders(key)
            Next key
        End If
    
        .send body
        
        RaiseEvent forRequestHTTP(.status)
        
        With This
            .AllHeaders = http.getAllResponseHeaders
            If Me.header <> Empty Then .getHeader = http.getResponseHeader(header)
'            .responseText = http.responseText
            
            Select Case response
                Case Is = "body"
                    .responseText = http.responseBody
                Case Is = "xml"
                    .responseText = http.responseXML
                Case Is = "stream"
                    .responseText = http.responseStream
                Case Else
                    .responseText = http.responseText
            End Select
        
        End With
        
        httpRequest = This.responseText
        
    End With
    
    Set http = Nothing
    
End Function

Public Function AllHeaders() As String
    AllHeaders = This.AllHeaders
End Function

Public Property Get getHeader() As String
    getHeader = This.getHeader
End Property

