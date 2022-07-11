Attribute VB_Name = "Mod_Get_Fields"

Private Function GetFields(ByVal contenido As Dictionary) As String()
        
        Dim arrStr() As String
        Dim index As Integer
        Dim value As Dictionary
        Dim i As Integer
        
        index = contenido("files").count - 1
        ReDim arrStr(0 To index, 2)
        
        For i = 0 To index
            Set value = contenido("files")(i + 1)
            arrStr(i, 0) = value("id")
            arrStr(i, 1) = value("name")
            arrStr(i, 2) = value("parents")(1)
        Next
        
        GetFields = arrStr
        
End Function

Public Function GetFields2(ByVal content As Dictionary) As String()
    
     Dim arrStr() As String
        Dim index As Integer
        Dim value As Dictionary
        Dim i As Integer
        Dim size_file As Double
        Dim size_str As String
        
        On Error GoTo Cath
        
        index = content("files").count - 1
        ReDim arrStr(0 To index, 3)
        
        For i = 0 To index
            Set value = content("files")(i + 1)
            arrStr(i, 0) = value("name")
            
            If (value("owners")(1)("me")) = True Then
                arrStr(i, 1) = "Yo"
            Else
                arrStr(i, 1) = "Otro"
            End If
            
            arrStr(i, 2) = JsonConverter.ParseIso(value("modifiedTime"))
            
            size_file = Val(value("size"))

            Select Case size_file
                Case Is >= WorksheetFunction.Power(1024, 4)
                    size_str = (Format(size_file / (WorksheetFunction.Power(1024, 4)), "#,##0.0")) & "TB"
                Case Is >= WorksheetFunction.Power(1024, 3)
                    size_str = (Format(size_file / (WorksheetFunction.Power(1024, 3)), , "#,##0.0")) & "GB"
                Case Is >= WorksheetFunction.Power(1024, 2)
                    size_str = (Format(size_file / (WorksheetFunction.Power(1024, 2)), "#,##0.0")) & "MB"
                Case Is >= 1024
                    size_str = Format(size_file / 1024, "#,##0.0") & "KB"
                Case Else
                    size_str = Format(size_file, "#,##0.0") & "Bytes"
            End Select
            
            arrStr(i, 3) = size_str
            
        Next
        
        GetFields2 = arrStr
        Exit Function
        
Cath:
Stop
    GetFields2 = arrStr
    Debug.Print Err.description; Chr(9); Err.Number
    On Error GoTo 0
    
End Function

Private Function IsArrayEmpty(ByRef arrStr() As String, Optional container As String) As Boolean

    On Error GoTo Cath

   container = arrStr(0)
    IsArrayEmpty = False

    Exit Function

Cath:

    IsArrayEmpty = True
    On Error GoTo 0

End Function
