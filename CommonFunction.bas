Attribute VB_Name = "CommonFunction"
Option Explicit

'------------------------------------------------------------------------------
' ## 同名ブックの起動有無確認
'------------------------------------------------------------------------------
Public Function IsDuplicateBook _
    (ByVal confirmation_filename As String) As Boolean
    
    IsDuplicateBook = False
    
    Dim openingFile As Workbook
    For Each openingFile In Workbooks
        If openingFile.Name = confirmation_filename Then
            IsDuplicateBook = True
            Exit Function
        End If
    Next openingFile
    
End Function

'------------------------------------------------------------------------------
' ## 配列版IsEmpty
'------------------------------------------------------------------------------
Public Function IsEmptyArray(ByRef confirmation_array As Variant) As Boolean
    
    On Error GoTo Error_Handler
    
    ' エラーまたは最大要素数が0未満の場合は空
    IsEmptyArray = IIf(UBound(confirmation_array) < 0, True, False)
    
    Exit Function
    
Error_Handler:
    IsEmptyArray = True
    
End Function
