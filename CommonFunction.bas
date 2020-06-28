Attribute VB_Name = "CommonFunction"
Option Explicit

'------------------------------------------------------------------------------
' ## 同名ブックが開いているかの確認
'------------------------------------------------------------------------------
Public Function ConfirmDuplicateFile(ByVal open_filename As String) As Boolean
    
    ConfirmDuplicateFile = False
    
    Dim openingFile As Workbook
    For Each openingFile In Workbooks
        
        If openingFile.Name = open_filename Then
            ConfirmDuplicateFile = True
            MsgBox "同名ブックが開かれているため処理を中断しました。"
            Exit Function
        End If
        
    Next openingFile
    
End Function
