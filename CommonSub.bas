Attribute VB_Name = "CommonSub"
Option Explicit

'------------------------------------------------------------------------------
' ## 既存のExcelファイルを読み取り専用で開く
'------------------------------------------------------------------------------
Public Sub OpenBookReadOnly(ByVal open_filepath As String, _
                            ByRef open_file As Workbook)
    
    Dim openFileName As String
    openFileName = Dir(open_filepath)
    
    If Not openFileName Like "*.xls?" Then
        MsgBox "Excelファイルを指定して下さい。", vbCritical
        Exit Sub
    End If
    
    ' 同名ブックの起動有無確認
    If CommonFunction.IsDuplicateBook(openFileName) Then
        MsgBox "同名ブックが開かれているため処理を中断しました。", vbCritical
        Exit Sub
    End If
    
    Workbooks.Open Filename:=open_filepath, ReadOnly:=True
    Set open_file = Workbooks(openFileName)
    
End Sub
