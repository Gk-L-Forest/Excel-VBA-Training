Attribute VB_Name = "CommonSub"
Option Explicit

'------------------------------------------------------------------------------
' ## 既存のExcelファイルを読み取り専用で開く
'------------------------------------------------------------------------------
Public Sub OpenExcelFile(ByVal open_filepath As String, _
                         ByRef open_file As Workbook)
    
    'Dim openFilePath As String
    'openFilePath = _
    '    Application.GetOpenFilename("Microsoft Excelブック, *.xls?")
    
    If open_filepath = "False" Then
        MsgBox "ファイルが選択されていません。"
        Exit Sub
    Else
        ' 同名ブックが開いているかの確認
        Dim openFileName As String
        openFileName = Dir(open_filepath)
        If CommonFunction.ConfirmDuplicateFile(openFileName) Then Exit Sub
    End If
    
    CommonProperty.AccelerationMode = True
    Workbooks.Open FileName:=open_filepath, ReadOnly:=True
    Set open_file = Workbooks(openFileName)
    CommonProperty.AccelerationMode = False
    
End Sub
