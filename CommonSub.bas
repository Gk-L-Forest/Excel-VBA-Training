Attribute VB_Name = "CommonSub"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 既存のExcelファイルを読み取り専用で開く
'------------------------------------------------------------------------------
Public Sub OpenExcelFile(ByVal open_filepath As String, _
                         ByRef open_file As Workbook)
    
    ' 同名ブックの起動有無確認
    Dim openFileName As String
    openFileName = Dir(open_filepath)
    If Not CommonFunction.ConfirmDuplicateBook(openFileName) Then Exit Sub
    
    ' HACK: アドインではエラーになるためAccelerationModeを使用していない
    Workbooks.Open FileName:=open_filepath, ReadOnly:=True
    Set open_file = Workbooks(openFileName)
    
End Sub
