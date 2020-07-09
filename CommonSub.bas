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
    
    Dim openFileName As String
    openFileName = Dir(open_filepath)
    
    If Not openFileName Like "*.xls?" Then
        MsgBox "Excelファイルを指定して下さい。", vbCritical
        Exit Sub
    End If
    
    ' 同名ブックの起動有無確認
    If CommonFunction.ConfirmDuplicateBook(openFileName) Then
        ' HACK: アドインでエラーになるためAccelerationModeを使用していない
        Workbooks.Open FileName:=open_filepath, ReadOnly:=True
        Set open_file = Workbooks(openFileName)
    End If
    
End Sub
