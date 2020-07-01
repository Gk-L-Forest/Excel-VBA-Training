Attribute VB_Name = "CommonFunction"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 同名ブックの起動有無確認
'------------------------------------------------------------------------------
Public Function ConfirmDuplicateBook _
    (ByVal confirm_filename As String) As Boolean
    
    ConfirmDuplicateBook = True
    
    Dim openingFile As Workbook
    For Each openingFile In Workbooks
        If openingFile.Name = confirm_filename Then
            ConfirmDuplicateBook = False
            MsgBox "同名ブックが開かれているため処理を中断しました。"
            Exit Function
        End If
    Next openingFile
    
End Function

'------------------------------------------------------------------------------
' ## 既存ファイルの存在確認
'------------------------------------------------------------------------------
Public Function ConfirmExistingFile _
    (ByVal confirm_filepath As String) As Boolean
    
    ConfirmExistingFile = True
    
    If Dir(confirm_filepath) <> "" Then
        ConfirmExistingFile = False
        MsgBox "同名ファイルが存在するため処理を中断しました。"
        Exit Function
    End If
    
End Function

