Attribute VB_Name = "SaveConfig"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 設定ファイルへの書き出し
'------------------------------------------------------------------------------
Public Sub SaveConfig(ByVal config_filename As String, _
                      ByVal config_data As String)
    
    Dim configFilePath As String
    configFilePath = ThisWorkbook.Path & "\config\" & config_filename
    
    Open configFilePath For Output As #1
        Print #1, config_data
    Close #1
    
End Sub

