VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'
' [規則化変換プログラム] v1.0   2020/07/16 G.O.
'
' シートごとに書かれた帳票を不要シートおよび不要行を除外し
' 値のみのデータベース形式へ変換/逆変換を行うプログラム
'
'
' ##コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## アドイン起動時
'------------------------------------------------------------------------------
Private Sub Workbook_Open()
    Load ConversionForm
    ConversionForm.Show vbModeless
End Sub

'------------------------------------------------------------------------------
' ## アドイン終了時
'------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' ブックが先に閉じてしまった場合エラーになるため
    ' ブックを閉じる前に必ずUnloadしUserForm_Terminateを経由させる
    Unload ConversionForm
End Sub
