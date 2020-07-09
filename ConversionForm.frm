VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConversionForm 
   Caption         =   "規則化変換プログラム"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6660
   OleObjectBlob   =   "ConversionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ConversionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 設定ファイルのファイル名
'------------------------------------------------------------------------------
Private Const SHEET_CONFIG = "ExclusionarySheet.config"
Private Const ROW_CONFIG = "ExclusionaryRow.config"

'------------------------------------------------------------------------------
' ## 変換ボタン
'------------------------------------------------------------------------------
Private Sub ConversionButton_Click()
    
    ' 設定値読み込み
    Dim exclusionarySheet() As String
    Dim exclusionaryRow() As String
    
    If ExclusionarySheetBox.Value <> "" Then
        Call LoadConfig.LoadExclusionarySheet _
            (ExclusionarySheetBox.Value, exclusionarySheet)
    End If
    
    If ExclusionaryRowBox.Value <> "" Then
        Call LoadConfig.LoadExclusionaryRow _
            (ExclusionaryRowBox.Value, exclusionaryRow)
        If IsEmpty(exclusionaryRow) Then Exit Sub
    End If
    
    ' 変換実行
    With ConversionForm.FileDorpView.ListItems
        If .Count = 1 Then
            Dim sourceFilePath As String
            sourceFilePath = .Item(1).SubItems(1)
            
            Call ConvertDatabase.ConvertDatabase _
                (sourceFilePath, exclusionarySheet, exclusionaryRow)
        Else
            MsgBox "ファイルが指定されていません。", vbExclamation
        End If
    End With
    
    ' 設定値保存
    Call SaveConfig.SaveConfig(SHEET_CONFIG, ExclusionarySheetBox.Value)
    Call SaveConfig.SaveConfig(ROW_CONFIG, ExclusionaryRowBox.Value)
    
End Sub

'------------------------------------------------------------------------------
' ## ファイルドロップ時の動作
'------------------------------------------------------------------------------
Private Sub FileDorpView_OLEDragDrop _
    (Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, _
     Shift As Integer, x As Single, y As Single)
    
    If Not Data.Files.Count = 1 Then
        MsgBox "変換するファイルは1つにして下さい。", vbExclamation
        Exit Sub
    End If
    
    ' 上書きのため一度Clear
    FileDorpView.ListItems.Clear
    
    With FileDorpView.ListItems.Add
        .Text = Dir(Data.Files(1))
        .SubItems(1) = Data.Files(1)
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## フォーム初期化
'
' ここで指定しているプロパティは以下の通り
' ・サイズ関係を除く動作上必須のもの
' ・コードでしか指定できないもの
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    With FileDorpView
        .OLEDropMode = ccOLEDropManual  ' D&Dの有効化
        .View = lvwReport               ' 表示形式
        .LabelEdit = lvwManual          ' 内容の編集
        .AllowColumnReorder = True      ' 列幅の変更
        .FullRowSelect = True           ' 行全体の選択
        .Gridlines = True               ' グリッド線表示
        
        .ColumnHeaders.Add Text:="ファイル名", Width:=100
        .ColumnHeaders.Add Text:="ファイルパス", Width:=400
    End With
    
    With ExclusionarySheetBox
        .MultiLine = True                   ' 改行の有効化
        .ScrollBars = fmScrollBarsVertical  ' スクロールバー
        
        Dim configSheet As String
        Call LoadConfig.LoadConfig(SHEET_CONFIG, configSheet)
        .Value = configSheet
    End With
    
    With ExclusionaryRowBox
        Dim configRow As String
        Call LoadConfig.LoadConfig(ROW_CONFIG, configRow)
        .Value = configRow
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## フォームと同時にブックを閉じる
'------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    
    'ThisWorkbook.Close
    
End Sub
