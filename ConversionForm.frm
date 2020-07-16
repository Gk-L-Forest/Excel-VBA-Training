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
Option Explicit

'------------------------------------------------------------------------------
' ## 設定ファイルのファイル名
'------------------------------------------------------------------------------
Private Const SHEET_CONFIG As String = "\ExclusionarySheet.config"
Private Const ROW_CONFIG As String = "\ExclusionaryRow.config"

'------------------------------------------------------------------------------
' ## フォーム初期化
'
' ここで指定しているプロパティは以下の通り
' *サイズ関係を除く動作上必須のもの
' *コードでしか指定できないもの
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
    End With
    
    ' 設定ファイルの読み込み
    ExclusionarySheetBox.Value = CommitConfig.LoadConfig(SHEET_CONFIG)
    ExclusionaryRowBox.Value = CommitConfig.LoadConfig(ROW_CONFIG)
    
End Sub

'------------------------------------------------------------------------------
' ## ファイルドロップ時の動作(要参照設定)
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
' ## 変換ボタン
'------------------------------------------------------------------------------
Private Sub ConversionButton_Click()
    
    With ConversionForm.FileDorpView.ListItems
        If .Count = 1 Then
            Dim sourceFilePath As String
            Dim exclusionarySheet() As String
            Dim exclusionaryRow() As String
            
            sourceFilePath = .Item(1).SubItems(1)
            
            ' 設定値読み込み
            Call CommitConfig.LoadExclusionarySheet _
                (ExclusionarySheetBox.Value, exclusionarySheet)
            Call CommitConfig.LoadExclusionaryRow _
                (ExclusionaryRowBox.Value, exclusionaryRow)
            
            ' 変換実行
            Call ConvertDatabase.ConvertDatabase _
                (sourceFilePath, exclusionarySheet, exclusionaryRow)
            
            ' 設定フォルダの準備
            Call CommitConfig.PrepareConfigFolder
            
            ' 設定値保存
            Call CommitConfig.SaveConfig _
                (SHEET_CONFIG, ExclusionarySheetBox.Value)
            Call CommitConfig.SaveConfig _
                (ROW_CONFIG, ExclusionaryRowBox.Value)
        Else
            MsgBox "ファイルが指定されていません。", vbExclamation
        End If
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## 逆変換ボタン
'------------------------------------------------------------------------------
Private Sub ReversionButton_Click()
    
    With ConversionForm.FileDorpView.ListItems
        If .Count = 1 Then
            Dim sourceFilePath As String
            sourceFilePath = .Item(1).SubItems(1)
            
            ' 変換実行
            Call ConvertReverse.ConvertReverse(sourceFilePath)
        Else
            MsgBox "ファイルが指定されていません。", vbExclamation
        End If
    End With
    
End Sub

'------------------------------------------------------------------------------
' ## フォームと同時にブックを閉じる
'------------------------------------------------------------------------------
Private Sub UserForm_Terminate()
    
    ' 編集時はコメントアウト
    'ThisWorkbook.Close
    
End Sub
