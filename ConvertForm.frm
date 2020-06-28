VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ConvertForm 
   Caption         =   "UserForm1"
   ClientHeight    =   2115
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
   OleObjectBlob   =   "ConvertForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ConvertForm"
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

Private Sub ConvertButton_Click()
    With ConvertForm.OpenFileView.ListItems
        If .Count = 1 Then
            Call ConvertMain.ConvertDatabase(.Item(1).SubItems(1))
        Else
            MsgBox "ファイルが指定されていません。"
        End If
    End With
End Sub

Private Sub OpenFileView_OLEDragDrop(Data As MSComctlLib.DataObject, _
                                         Effect As Long, _
                                         Button As Integer, _
                                         Shift As Integer, _
                                         x As Single, _
                                         y As Single)
    
    If Not Data.Files.Count = 1 Then
        MsgBox "変換するファイルは1つにして下さい。"
        Exit Sub
    End If
    
    OpenFileView.ListItems.Clear
    With OpenFileView.ListItems.Add
        .Text = Dir(Data.Files(1))
        .SubItems(1) = Data.Files(1)
    End With
    
End Sub

Private Sub UserForm_Initialize()
    
    With ConvertForm
        
        .Caption = "規則化変換プログラム"
        
    End With
    
    With OpenFileView
        
        .OLEDropMode = ccOLEDropManual  ' D&Dの有効化
        .View = lvwReport               ' 表示形式
        .LabelEdit = lvwManual          ' 内容の編集
        .AllowColumnReorder = True      ' 列幅の変更
        .FullRowSelect = True           ' 行全体の選択
        .Gridlines = True               ' グリッド線表示
        
        .ColumnHeaders.Add Text:="ファイル名", Width:=150
        .ColumnHeaders.Add Text:="ファイルパス", Width:=450
        
    End With
    
End Sub

Private Sub UserForm_Terminate()
    'ThisWorkbook.Close
End Sub
