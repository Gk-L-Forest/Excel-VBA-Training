Attribute VB_Name = "ConvertMain"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

Public Const ADDITION_COLUMN As Long = 2

'------------------------------------------------------------------------------
' ## シートごとに書かれた帳票のデータベース形式への変換メインプログラム
'
' 任意の形式で書かれたExcelファイルをシート名と行番号を保持しつつ
' すべてのシートをマージしデータベース形式の表へ変換する
'------------------------------------------------------------------------------
Private Sub ConvertDatabase()
    
    ' 元ファイルを開く
    Dim originFile As Workbook
    Call GeneralRoutine.OpenExcelFile(originFile)
    If originFile Is Nothing Then Exit Sub
    
    ' 出力ファイルを作成
    Dim databaseFile As Workbook
    Call CreateNewFile(originFile, databaseFile)
    If databaseFile Is Nothing Then Exit Sub
    
    GeneralRoutine.AccelerationMode = True
    
    ' 最大列確認および総行数記憶
    Dim rowSize As Long: rowSize = 0
    Dim columnSize As Long: columnSize = 0
    Call LoadData.FetchMatrixSize(originFile, rowSize, columnSize)
    
    ' シート名および行番号の要素を確保
    columnSize = columnSize + ADDITION_COLUMN
    
    ' シート名/行番号と共にすべてのデータを配列へ格納
    Dim dataArray() As Variant
    ReDim dataArray(1 To rowSize, 1 To columnSize)
    Call LoadData.StoreToArray(originFile, dataArray)
    originFile.Close SaveChanges:=False
    
    ' ヘッダの生成
    Dim columnName() As String
    ReDim columnName(1 To 1, 1 To columnSize)
    columnName(1, 1) = "シート名"
    columnName(1, 2) = "行番号"
    Call LoadData.CreateHeader(columnName)
    
    ' ファイルへ出力
    With databaseFile.Sheets(1)
        .Cells(1, 1).Resize(1, columnSize).NumberFormatLocal = "@"
        .Cells(1, 1).Resize(1, columnSize) = columnName
        .Cells(2, 1).Resize(rowSize, columnSize).NumberFormatLocal = "@"
        .Cells(2, 1).Resize(rowSize, columnSize) = dataArray
    End With
    
    GeneralRoutine.AccelerationMode = False
    databaseFile.Save
    
    MsgBox "シートのマージが完了しました"
    
End Sub
