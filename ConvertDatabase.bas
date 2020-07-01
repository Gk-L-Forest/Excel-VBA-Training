Attribute VB_Name = "ConvertDatabase"
'------------------------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'------------------------------------------------------------------------------
Option Explicit

'------------------------------------------------------------------------------
' ## 列要素の付加数
'------------------------------------------------------------------------------
Private Const ADDITION_COLUMN As Long = 2

'------------------------------------------------------------------------------
' ## シートごとに書かれた帳票のデータベース形式への変換プログラム
'
' 任意の形式で書かれたExcelファイルをシート名と行番号を保持しつつ
' すべてのシートをマージしデータベース形式の表へ変換する
'------------------------------------------------------------------------------
Public Sub ConvertDatabase(ByVal source_filepath As String)
    
    ' 開始時間計測用
    'Dim startTime As Double
    'startTime = Timer
    
    ' TODO: Excelファイルかどうかを判定し選別する
    ' 元ファイルを開く
    Dim sourceFile As Workbook
    Call CommonSub.OpenExcelFile(source_filepath, sourceFile)
    If sourceFile Is Nothing Then Exit Sub
    
    ' 出力ファイルを作成
    Dim dataFile As Workbook
    Call createNewFile(sourceFile, dataFile)
    If dataFile Is Nothing Then
        sourceFile.Close SaveChanges:=False
        Exit Sub
    End If
    
    CommonProperty.AccelerationMode = True
    
    ' 最大列確認および総行数記憶
    Dim rowSize As Long: rowSize = 0
    Dim columnSize As Long: columnSize = 0
    Call fetchMatrixSize(sourceFile, rowSize, columnSize)
    
    ' シート名および行番号の要素を確保
    columnSize = columnSize + ADDITION_COLUMN
    
    ' シート名/行番号と共にすべてのデータを配列へ格納
    Dim dataArray() As Variant
    ReDim dataArray(1 To rowSize, 1 To columnSize)
    Call storeToArray(sourceFile, dataArray)
    sourceFile.Close SaveChanges:=False
    
    ' ヘッダの生成
    Dim columnName() As String
    ReDim columnName(1 To 1, 1 To columnSize)
    columnName(1, 1) = "シート名"
    columnName(1, 2) = "行番号"
    Call createHeader(columnName)
    
    ' ファイルへ出力
    With dataFile.Sheets(1)
        .Cells(1, 1).Resize(1, columnSize).NumberFormatLocal = "@"
        .Cells(1, 1).Resize(1, columnSize) = columnName
        .Cells(2, 1).Resize(rowSize, columnSize).NumberFormatLocal = "@"
        .Cells(2, 1).Resize(rowSize, columnSize) = dataArray
    End With
    
    CommonProperty.AccelerationMode = False
    dataFile.Save
    
    ' 終了時間計測用
    'Dim endTime As Double
    'endTime = Timer
    
    MsgBox "データベース形式への変換が完了しました。"
    'MsgBox endTime - startTime
    
End Sub

'------------------------------------------------------------------------------
' ## "元ファイル名_編集用"の出力ファイルを作成する
'------------------------------------------------------------------------------
Private Sub createNewFile(ByRef source_file As Workbook, _
                          ByRef new_file As Workbook)
    
    ' ファイル名の生成
    Dim extensionPoint As Long
    Dim newFileName As String
    extensionPoint = InStrRev(source_file.Name, ".")
    newFileName = Left(source_file.Name, extensionPoint - 1) & "_編集用.xlsx"
    
    Dim newFilePath As String
    newFilePath = source_file.Path & "\" & newFileName
    
    ' 同名ブックの起動有無確認および既存ファイルの存在確認
    If Not CommonFunction.ConfirmDuplicateBook(newFileName) Then Exit Sub
    If Not CommonFunction.ConfirmExistingFile(newFilePath) Then Exit Sub
    
    Set new_file = Workbooks.Add
    new_file.SaveAs FileName:=newFilePath
    
End Sub

'------------------------------------------------------------------------------
' ## 最大列確認および総行数記憶
'------------------------------------------------------------------------------
Public Sub fetchMatrixSize(ByRef source_file As Workbook, _
                           ByRef row_size As Long, ByRef column_size As Long)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim bufferSize As Long
    
    For Each currentSheet In source_file.Worksheets
        
        currentData = currentSheet.UsedRange
        
        If Not IsEmpty(currentData) Then
            
            row_size = row_size + UBound(currentData, 1)
            
            bufferSize = UBound(currentData, 2)
            If bufferSize > column_size Then column_size = bufferSize
            
            Erase currentData
            
        End If
        
    Next currentSheet
    
End Sub

'------------------------------------------------------------------------------
' ## シート名/行番号を付加し配列へ格納
'------------------------------------------------------------------------------
Public Sub storeToArray(ByRef source_file As Workbook, _
                        ByRef data_array() As Variant)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim i_row As Long, j_col As Long
    Dim db_row As Long, db_col As Long
    
    db_row = 0
    For Each currentSheet In source_file.Worksheets
        
        currentData = currentSheet.UsedRange
        
        If Not IsEmpty(currentData) Then
            
            For i_row = 1 To UBound(currentData, 1)
                
                db_row = db_row + 1
                
                data_array(db_row, 1) = currentSheet.Name
                data_array(db_row, 2) = i_row
                
                For j_col = 1 To UBound(currentData, 2)
                    
                    db_col = ADDITION_COLUMN + j_col
                    data_array(db_row, db_col) = currentData(i_row, j_col)
                    
                Next j_col
                
            Next i_row
            
            Erase currentData
            
        End If
        
    Next currentSheet
    
End Sub

'------------------------------------------------------------------------------
' ## ヘッダの生成
'
' 暫定的に"列**"としているがフォーム入力等で識別子を与えるべき
'------------------------------------------------------------------------------
Public Sub createHeader(ByRef column_name() As String)
    
    Dim i As Long, n_col As Long
    
    For i = 1 + ADDITION_COLUMN To UBound(column_name, 2)
        n_col = i - ADDITION_COLUMN
        column_name(1, i) = "列" & n_col
    Next
    
End Sub
