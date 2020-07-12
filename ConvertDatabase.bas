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
Public Sub ConvertDatabase(ByVal source_filepath As String, _
                           ByRef exclusionary_sheet() As String, _
                           ByRef exclusionary_row() As String)
    
    Dim sourceFile As Workbook
    Dim dataFile As Workbook
    
    ' 元ファイルを開く
    Call CommonSub.OpenExcelFile(source_filepath, sourceFile)
    If sourceFile Is Nothing Then Exit Sub
    
    ' 出力ファイルを作成
    Call createNewFile(sourceFile, dataFile)
    If dataFile Is Nothing Then
        sourceFile.Close SaveChanges:=False
        Exit Sub
    End If
    
    CommonProperty.AccelerationMode = True
    
    Dim rowSize As Long: rowSize = 0
    Dim columnSize As Long: columnSize = 0
    Dim columnName() As String
    Dim dataArray() As Variant
    
    ' 総行数および最大列記憶
    Call fetchMatrixSize(sourceFile, rowSize, columnSize)
    
    ' シート名および行番号の要素を確保
    columnSize = columnSize + ADDITION_COLUMN
    
    ' ヘッダの生成
    ReDim columnName(1 To 1, 1 To columnSize)
    columnName(1, 1) = "シート名"
    columnName(1, 2) = "行番号"
    Call createHeader(columnName)
    
    ' シート名/行番号の付加および配列への格納
    ReDim dataArray(1 To rowSize, 1 To columnSize)
    Call storeToArray(sourceFile, dataArray, _
        exclusionary_sheet, exclusionary_row)
    sourceFile.Close SaveChanges:=False
    
    ' ファイルへ出力
    Call outputData(dataFile, rowSize, columnSize, columnName, dataArray)
    
    CommonProperty.AccelerationMode = False
    dataFile.Save
    
    MsgBox "データベース形式への変換が完了しました。", vbInformation
    
End Sub

'------------------------------------------------------------------------------
' ## "元ファイル名_編集用.xlsx"の出力ファイルを同階層に作成
'------------------------------------------------------------------------------
Private Sub createNewFile(ByRef source_file As Workbook, _
                          ByRef new_file As Workbook)
    
    ' ファイル名生成
    Dim extensionPoint As Long
    Dim newFileName As String
    extensionPoint = InStrRev(source_file.Name, ".")
    newFileName = Left(source_file.Name, extensionPoint - 1) & "_編集用.xlsx"
    
    ' ファイルパス生成(元ファイルと同階層)
    Dim newFilePath As String
    newFilePath = source_file.Path & "\" & newFileName
    
    ' 同名ブックの起動有無確認および既存ファイルの存在確認
    If CommonFunction.ConfirmDuplicateBook(newFileName) Then
        If CommonFunction.ConfirmExistingFile(newFilePath) Then
            Set new_file = Workbooks.Add
            new_file.SaveAs FileName:=newFilePath
        End If
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 総行数および最大列記憶
'------------------------------------------------------------------------------
Private Sub fetchMatrixSize(ByRef source_file As Workbook, _
                            ByRef row_size As Long, _
                            ByRef column_size As Long)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim bufferSize As Long
    
    ' HACK: 現状は最大値を記憶しているため値格納と同様に最適化する
    For Each currentSheet In source_file.Worksheets
        ' UsedRangeで最大列および最大行の取得短縮化
        currentData = currentSheet.UsedRange
        If Not IsEmpty(currentData) Then
            ' 総行数の更新および最大列の確認/更新
            row_size = row_size + UBound(currentData, 1)
            bufferSize = UBound(currentData, 2)
            If bufferSize > column_size Then column_size = bufferSize
            
            Erase currentData
        End If
    Next currentSheet
    
End Sub

'------------------------------------------------------------------------------
' ## シート名/行番号の付加および配列への格納
'------------------------------------------------------------------------------
Private Sub storeToArray(ByRef source_file As Workbook, _
                         ByRef data_array() As Variant, _
                         ByRef exclusionary_sheet() As String, _
                         ByRef exclusionary_row() As String)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim current_row As Long, current_col As Long
    Dim data_row As Long, data_col As Long
    
    Dim i As Long
    Dim skipSheet As Long
    Dim skipRow As Long
    
    data_row = 0
    For Each currentSheet In source_file.Worksheets
        ' 除外シート名を照合
        skipSheet = 1
        If Not CommonFunction.IsEmptyArray(exclusionary_sheet) Then
            For i = 0 To UBound(exclusionary_sheet)
                skipSheet = skipSheet * StrComp _
                    (exclusionary_sheet(i), currentSheet.Name)
            Next i
        End If
        If skipSheet <> 0 Then
            ' UsedRangeで配列化短縮化
            currentData = currentSheet.UsedRange
            If Not IsEmpty(currentData) Then
                For current_row = 1 To UBound(currentData, 1)
                    ' 除外行番号を照合
                    skipRow = 1
                    If Not CommonFunction.IsEmptyArray(exclusionary_row) Then
                        For i = 0 To UBound(exclusionary_row)
                            skipRow = skipRow * StrComp _
                                (exclusionary_row(i), current_row)
                        Next i
                    End If
                    If skipRow <> 0 Then
                        data_row = data_row + 1
                        ' シート名/行番号の付加
                        data_array(data_row, 1) = currentSheet.Name
                        data_array(data_row, 2) = current_row
                        ' 列要素の付加数を考慮して配列へ格納
                        For current_col = 1 To UBound(currentData, 2)
                            data_col = ADDITION_COLUMN + current_col
                            data_array(data_row, data_col) = _
                                currentData(current_row, current_col)
                        Next current_col
                    End If
                Next current_row
                
                Erase currentData
            End If
        End If
    Next currentSheet
    
End Sub

'------------------------------------------------------------------------------
' ## ヘッダの生成
'------------------------------------------------------------------------------
Private Sub createHeader(ByRef column_name() As String)
    
    Dim i As Long
    Dim n_col As Long
    
    For i = 1 + ADDITION_COLUMN To UBound(column_name, 2)
        n_col = i - ADDITION_COLUMN
        ' 暫定的に"列**"としている
        column_name(1, i) = "列" & n_col
    Next
    
End Sub

'------------------------------------------------------------------------------
' ## 出力ファイルへの書き込み
'------------------------------------------------------------------------------
Private Sub outputData(ByRef data_file As Workbook, _
                       ByVal row_size As Long, _
                       ByVal column_size As Long, _
                       ByRef column_name() As String, _
                       ByRef data_array() As Variant)
    
    With data_file.Sheets(1)
        .Cells(1, 1).Resize(1, column_size).NumberFormatLocal = "@"
        .Cells(1, 1).Resize(1, column_size) = column_name
        .Cells(2, 1).Resize(row_size, column_size).NumberFormatLocal = "@"
        .Cells(2, 1).Resize(row_size, column_size) = data_array
    End With
    
End Sub
