Attribute VB_Name = "ConvertReverse"
Option Explicit

'------------------------------------------------------------------------------
' ## ヘッダ行数
'------------------------------------------------------------------------------
Private Const HEADER_LINENO As Long = 1

'------------------------------------------------------------------------------
' ## データベース形式から元ファイルへの逆変換プログラム
'
' ConverDatabaseにて作成したデータを元ファイルへ返す
'------------------------------------------------------------------------------
Public Sub ConvertReverse(ByVal source_filepath As String)
    
    ' 逆変換についての確認
    Dim confirmationMessage As VbMsgBoxResult
    confirmationMessage = _
        MsgBox("元ファイルへの逆変換を行います。問題ありませんか？", _
            vbYesNo + vbQuestion)
    If confirmationMessage = vbNo Then Exit Sub
    
    ' 元ファイルからデータベース形式ファイルのパス生成
    Dim extensionPoint As Long
    Dim dataFilePath As String
    extensionPoint = InStrRev(source_filepath, ".")
    dataFilePath = Left(source_filepath, extensionPoint - 1) & "_編集用.xlsx"
    
    If Dir(dataFilePath) = "" Then
        MsgBox "編集用ファイルが存在しません。", vbCritical
        Exit Sub
    End If
    
    ' データベース形式ファイルを開く
    Dim dataFile As Workbook
    Call CommonSub.OpenBookReadOnly(dataFilePath, dataFile)
    If dataFile Is Nothing Then Exit Sub
    
    ' 元ファイルを開く
    Dim sourceFile As Workbook
    Call openSourceFile(source_filepath, sourceFile)
    If sourceFile Is Nothing Then
        dataFile.Close SaveChanges:=False
        Exit Sub
    End If
    
    CommonProperty.AccelerationMode = True
    
    ' データベース形式のデータを配列に格納
    Dim dataArray As Variant
    dataArray = dataFile.Sheets(1).UsedRange
    dataFile.Close SaveChanges:=False
    
    ' 元ファイルへデータを戻す
    Call returnData(dataArray, sourceFile)
    
    CommonProperty.AccelerationMode = False
    
    ' 上書き保存の確認
    Dim saveMessage As VbMsgBoxResult
    saveMessage = _
        MsgBox("元ファイルへの逆変換が完了しました。保存しますか？", _
            vbYesNo + vbInformation)
    If saveMessage = vbYes Then sourceFile.Save
    
End Sub

'------------------------------------------------------------------------------
' ## 元ファイルを開く
'------------------------------------------------------------------------------
Private Sub openSourceFile(ByVal source_filepath As String, _
                           ByRef source_file As Workbook)
    
    Dim sourceFileName As String
    sourceFileName = Dir(source_filepath)
    
    ' 同名ブックの起動有無確認
    If CommonFunction.IsDuplicateBook(sourceFileName) Then
        MsgBox "同名ブックが開かれているため処理を中断しました。", vbCritical
        Exit Sub
    End If
    
    Workbooks.Open Filename:=source_filepath
    Set source_file = Workbooks(sourceFileName)
    
End Sub

'------------------------------------------------------------------------------
' ## 元ファイルへデータを戻す
'------------------------------------------------------------------------------
Private Sub returnData(ByRef data_array As Variant, _
                       ByRef source_file As Workbook)
    
    Dim data_row As Long, data_col As Long
    Dim sheet_row As Long, sheet_col As Long
    Dim sheetName As String
    Dim currentSheet As Worksheet
    
    For data_row = 1 + HEADER_LINENO To UBound(data_array, 1)
        
        sheetName = data_array(data_row, 1)
        sheet_row = data_array(data_row, 2)
        
        ' シート名または行番号が空白の場合はスキップ
        If sheetName <> "" Or sheet_row < 1 Then GoTo Continue_data_row
        
        ' 行データを元ファイルへ戻す
        Set currentSheet = source_file.Worksheets(sheetName)
        For data_col = 1 + ADDITION_COLUMN To UBound(data_array, 2)
            sheet_col = data_col - ADDITION_COLUMN
            currentSheet.Cells(sheet_row, sheet_col).Value = _
                data_array(data_row, data_col)
        Next data_col
        
Continue_data_row:
        
    Next data_row
    
    Set currentSheet = Nothing
    
End Sub
