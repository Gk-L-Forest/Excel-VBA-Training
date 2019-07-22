Attribute VB_Name = "ConvertReverse"
Option Explicit

'------------------------------------------------------------------------------
' ## ヘッダ行数
'------------------------------------------------------------------------------
Private Const HEADER_LINENO As Long = 1

'------------------------------------------------------------------------------
' ## データベース形式から元フォーマットへの逆変換プログラム
'
' ConverDatabaseにて作成したデータを元フォーマットへ返す
'------------------------------------------------------------------------------
Public Sub ConvertSource()
    
    ' データベース形式ファイルを開く
    Dim dataFile As Workbook
    Call GeneralRoutine.OpenExcelFile(dataFile)
    If dataFile Is Nothing Then Exit Sub
    
    ' データファイルのファイル名が条件外の場合閉じて終了
    If Not dataFile.Name Like "*_編集用.xlsx" Then
        dataFile.Close SaveChanges:=False
        MsgBox "[*_編集用.xlsx]のファイルを選択してください。"
        Exit Sub
    End If
    
    ' 元ファイルを開く
    Dim sourceFile As Workbook
    Call openSourceFile(dataFile, sourceFile)
    If sourceFile Is Nothing Then Exit Sub
    
    GeneralRoutine.AccelerationMode = True
    
    ' データベース形式のデータを配列に格納
    Dim dataArray As Variant
    dataArray = dataFile.Sheets(1).UsedRange
    dataFile.Close SaveChanges:=False
    
    ' 元フォーマットのファイルへ出力
    Dim db_row As Long, db_col As Long
    Dim n_row As Long, n_col As Long
    Dim currentSheet As Worksheet
    For db_row = 1 + HEADER_LINENO To UBound(dataArray, 1)
        
        Set currentSheet = sourceFile.Worksheets(dataArray(db_row, 1))
        n_row = dataArray(db_row, 2)
        
        ' シートへの出力処理(もう少し改善したい)
        For db_col = 1 + ADDITION_COLUMN To UBound(dataArray, 2)
            
            n_col = db_col - ADDITION_COLUMN
            currentSheet.Cells(n_row, n_col).Value = dataArray(db_row, db_col)
            
        Next db_col
        
    Next db_row
    
    GeneralRoutine.AccelerationMode = False
    sourceFile.Save
    
    MsgBox "元ファイルへの逆変換が完了しました。"
    
End Sub

'------------------------------------------------------------------------------
' ## "元ファイル名_編集用"から元ファイルを開く
'------------------------------------------------------------------------------
Private Sub openSourceFile(ByRef data_file As Workbook, _
                           ByRef source_file As Workbook)
    
    Dim lowlinePoint As Long
    Dim sourceFileName As String
    lowlinePoint = InStrRev(data_file.Name, "_")
    sourceFileName = Left(data_file.Name, lowlinePoint - 1) & ".xlsm"
    
    Dim sourceFilePath As String
    sourceFilePath = data_file.Path & "\" & sourceFileName
    
    ' エラーになりうる場合はデータベース形式ファイルを閉じて終了
    If Dir(sourceFilePath) = "" Then
        MsgBox "元ファイルが見つかりません。"
        data_file.Close SaveChanges:=False
        Exit Sub
    ElseIf ConfirmDuplicateFile(sourceFileName) Then
        data_file.Close SaveChanges:=False
        Exit Sub
    End If
    
    GeneralRoutine.AccelerationMode = True
    Workbooks.Open FileName:=sourceFilePath
    Set source_file = Workbooks(sourceFileName)
    GeneralRoutine.AccelerationMode = False
    
End Sub
