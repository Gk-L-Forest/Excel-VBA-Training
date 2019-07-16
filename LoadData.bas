Attribute VB_Name = "LoadData"
Option Explicit

'------------------------------------------------------------------------------
' ## 最大列確認および総行数記憶
'------------------------------------------------------------------------------
Public Sub FetchMatrixSize(ByRef origin_file As Workbook, ByRef row_size As Long, ByRef column_size As Long)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim bufferSize As Long
    
    For Each currentSheet In origin_file.Worksheets
        
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
Public Sub StoreToArray(ByRef origin_file As Workbook, ByRef data_array() As Variant)
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim i_row As Long, j_col As Long
    Dim db_row As Long, db_col As Long
    
    db_row = 0
    For Each currentSheet In origin_file.Worksheets
        
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
Public Sub CreateHeader(ByRef column_name() As String)
    
    Dim i As Long, n_col As Long
    
    For i = 1 + ADDITION_COLUMN To UBound(column_name, 2)
        n_col = i - ADDITION_COLUMN
        column_name(1, i) = "列" & n_col
    Next
    
End Sub
