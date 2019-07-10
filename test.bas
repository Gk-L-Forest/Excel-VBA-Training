Attribute VB_Name = "ConvertMain"
'--------------------------------------------------------------
' ## コーディングガイドライン
'
' [You.Activate|VBAコーディングガイドライン]に準拠する
' (http://www.thom.jp/vbainfo/codingguideline.html)
'
'--------------------------------------------------------------
Option Explicit

'--------------------------------------------------------------
' ## 規則化変換メインプログラム
'
' 仮作成のプログラム
' 一通りできたらクラス等へ分ける
'
'--------------------------------------------------------------
Private Sub ConvertDatabase()
    
    Dim originFilePath As String
    originFilePath = Application.GetOpenFilename("Microsoft Excelブック, *.xls?")
    If originFilePath = "False" Then
        MsgBox "ファイル選択がキャンセルされました。"
        Exit Sub
    End If
    
    ' ** カレントディレクトリを事前に移動しないと気になる？ **
    
    Dim originFile As Workbook
    Workbooks.Open Filename:=originFilePath, ReadOnly:=True
    Set originFile = Workbooks(Dir(originFilePath))
    
    ' ** 現状は出力先をThisWorkbookとするが後でAddした新規ファイルに変更する **
    
    Dim databaseFile As Workbook
    Set databaseFile = ThisWorkbook
    
    ' ** はじめに最大列確認＆総行数記憶を行う方向で設計変更/その是非は動かしてから確認 **
    
    Dim currentSheet As Worksheet
    Dim currentData As Variant
    Dim rowSize As Long, columnSize As Long, bufferSize As Long
    
    rowSize = 0
    columnSize = 0
    For Each currentSheet In originFile.Worksheets
        
        currentData = currentSheet.UsedRange
        
        rowSize = rowSize + UBound(currentData, 1)
        
        bufferSize = UBound(currentData, 2)
        If bufferSize > columnSize Then columnSize = bufferSize
        
        Erase currentData
        
    Next currentSheet
    
    
    Const ADDITION_COLUMN As Long = 2
    columnSize = columnSize + ADDITION_COLUMN
    
    Dim dataBase() As Variant
    ReDim dataBase(1 To rowSize, 1 To columnSize)
    
    Dim i_row As Long, j_col As Long
    Dim db_row As Long, db_col As Long
    
    db_row = 0
    For Each currentSheet In originFile.Worksheets
        
        currentData = currentSheet.UsedRange
        
        For i_row = 1 To UBound(currentData, 1)
            
            db_row = db_row + 1
            
            dataBase(db_row, 1) = currentSheet.Name
            dataBase(db_row, 2) = i_row
            
            For j_col = 1 To UBound(currentData, 2)
                
                db_col = ADDITION_COLUMN + j_col
                dataBase(db_row, db_col) = currentData(i_row, j_col)
                
            Next j_col
            
        Next i_row
        
        Erase currentData
        
    Next currentSheet
    
    Dim headData() As String
    ReDim headData(1 To 1, 1 To columnSize)
    
    Dim i As Long, n_col As Long
    
    headData(1, 1) = "シート名"
    headData(1, 2) = "行番号"
    For i = 1 + ADDITION_COLUMN To columnSize
        n_col = i - ADDITION_COLUMN
        headData(1, i) = "要素" & n_col
    Next
    
    With ThisWorkbook.ActiveSheet
        '.Range(.Cells(1, 1)).Resize(1, columnSize) = headData
        '.Range(.Cells(1, 1)).Resize(rowSize, columnSize) = dataBase
        .Range(.Cells(1, 1), .Cells(1, columnSize)) = headData
        .Range(.Cells(2, 1), .Cells(1 + rowSize, columnSize)) = dataBase
    End With
    
    originFile.Close
    
    MsgBox "シートのマージが完了しました"
    
End Sub
