Attribute VB_Name = "GeneralRoutine"
Option Explicit

'------------------------------------------------------------------------------
' ## 列要素の付加数
'------------------------------------------------------------------------------
Public Const ADDITION_COLUMN As Long = 2

'------------------------------------------------------------------------------
' ## 画面更新/イベント検知/自動計算の制御
'------------------------------------------------------------------------------
Public Property Let AccelerationMode(ByVal flg As Boolean)
    With Application
        .ScreenUpdating = Not flg
        .EnableEvents = Not flg
        .Calculation = IIf(flg, xlCalculationManual, xlCalculationAutomatic)
    End With
End Property

'------------------------------------------------------------------------------
' ## 既存のExcelファイルを読み取り専用で開く(汎用)
'------------------------------------------------------------------------------
Public Sub OpenExcelFile(ByRef open_file As Workbook)
    
    Dim openFilePath As String
    openFilePath = Application.GetOpenFilename("Microsoft Excelブック, *.xls?")
    
    If openFilePath = "False" Then
        MsgBox "ファイル選択がキャンセルされました。"
        Exit Sub
    Else
        ' 同名ブックが開いているかの確認
        Dim openFileName As String
        openFileName = Dir(openFilePath)
        If ConfirmDuplicateFile(openFileName) Then Exit Sub
    End If
    
    GeneralRoutine.AccelerationMode = True
    Workbooks.Open FileName:=openFilePath, ReadOnly:=True
    Set open_file = Workbooks(openFileName)
    GeneralRoutine.AccelerationMode = False
    
End Sub

'------------------------------------------------------------------------------
' ## 同名ブックが開いているかの確認(汎用)
'------------------------------------------------------------------------------
Public Function ConfirmDuplicateFile(ByVal open_filename As String) As Boolean
    
    ConfirmDuplicateFile = False
    
    Dim openingFile As Workbook
    For Each openingFile In Workbooks
        
        If openingFile.Name = open_filename Then
            ConfirmDuplicateFile = True
            MsgBox "同名ブックが開かれているため処理を中断しました。"
            Exit Function
        End If
        
    Next openingFile
    
End Function
