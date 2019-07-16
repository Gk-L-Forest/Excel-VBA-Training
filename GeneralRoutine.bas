Attribute VB_Name = "GeneralRoutine"
Option Explicit

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
' ## 既存のExcelファイルを開く(汎用)
'------------------------------------------------------------------------------
Public Sub OpenExcelFile(ByRef open_file As Workbook)
    
    Dim openFilePath As String
    openFilePath = Application.GetOpenFilename("Microsoft Excelブック, *.xls?")
    
    If openFilePath = "False" Then
        MsgBox "ファイル選択がキャンセルされました。"
        Exit Sub
    Else
        
        Dim openFileName As String
        openFileName = Dir(openFilePath)
        
        Dim openingFile As Workbook
        For Each openingFile In Workbooks
            
            If openingFile.Name = openFileName Then
                MsgBox "同名ブックが既に開いています。"
                Exit Sub
            End If
            
        Next openingFile
        
    End If
    
    GeneralRoutine.AccelerationMode = True
    Workbooks.Open FileName:=openFileName, ReadOnly:=True
    Set open_file = Workbooks(openFileName)
    GeneralRoutine.AccelerationMode = False
    
End Sub

'------------------------------------------------------------------------------
' ## "元ファイル名_編集用"の出力ファイルを作成する
'------------------------------------------------------------------------------
Public Sub CreateNewFile(ByRef origin_file As Workbook, ByRef new_file As Workbook)
    
    Dim extensionPoint As Long
    Dim newFileName As String
    extensionPoint = InStrRev(origin_file.Name, ".")
    newFileName = Left(origin_file.Name, extensionPoint - 1) & "_編集用.xlsx"
    
    Dim newFilePath As String
    newFilePath = origin_file.Path & "\" & newFileName
    
    If Dir(newFilePath) <> "" Then
        MsgBox "同名ファイルが存在するため処理を中断しました。"
        origin_file.Close SaveChanges:=False
        Exit Sub
    Else
        GeneralRoutine.AccelerationMode = True
        Set new_file = Workbooks.Add
        new_file.SaveAs FileName:=newFilePath
        GeneralRoutine.AccelerationMode = False
    End If
    
End Sub
