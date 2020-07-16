Attribute VB_Name = "CommitConfig"
Option Explicit

'------------------------------------------------------------------------------
' ## 設定フォルダのフォルダ名
'------------------------------------------------------------------------------
Private Const CONFIG_FOLDER As String = "\config"

'------------------------------------------------------------------------------
' ## 設定ファイルの読み込み
'------------------------------------------------------------------------------
Public Function LoadConfig(ByVal config_filename As String) As String
    
    LoadConfig = ""
    
    Dim configFilePath As String
    configFilePath = ThisWorkbook.Path & CONFIG_FOLDER & config_filename
    
    If Dir(configFilePath) = "" Then Exit Function
    
    Dim bufferData As String
    Open configFilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, bufferData
            LoadConfig = LoadConfig & bufferData & vbCrLf
        Loop
        LoadConfig = Left(LoadConfig, Len(LoadConfig) - Len(vbCrLf))
    Close #1
    
End Function

'------------------------------------------------------------------------------
' ## 除外するシート名の読み込み
'------------------------------------------------------------------------------
Public Sub LoadExclusionarySheet(ByVal config_sheetname As String, _
                                 ByRef exclusionary_sheet() As String)
    
    Dim bufferSheetName() As String
    bufferSheetName = Split(config_sheetname, vbCrLf)
    
    ' 半角/全角スペースのみを除いて除外シート名を記憶
    Dim i As Long
    Dim blankChecker As String
    Dim n_sheet As Long: n_sheet = 0
    For i = 0 To UBound(bufferSheetName)
        blankChecker = Replace(Replace(bufferSheetName(i), " ", ""), "　", "")
        If blankChecker <> "" Then
            ReDim Preserve exclusionary_sheet(0 To n_sheet)
            exclusionary_sheet(n_sheet) = bufferSheetName(i)
            n_sheet = n_sheet + 1
        End If
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## 除外する行番号の読み込み
'------------------------------------------------------------------------------
Public Sub LoadExclusionaryRow(ByVal config_rownumber As String, _
                               ByRef exclusionary_row() As String)
    
    If config_rownumber = "" Then Exit Sub
    
    ' 設定値が数字,ハイフン,カンマのみか確認
    If Not validateSingleCharacter(config_rownumber) Then Exit Sub
    
    Dim bufferRow() As String
    bufferRow = Split(config_rownumber, ",")
    
    ' カンマ区切りの要素ごとの確認
    Dim i As Long
    Dim hyphenPosition As Long
    Dim fullRowNumber As String: fullRowNumber = ""
    For i = 0 To UBound(bufferRow)
        If bufferRow(i) = "" Then
            MsgBox "カンマ区切りの指定が不正です。", vbCritical
            Exit Sub
        End If
        ' 範囲入力の確認
        hyphenPosition = InStr(bufferRow(i), "-")
        If hyphenPosition > 0 Then
            ' 範囲入力(ハイフン有り)の場合は行番号へ変換
            Call convertRowNumber(hyphenPosition, bufferRow(i))
            ' 変換出来ていない場合は終了
            If InStr(bufferRow(i), "-") > 0 Then Exit Sub
        End If
        ' カンマ区切りの行番号として記憶
        fullRowNumber = fullRowNumber & bufferRow(i)
        If i <> UBound(bufferRow) Then fullRowNumber = fullRowNumber & ","
    Next i
    
    exclusionary_row = Split(fullRowNumber, ",")
    
End Sub

'------------------------------------------------------------------------------
' ## 除外する行番号の入力が数字,ハイフン,カンマのみか確認
'------------------------------------------------------------------------------
Private Function validateSingleCharacter _
    (ByVal config_rownumber As String) As Boolean
    
    validateSingleCharacter = True
    
    Dim i As Long
    Dim singleCharacter As String
    
    For i = 1 To Len(config_rownumber)
        singleCharacter = Mid(config_rownumber, i, 1)
        If Not IsNumeric(singleCharacter) _
        And singleCharacter <> "-" _
        And singleCharacter <> "," Then
            validateSingleCharacter = False
            MsgBox "設定値に数字,ハイフン,カンマ以外が含まれています。", _
                vbCritical
            Exit Function
        End If
    Next i
    
End Function

'------------------------------------------------------------------------------
' ## 範囲入力を各行番号へ変換
'------------------------------------------------------------------------------
Private Sub convertRowNumber(ByVal hyphen_position As Long, _
                             ByRef row_range As String)
    
    Dim smallNumber As String
    Dim largeNumber As String
    
    smallNumber = Left(row_range, hyphen_position - 1)
    largeNumber = Mid(row_range, hyphen_position + 1)
    
    ' 範囲入力として問題が無いか確認
    If Not IsNumeric(smallNumber) Or Not IsNumeric(largeNumber) Then
        MsgBox "行番号の範囲入力の数値が不正です。", vbCritical
        Exit Sub
    ElseIf CLng(smallNumber) >= CLng(largeNumber) Then
        MsgBox "行番号の範囲入力の大小が不正です。", vbCritical
        Exit Sub
    End If
    
    row_range = ""
    
    ' カンマ区切りの行番号へ変換
    Dim i As Long
    For i = CLng(smallNumber) To CLng(largeNumber)
        row_range = row_range & i
        If i <> CLng(largeNumber) Then row_range = row_range & ","
    Next i
    
End Sub

'------------------------------------------------------------------------------
' ## 設定フォルダの存在確認および作成
'------------------------------------------------------------------------------
Public Sub PrepareConfigFolder()
    
    Dim configFolderPath As String
    configFolderPath = ThisWorkbook.Path & CONFIG_FOLDER
    
    If Dir(configFolderPath, vbDirectory) = "" Then
        MkDir configFolderPath
    End If
    
End Sub

'------------------------------------------------------------------------------
' ## 設定ファイルへの書き出し
'------------------------------------------------------------------------------
Public Sub SaveConfig(ByVal config_filename As String, _
                      ByVal config_data As String)
    
    Dim configFilePath As String
    configFilePath = ThisWorkbook.Path & CONFIG_FOLDER & config_filename
    
    Open configFilePath For Output As #1
        Print #1, config_data
    Close #1
    
End Sub
