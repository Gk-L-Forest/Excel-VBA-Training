Attribute VB_Name = "CommonProperty"
Option Explicit

'------------------------------------------------------------------------------
' ## 画面更新/イベント検知/自動計算の制御
'------------------------------------------------------------------------------
Public Property Let AccelerationMode(ByVal flg As Boolean)
    
    ' ブックが一つも開いていないとCalculationはエラーになるため注意
    With Application
        .ScreenUpdating = Not flg
        .EnableEvents = Not flg
        .Calculation = IIf(flg, xlCalculationManual, xlCalculationAutomatic)
    End With
    
End Property
