Attribute VB_Name = "modOptimize"
Option Explicit

Dim prevCalc, prevEvents, prevScreen, prevPageBreaks
Dim prevDisplayAlerts, prevEnableEvents, prevCalculation

Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub

Sub OptimizeOn()
    'Optimize VBA exectution
    prevCalc = Application.Calculation: Application.Calculation = xlCalculationManual
    prevEvents = Application.EnableEvents: Application.EnableEvents = False
    prevScreen = Application.ScreenUpdating: Application.ScreenUpdating = False
    prevPageBreaks = ActiveSheet.DisplayPageBreaks: ActiveSheet.DisplayPageBreaks = False
End Sub

Sub OptimizeOff()
    'Turn off VBA optimization
    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    ActiveSheet.DisplayPageBreaks = prevPageBreaks
End Sub
