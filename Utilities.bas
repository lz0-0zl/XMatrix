Attribute VB_Name = "Utilities"
'@Folder "ABoiler.Utilities"
Option Explicit

Enum eNumberFormat
  NumberFormat_automatic
  NumberFormat_noDec
  NumberFormat_oneDec
  NumberFormat_twoDec
  NumberFormat_threeDec
  NumberFormat_fourDec
  NumberFormat_noFormat
End Enum

Private SharedInstance As EngLogger

Sub resetEngLogger(): Set SharedInstance = Nothing: End Sub

Function getEngLoggerInstance(pComponent As String) As EngLogger
If SharedInstance Is Nothing Then
  Set SharedInstance = New EngLogger
  SharedInstance.init "RUN " & Format(Now, "yyyymmdd-HHNNSS"), component:=pComponent, minLevel:=EngLogLevel_Error, textPath:=ThisWorkbook.path
End If

SharedInstance.setComponent = pComponent

Set getEngLoggerInstance = SharedInstance
End Function

Function rndBtwnDouble(loBound As Double, upBound As Double) As Double
Randomize
rndBtwnDouble = (upBound - loBound) * Rnd + loBound
End Function

Function rndBtwnInteger(loBound As Integer, upBound As Integer) As Integer
Randomize
rndBtwnInteger = Int((upBound - loBound + 1) * Rnd + loBound)
End Function

Sub formatNumbersWithFormula(ws As Worksheet)
'If ws Is Nothing Then Set ws = ThisWorkbook.ActiveSheet

resetEngLogger
ws.Activate
ws.Cells.clear

ws.Cells.Font.name = "Consolas"
ws.Cells.VerticalAlignment = xlCenter
ws.Cells.HorizontalAlignment = xlCenter

ws.Columns(1).IndentLevel = 1
ws.Columns(1).HorizontalAlignment = xlLeft

Dim lastCol As Integer
lastCol = Cells(1, Columns.Count).Column

Dim rng As Range: Set rng = ws.Range(Columns(2), Columns(lastCol))

Dim ds As String, Ts As String
ds = Application.International(xlDecimalSeparator)
Ts = Application.International(xlThousandsSeparator)

rng.FormatConditions.Delete
rng.NumberFormatLocal = "#" & Ts & "##0"

rng.FormatConditions.Add Type:=xlExpression, Formula1:="=abs(B1)<1"
rng.FormatConditions.Add Type:=xlExpression, Formula1:="=abs(B1)<10"
rng.FormatConditions.Add Type:=xlExpression, Formula1:="=abs(B1)<100"
rng.FormatConditions.Add Type:=xlExpression, Formula1:="=abs(B1)<1000"

rng.FormatConditions(1).NumberFormat = "#" & Ts & "##0" & ds & "0000"
rng.FormatConditions(2).NumberFormat = "#" & Ts & "##0" & ds & "000"
rng.FormatConditions(3).NumberFormat = "#" & Ts & "##0" & ds & "00"
rng.FormatConditions(4).NumberFormat = "#" & Ts & "##0" & ds & "0"
End Sub

Function numberToString(value As Double, Optional numFormat As eNumberFormat = NumberFormat_automatic) As String
Dim ds As String, Ts As String
ds = Application.International(xlDecimalSeparator)
Ts = Application.International(xlThousandsSeparator)

Select Case numFormat
Case NumberFormat_noDec: numberToString = FormatNumber(value, 0, vbTrue, vbFalse, vbTrue)
Case NumberFormat_oneDec: numberToString = FormatNumber(value, 1, vbTrue, vbFalse, vbTrue)
Case NumberFormat_twoDec: numberToString = FormatNumber(value, 2, vbTrue, vbFalse, vbTrue)
Case NumberFormat_threeDec: numberToString = FormatNumber(value, 3, vbTrue, vbFalse, vbTrue)
Case NumberFormat_fourDec: numberToString = FormatNumber(value, 4, vbTrue, vbFalse, vbTrue)
Case NumberFormat_automatic
  If Abs(value) < 1 Then
    numberToString = FormatNumber(value, 4, vbTrue, vbFalse, vbTrue)
  ElseIf Abs(value) < 10 Then
    numberToString = FormatNumber(value, 3, vbTrue, vbFalse, vbTrue)
  ElseIf Abs(value) < 100 Then
    numberToString = FormatNumber(value, 2, vbTrue, vbFalse, vbTrue)
  ElseIf Abs(value) < 1000 Then
    numberToString = FormatNumber(value, 1, vbTrue, vbFalse, vbTrue)
  Else
    numberToString = FormatNumber(value, 0, vbTrue, vbFalse, vbTrue)
  End If
Case NumberFormat_noFormat: numberToString = value
End Select
End Function

Sub fastOn()
With Application
  .ScreenUpdating = False
  .EnableEvents = False
  .Calculation = xlCalculationManual
End With
End Sub

Sub fastOff()
With Application
  .ScreenUpdating = True
  .EnableEvents = True
  .Calculation = xlCalculationAutomatic
End With
End Sub






