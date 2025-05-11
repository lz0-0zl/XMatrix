Attribute VB_Name = "Utilities"
'@Folder("VBAProject.Utilities")
Option Explicit

Function RndBtwnDouble(loBound As Double, upBound As Double) As Double
Randomize
RndBtwnDouble = (upBound - loBound) * Rnd + loBound
End Function

Function RndBtwnInteger(loBound As Integer, upBound As Integer) As Integer
Randomize
RndBtwnInteger = Int((upBound - loBound + 1) * Rnd + loBound)
End Function
