Attribute VB_Name = "Utilities"
'@Folder("VBAProject.Utilities")
Option Explicit

' ============================================================================
' Purpose: Generate a random Double value within a specified range.
' Description: Returns a random Double value between loBound (inclusive)
'              and upBound (exclusive).
' Parameters:
'   loBound [Double] - The lower bound of the random range (inclusive).
'   upBound [Double] - The upper bound of the random range (exclusive).
' Return:
'   [Double] - A random number >= loBound and < upBound.
' Dependencies:
'   Uses VBA's Randomize and Rnd functions.
' ============================================================================
Function RndBtwnDouble(loBound As Double, upBound As Double) As Double
Randomize
RndBtwnDouble = (upBound - loBound) * Rnd + loBound
End Function

' ============================================================================
' Purpose: Generate a random Integer value within a specified range.
' Description: Returns a random Integer value between loBound
'              and upBound (both inclusive).
' Parameters:
'   loBound [Integer] - The lower bound of the random range (inclusive).
'   upBound [Integer] - The upper bound of the random range (inclusive).
' Return:
'   [Integer] - A random number >= loBound and <= upBound.
' Dependencies:
'   Uses VBA's Randomize, Rnd, and Int functions.
' ============================================================================
Function RndBtwnInteger(loBound As Integer, upBound As Integer) As Integer
Randomize
RndBtwnInteger = Int((upBound - loBound + 1) * Rnd + loBound)
End Function