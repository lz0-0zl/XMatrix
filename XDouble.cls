VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject.Utilities")
Option Explicit

Public Event ChangedValue(oldValue As Double, newValue As Double)

Public name As String
Public Frmt As String
Public Unit As String

Private m_value As Double

Private Sub class_initialize()
Frmt = "null"
Unit = "null"
name = "null"
End Sub

'@DefaultMember
Property Get Value(Optional disableEvent As Boolean = False) As Double
Attribute Value.VB_UserMemId = 0
Value = m_value
End Property

Property Let Value(Optional disableEvent As Boolean = False, pvalue As Double)
Dim oldValue As Double

Application.EnableEvents = True
oldValue = m_value
m_value = pvalue

If Not disableEvent Then RaiseEvent ChangedValue(oldValue, pvalue)
End Property

Function ToStringWithPropName() As String
Dim tempToString As String: tempToString = "null"

ToStringWithPropName = ToString
If Not name = "null" Then tempToString = name
If Not Unit = "null" Then tempToString = tempToString & " [" & Unit & "]"

If Not tempToString = "null" Then
  tempToString = tempToString & ".: "
  Do While Len(tempToString) < 60
    tempToString = Replace(tempToString, ".:", "..:")
  Loop
  ToStringWithPropName = tempToString & ToStringWithPropName
End If
End Function

Function ToString() As String
If Abs(Value) < 1 Then
  ToString = Format(Value, "#,##0.0000")
ElseIf Abs(Value) < 10 Then
  ToString = Format(Value, "#,##0.000")
ElseIf Abs(Value) < 100 Then
  ToString = Format(Value, "#,##0.00")
ElseIf Abs(Value) < 1000 Then
  ToString = Format(Value, "#,##0.0")
Else
  ToString = Format(Value, "#,##0")
End If

If Not Frmt = "null" Then ToString = Format(Value, Frmt)
End Function

Function Clone() As XDouble
Set Clone = New XDouble

Clone.Value = Value
Clone.Frmt = Frmt
Clone.name = name
Clone.Unit = Unit
End Function
