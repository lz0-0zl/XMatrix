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
'===========================================================================================================================
' XDouble.cls
' A utility class for encapsulating a Double value with additional properties (name, format, unit)
' and event notification on value changes. Provides formatted string representations and cloning.
'===========================================================================================================================

Option Explicit

'===========================================================================================================================
' Purpose: Notifies listeners when the Value property changes.
' Description: Raised whenever the Value property is set to a new value (unless event is disabled).
' Parameters:
'   oldValue - Double: The previous value before change.
'   newValue - Double: The new value assigned.
' Return: None (Event)
' Dependencies: None
'===========================================================================================================================
Public Event ChangedValue(oldValue As Double, newValue As Double)

'===========================================================================================================================
' Purpose: Stores the name of the value for display or identification.
' Description: Used in formatted string output and for identification.
' Parameters: None
' Return: String
' Dependencies: None
'===========================================================================================================================
Public name As String

'===========================================================================================================================
' Purpose: Stores the format string for displaying the value.
' Description: Used by ToString to format the value output.
' Parameters: None
' Return: String
' Dependencies: None
'===========================================================================================================================
Public Frmt As String

'===========================================================================================================================
' Purpose: Stores the unit of the value (e.g., "kg", "m").
' Description: Used in formatted output to indicate the unit.
' Parameters: None
' Return: String
' Dependencies: None
'===========================================================================================================================
Public Unit As String

'===========================================================================================================================
' Purpose: Internal storage for the encapsulated Double value.
' Description: Holds the actual value for the Value property.
' Parameters: None
' Return: Double
' Dependencies: None
'===========================================================================================================================
Private m_value As Double

'===========================================================================================================================
' Purpose: Initializes the class with default property values.
' Description: Sets name, Frmt, and Unit to "null" on object creation.
' Parameters: None
' Return: None
' Dependencies: None
'===========================================================================================================================
Private Sub class_initialize()
' Set default values for properties when the object is created
Frmt = "null"
Unit = "null"
name = "null"
End Sub

'===========================================================================================================================
' Purpose: Gets the encapsulated Double value.
' Description: Returns the current value stored in m_value.
' Parameters:
'   disableEvent (Optional, Boolean): Not used in getter, present for interface symmetry.
' Return: Double - The current value.
' Dependencies: None
'===========================================================================================================================
'@DefaultMember
Property Get Value(Optional disableEvent As Boolean = False) As Double
Attribute Value.VB_UserMemId = 0
' Return the stored value
Value = m_value
End Property

'===========================================================================================================================
' Purpose: Sets the encapsulated Double value.
' Description: Updates m_value and raises ChangedValue event unless disabled.
' Parameters:
'   disableEvent (Optional, Boolean): If True, suppresses the ChangedValue event.
'   pvalue (Double): The new value to assign.
' Return: None
' Dependencies: Application.EnableEvents (Excel VBA)
'===========================================================================================================================
Property Let Value(Optional disableEvent As Boolean = False, pvalue As Double)
Dim oldValue As Double

' Ensure events are enabled (relevant for Excel VBA)
Application.EnableEvents = True
' Store the old value for event notification
oldValue = m_value
' Update the stored value
m_value = pvalue

' Raise the ChangedValue event if not disabled
If Not disableEvent Then RaiseEvent ChangedValue(oldValue, pvalue)
End Property

'===========================================================================================================================
' Purpose: Returns a formatted string including the property name and unit, padded for alignment, followed by the value.
' Description: Combines name, unit, and formatted value for display.
' Parameters: None
' Return: String - The formatted string.
' Dependencies: ToString method
'===========================================================================================================================
Function ToStringWithPropName() As String
Dim tempToString As String: tempToString = "null"

' Start with the formatted value
ToStringWithPropName = ToString
' If name is set, use it
If Not name = "null" Then tempToString = name
' If unit is set, append it in brackets
If Not Unit = "null" Then tempToString = tempToString & " [" & Unit & "]"

' If either name or unit is set, pad for alignment and append value
If Not tempToString = "null" Then
  tempToString = tempToString & ".: "
  ' Pad with dots for alignment up to 60 characters
  Do While Len(tempToString) < 60
    tempToString = Replace(tempToString, ".:", "..:")
  Loop
  ' Concatenate the formatted value
  ToStringWithPropName = tempToString & ToStringWithPropName
End If
End Function

'===========================================================================================================================
' Purpose: Returns the value as a formatted string.
' Description: Uses different formats depending on the magnitude of the value, or uses the custom format if Frmt is set.
' Parameters: None
' Return: String - The formatted value.
' Dependencies: Format (VBA built-in)
'===========================================================================================================================
Function ToString() As String
' Format the value based on its magnitude for better readability
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

' If a custom format is set, use it instead
If Not Frmt = "null" Then ToString = Format(Value, Frmt)
End Function

'===========================================================================================================================
' Purpose: Creates and returns a new XDouble object with the same value, format, name, and unit.
' Description: Performs a shallow copy of the current object.
' Parameters: None
' Return: XDouble - The cloned object.
' Dependencies: XDouble class
'===========================================================================================================================
Function Clone() As XDouble
' Create a new instance of XDouble
Set Clone = New XDouble

' Copy all properties to the new instance
Clone.Value = Value
Clone.Frmt = Frmt
Clone.name = name
Clone.Unit = Unit
End Function