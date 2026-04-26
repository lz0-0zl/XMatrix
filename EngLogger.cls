VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EngLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ABoiler.Utilities"
'IgnoreModule
Option Explicit

'========================================================
' ******************** Types / Enums ********************
'========================================================
Public Enum eEngLogLevel
  EngLogLevel_Debug = 10                                                        ' Detailed diagnostic information
  EngLogLevel_Info = 20                                                         ' Normal workflow or status messages
  EngLogLevel_Warn = 30                                                         ' Something unusual, calculation continues
  EngLogLevel_Error = 40                                                        ' Serious issue, may stop calculation
End Enum

Private Type StepInfo
  stepId As String
  scope As String
  startTick As Double
End Type

'================================================
'******************** Fields ********************
'================================================
Private mRunId As String
Private mComponent As String
Private mMinLevel As eEngLogLevel
Private mLogCsvPath As String
Private mEnableCsv As Boolean
Private mLogTextPath As String
Private mEnableText As Boolean
Private mLogToImmediate As Boolean
Private mMaxTextSize As Long
Private mBackups As Long
Private mInitialized As Boolean

Private mSteps() As StepInfo
Private mStepCount As Long

Private mInfoCount As Long
Private mWarningCount As Long
Private mErrorCount As Long
Private mIssues() As String

'=====================================================
' ******************** Properties ********************
'=====================================================
Property Let setComponent(value As String): mComponent = value: End Property

Property Get warningCount() As Long: warningCount = mWarningCount: End Property

Property Get errorCount() As Long: errorCount = mErrorCount: End Property

Property Get hasSucceeded() As Boolean: hasSucceeded = (mErrorCount = 0): End Property

Property Get hasWarnings() As Boolean: hasWarnings = (mWarningCount > 0): End Property

'=========================================================
' ******************** Public Methods ********************
'=========================================================
Public Sub msgDebug(ByVal msg As String, _
                    Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
writeLog EngLogLevel_Debug, stepId, scope, msg
End Sub

Public Sub msgInfo(ByVal msg As String, _
                   Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
mInfoCount = mInfoCount + 1

If mInfoCount > UBound(mIssues) Then
  ReDim Preserve mIssues(1 To UBound(mIssues) + 1)
End If

mIssues(UBound(mIssues)) = levelName(EngLogLevel_Info) & ": " & msg

writeLog EngLogLevel_Info, stepId, scope, msg
End Sub

Public Sub msgWarn(ByVal msg As String, _
                   Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
mWarningCount = mWarningCount + 1

writeLog EngLogLevel_Warn, stepId, scope, msg
End Sub

Public Sub msgError(ByVal msg As String, _
                    Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
mErrorCount = mErrorCount + 1

writeLog EngLogLevel_Error, stepId, scope, msg
End Sub

' -------- Structured messages -------------------------------
Public Sub msgParam(ByVal name As String, ByVal value As Variant, _
                    Optional ByVal stepId As String = "", Optional ByVal scope As String = "")

writeLog EngLogLevel_Debug, stepId, scope, "PARAM " & name & "=" & CStr(value)
End Sub

Public Sub MsgMetric(ByVal name As String, ByVal value As Double, _
                     Optional ByVal stepId As String = "", Optional ByVal scope As String = "")

writeLog EngLogLevel_Info, stepId, scope, "METRIC " & name & "=" & numberToString(value)
End Sub

Public Sub msgResult(ByVal name As String, ByVal value As Variant, _
                     Optional ByVal stepId As String = "", Optional ByVal scope As String = "")

writeLog EngLogLevel_Info, stepId, scope, "RESULT " & name & "=" & CStr(value)
End Sub

Public Sub msgNote(ByVal msg As String, Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
writeLog EngLogLevel_Info, stepId, scope, "NOTE " & msg
End Sub

Public Sub msgErrorFromErr(Optional ByVal stepId As String = "", Optional ByVal scope As String = "")
If Err.Number <> 0 Then
  writeLog EngLogLevel_Error, stepId, scope, "ERR " & Err.Number & " - " & Err.description
End If
End Sub

Public Function statusSummary() As String
If mErrorCount > 0 Then
  statusSummary = "FAILED with " & mErrorCount & " error(s)"
ElseIf mWarningCount > 0 Then
  statusSummary = "FINISHED with " & mWarningCount & " warning(s)"
Else
  statusSummary = "SUCCESS - no warnings or errors"
End If
End Function

' ========== Step Tracking ==========
Public Function stepStart(ByVal stepId As String, Optional ByVal scope As String = "") As String
pushStep stepId, scope
writeLog EngLogLevel_Info, stepId, scope, "STEP_START"
stepStart = stepId
End Function

Public Sub stepEnd(Optional ByVal extra As String = "")
If mStepCount = 0 Then Exit Sub

Dim S As StepInfo: S = popStep()
Dim dur As Double: dur = elapsed(S.startTick)

writeLog EngLogLevel_Info, S.stepId, S.scope, "STEP_END duration_s=" & Format$(dur, "0.000") & " " & extra
End Sub

Public Sub stepFail(Optional ByVal msg As String = "")
If mStepCount = 0 Then Exit Sub

Dim S As StepInfo: S = popStep()
Dim dur As Double: dur = Timer - S.startTick

writeLog EngLogLevel_Error, S.stepId, S.scope, "STEP_FAIL duration_s=" & Format$(dur, "0.000") & " " & msg
End Sub

'====================================================
' ******************** Lifecycle ********************
'====================================================
Public Sub init(ByVal runId As String, _
                Optional ByVal csvPath As String = "", _
                Optional ByVal component As String = "Main", _
                Optional ByVal minLevel As eEngLogLevel = EngLogLevel_Info, _
                Optional ByVal logToImmediate As Boolean = True, _
                Optional ByVal textPath As String = "", _
                Optional ByVal maxTextBytes As Long = 2& * 1024& * 1024&, _
                Optional ByVal backups As Long = 3)

mRunId = runId

mComponent = component
mMinLevel = minLevel

mLogCsvPath = csvPath & "\" & mRunId & ".csv"
mEnableCsv = (Len(csvPath) > 0)

mLogTextPath = textPath & "\" & mRunId & ".txt"
mEnableText = (Len(textPath) > 0)

mLogToImmediate = logToImmediate

mMaxTextSize = maxTextBytes
mBackups = backups

If mEnableCsv Then
  ensureFolderExists csvPath
  If Not fileExists(mLogCsvPath) Then
    appendText mLogCsvPath, "timestamp,level,run,component,step,scope,message" & vbCrLf
  End If
End If

If mEnableText Then
  ensureFolderExists textPath
  If Not fileExists(mLogTextPath) Then
    appendText mLogTextPath, "timestamp|level|run|component|step|scope|message" & vbCrLf
  End If
End If

ReDim mSteps(1 To 64)
mStepCount = 0

ReDim mIssues(1 To 1)
mInfoCount = 0
mWarningCount = 0
mErrorCount = 0

mInitialized = True
End Sub

Sub reset(Optional ByVal newRunId As String = "")
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

' ---- CSV ----
If mEnableCsv And Len(mLogCsvPath) > 0 Then
  ensureFolderExists mLogCsvPath
  If fileExists(mLogCsvPath) Then fso.DeleteFile mLogCsvPath, True
End If

' ---- TEXT ----
If mEnableText And Len(mLogTextPath) > 0 Then
  ensureFolderExists mLogTextPath

  If fileExists(mLogTextPath) Then fso.DeleteFile mLogTextPath, True

  ' --- Also remove rotated backups if we're doing a full delete ---
  Dim i As Long
  For i = 1 To mBackups
    If fileExists(mLogTextPath & "." & CStr(i)) Then fso.DeleteFile mLogTextPath & "." & CStr(i), True
  Next i
End If

If Len(newRunId) = 0 Then newRunId = "RUN-" & Format(Now, "yyyymmdd-HHNNSS")

init newRunId, mLogCsvPath, mComponent, mMinLevel, mLogToImmediate, mLogTextPath, mMaxTextSize, mBackups
End Sub

'==========================================================
' ******************** Private Methods ********************
'==========================================================
' ========== Core Engine: WriteLog ==========
Private Sub writeLog(ByVal level As eEngLogLevel, ByVal stepId As String, _
                     ByVal scope As String, ByVal msg As String)

If Not mInitialized Then Exit Sub
If level < mMinLevel Then Exit Sub

Dim Ts As String
Ts = Format(Now, "yyyy-mm-dd HH:NN:SS")

' Immediate window
If mLogToImmediate Then
Debug.Print Ts & " | " & levelName(level) & " | " & stepId & " | " & msg
End If

' CSV log
If mEnableCsv Then
  Dim line As String
  line = csv(Ts) & "," & csv(levelName(level)) & "," & csv(mRunId) & "," & csv(mComponent) & _
         "," & csv(stepId) & "," & csv(scope) & "," & csv(msg) & vbCrLf
  appendText mLogCsvPath, line
End If

' Optional text log
If mEnableText Then
  rotateIfNeeded mLogTextPath
  appendText mLogTextPath, Ts & "|" & levelName(level) & "|" & mRunId & "|" & _
                          mComponent & "|" & stepId & "|" & scope & "|" & msg & vbCrLf
End If
End Sub

Private Function levelName(ByVal level As eEngLogLevel) As String
Select Case level
Case EngLogLevel_Debug: levelName = "DEBUG"
Case EngLogLevel_Info:  levelName = "INFO"
Case EngLogLevel_Warn:  levelName = "WARN"
Case EngLogLevel_Error: levelName = "ERROR"
End Select
End Function

Private Function csv(ByVal S As String) As String
S = Replace(S, """", """""")
If InStr(S, ",") > 0 Then csv = """" & S & """" Else csv = S
End Function

Private Sub ensureFolderExists(ByVal path As String)
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
Dim folder As String: folder = fso.GetParentFolderName(path)
If Len(folder) > 0 Then
  If Not fso.FolderExists(folder) Then fso.CreateFolder folder
End If
End Sub

Private Function fileExists(ByVal path As String) As Boolean
fileExists = CreateObject("Scripting.FileSystemObject").fileExists(path)
End Function

Private Sub appendText(ByVal path As String, ByVal content As String)
Dim f As Integer: f = FreeFile
Open path For Append As #f
Print #f, content;
Close #f
End Sub

Private Sub rotateIfNeeded(ByVal path As String)
If mMaxTextSize = 0 Then Exit Sub

Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.fileExists(path) Then Exit Sub

Dim f As Object: Set f = fso.GetFile(path)
If f.Size < mMaxTextSize Then Exit Sub

Dim i As Long
For i = mBackups To 1 Step -1
  If fso.fileExists(path & "." & (i + 1)) Then fso.DeleteFile path & "." & (i + 1)
  If fso.fileExists(path & "." & i) Then fso.MoveFile path & "." & i, path & "." & (i + 1)
Next i

If fso.fileExists(path & ".1") Then fso.DeleteFile path & ".1"
fso.MoveFile path, path & ".1"
End Sub

' ========== Step Stack Handling ==========
Private Sub pushStep(ByVal stepId As String, ByVal scope As String)
mStepCount = mStepCount + 1

If mStepCount > UBound(mSteps) Then
  ReDim Preserve mSteps(1 To UBound(mSteps) * 2)
End If

mSteps(mStepCount).stepId = stepId
mSteps(mStepCount).scope = scope
mSteps(mStepCount).startTick = Timer
End Sub

Private Function popStep() As StepInfo
If mStepCount = 0 Then Exit Function

popStep = mSteps(mStepCount)
mStepCount = mStepCount - 1
End Function

Private Function elapsed(ByVal startTick As Double) As Double
Dim t As Double: t = Timer

If t < startTick Then t = t + 86400#

elapsed = t - startTick
End Function

