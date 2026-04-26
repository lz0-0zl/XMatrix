Attribute VB_Name = "AlgeLin"
'@Folder "ABoiler.Utilities"
Option Explicit

Private Function mLogger() As EngLogger: Set mLogger = getEngLoggerInstance("Algelin"): End Function

'=================================================
'****************** Properties *******************
'=================================================
Function arrNumberOfRows(matrixA() As Double) As Integer: arrNumberOfRows = UBound(matrixA, 1) - LBound(matrixA, 1) + 1: End Function
Function arrNumberOfCols(matrixA() As Double) As Integer: arrNumberOfCols = UBound(matrixA, 2) - LBound(matrixA, 2) + 1: End Function

Function arrIsSymmetric(matrixA() As Double) As Boolean
Dim i As Integer, j As Integer

arrIsSymmetric = False

If Not arrIsSquare(matrixA) Then Exit Function

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To i - 1
    If matrixA(i, j) <> matrixA(j, i) Then Exit Function
  Next
Next

arrIsSymmetric = True
End Function

Function arrIsSquare(matrixA() As Double) As Boolean

If arrNumberOfRows(matrixA) <> arrNumberOfCols(matrixA) Then
  arrIsSquare = False
Else
  arrIsSquare = True
End If
End Function

Function arrIsSameSize(matrixA() As Double, matrixB() As Double) As Boolean
If arrNumberOfRows(matrixA) = arrNumberOfRows(matrixB) And arrNumberOfCols(matrixA) = arrNumberOfCols(matrixB) Then
  arrIsSameSize = True
Else
  arrIsSameSize = False
End If
End Function

'=================================================
'***************** Aritimethics ******************
'=================================================
Function arrAdd(matrixA() As Double, matrixB() As Double) As Double()
Dim i As Integer, j As Integer

On Error GoTo ErrHandler

If Not arrIsSameSize(matrixA, matrixB) Then Err.Raise 5, "AlgeLin.bas", "Incompatible matrix dimensions"

Dim result() As Double
ReDim result(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixB, 2) To UBound(matrixB, 2))

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    result(i, j) = result(i, j) + matrixB(i, j)
  Next
Next

CleanExit:
  arrAdd = result
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "arrAdd"
End Function

Function arrSub(matrixA() As Double, matrixB() As Double) As Double()
Dim i As Integer, j As Integer

On Error GoTo ErrHandler

If Not arrIsSameSize(matrixA, matrixB) Then Err.Raise 5, "AlgeLin.bas", "Incompatible matrix dimensions"

Dim result() As Double
ReDim result(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixB, 2) To UBound(matrixB, 2))

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    result(i, j) = result(i, j) - matrixB(i, j)
  Next
Next

CleanExit:
  arrSub = result
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "arrSub"
End Function

Function arrMult(matrixA() As Double, matrixB() As Double) As Double()
Dim i As Integer, j As Integer, k As Integer

On Error GoTo ErrHandler

If Not arrNumberOfCols(matrixA) = arrNumberOfRows(matrixB) Then Err.Raise 5, "AlgeLin.bas", "Incompatible matrix dimensions"

Dim result() As Double
ReDim result(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixB, 2) To UBound(matrixB, 2))

For k = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      result(i, k) = result(i, k) + matrixA(i, j) * matrixB(j, k)
    Next
  Next
Next

CleanExit:
  arrMult = result
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "arrMult"
End Function

Private Sub addScaledIdentity(matrixA() As Double, scaleI As Double)
Dim i As Integer
For i = LBound(matrixA, 1) To UBound(matrixA, 1): matrixA(i, i) = matrixA(i, i) + scaleI: Next
End Sub

Function arrTranspose(matrixA() As Double) As Double()
Dim i As Integer, j As Integer

Dim result() As Double
ReDim result(LBound(matrixA, 2) To UBound(matrixA, 2), LBound(matrixA, 1) To UBound(matrixA, 1))

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    result(j, i) = matrixA(i, j)
  Next
Next

arrTranspose = result
End Function

Function arrTrace(matrixA() As Double) As Double
Dim i As Integer

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"

For i = LBound(matrixA, 1) To UBound(matrixA, 1): arrTrace = arrTrace + matrixA(i, i): Next

CleanExit:
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "arrTrace"
End Function

'=================================================
'***************** Initializing ******************
'=================================================
Sub arrToRandDouble(ByRef matrixA() As Double, min As Double, max As Double, Optional sym As Boolean = True)
Dim i As Integer, j As Integer

ReDim matrixA(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

If Not arrIsSquare(matrixA) Then sym = False

If sym Then
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If i <= j Then
        matrixA(i, j) = rndBtwnDouble(min, max)
      Else
        matrixA(i, j) = matrixA(j, i)
      End If
    Next
  Next
Else
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      matrixA(i, j) = rndBtwnDouble(min, max)
    Next
  Next
End If
End Sub

Sub arrToRandInteger(ByRef matrixA() As Double, min As Integer, max As Integer, Optional sym As Boolean = True)
Dim i As Integer, j As Integer

ReDim matrixA(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

If Not arrIsSquare(matrixA) Then sym = False

If sym Then
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If i <= j Then
        matrixA(i, j) = rndBtwnInteger(min, max)
      Else
        matrixA(i, j) = matrixA(j, i)
      End If
    Next
  Next
Else
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      matrixA(i, j) = rndBtwnInteger(min, max)
    Next
  Next
End If
End Sub

Sub arrToRandTriDouble(ByRef matrixA() As Double, min As Double, max As Double, tri As Integer, Optional sym As Boolean = True)
Dim i As Integer, j As Integer

ReDim matrixA(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

If Not arrIsSquare(matrixA) Then sym = False

If sym Then
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If i <= j Then
        If j > i - tri And j < i + tri Then
          matrixA(i, j) = rndBtwnDouble(min, max)
        Else
          matrixA(i, j) = 0
        End If
      Else
        matrixA(i, j) = matrixA(j, i)
      End If
    Next
  Next
Else
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If j > i - tri And j < i + tri Then
        matrixA(i, j) = rndBtwnDouble(min, max)
      Else
        matrixA(i, j) = 0
      End If
    Next
  Next
End If
End Sub

Sub arrToRandTriInteger(ByRef matrixA() As Double, min As Integer, max As Integer, tri As Integer, Optional sym As Boolean = True)
Dim i As Integer, j As Integer

ReDim matrixA(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

If Not arrIsSquare(matrixA) Then sym = False

If sym Then
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If i <= j Then
        If j > i - tri And j < i + tri Then
          matrixA(i, j) = rndBtwnInteger(min, max)
        Else
          matrixA(i, j) = 0
        End If
      Else
        matrixA(i, j) = matrixA(j, i)
      End If
    Next
  Next
Else
  For i = LBound(matrixA, 1) To UBound(matrixA, 1)
    For j = LBound(matrixA, 2) To UBound(matrixA, 2)
      If j > i - tri And j < i + tri Then
        matrixA(i, j) = rndBtwnInteger(min, max)
      Else
        matrixA(i, j) = 0
      End If
    Next
  Next
End If
End Sub

Sub arrToIdentity(ByRef matrixA() As Double, Optional factor As Double = 1)
Dim i As Integer

ReDim matrixA(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

For i = LBound(matrixA, 1) To UBound(matrixA, 1): matrixA(i, i) = 1 * factor: Next
End Sub

'=================================================
'************ Linear System Solvers **************
'=================================================
Sub luSolverInPlace(ByRef matrixA() As Double, ByRef matrixB() As Double)

Dim i As Long, j As Long, k As Long, pvt As Long, col As Long
Dim f As Double, maxVal As Double, tmp As Double

On Error GoTo ErrHandler

' ---------- Elimination with partial pivoting ----------
For k = LBound(matrixA, 1) To UBound(matrixA, 1) - 1
  ' --- Pivot search ---
  pvt = k
  maxVal = Abs(matrixA(k, k))
  For i = k + 1 To UBound(matrixA, 1)
    If Abs(matrixA(i, k)) > maxVal Then
      maxVal = Abs(matrixA(i, k))
      pvt = i
    End If
  Next i

  If maxVal = 0 Then Err.Raise 1, "Algelin.bas", "Singular matrix"

  ' --- Row swap in matrixA ---
  If pvt <> k Then
    For j = k To UBound(matrixA, 2)
      tmp = matrixA(k, j): matrixA(k, j) = matrixA(pvt, j): matrixA(pvt, j) = tmp
    Next j

    ' --- Row swap in matrixB ---
    For col = LBound(matrixB, 2) To UBound(matrixB, 2)
      tmp = matrixB(k, col): matrixB(k, col) = matrixB(pvt, col): matrixB(pvt, col) = tmp
    Next col
  End If

  ' --- Elimination ----
  For i = k + 1 To UBound(matrixA, 1)
    f = matrixA(i, k) / matrixA(k, k)
    matrixA(i, k) = 0

    For j = k + 1 To UBound(matrixA, 2)
      matrixA(i, j) = matrixA(i, j) - f * matrixA(k, j)
    Next j

    For col = LBound(matrixB, 2) To UBound(matrixB, 2)
      matrixB(i, col) = matrixB(i, col) - f * matrixB(k, col)
    Next col
  Next i
Next k

' ---- Back substitution (result stored in matrixB) ---
For col = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = UBound(matrixA, 1) To LBound(matrixA, 1) Step -1

    For j = i + 1 To UBound(matrixA, 2)
      matrixB(i, col) = matrixB(i, col) - matrixA(i, j) * matrixB(j, col)
    Next j

    matrixB(i, col) = matrixB(i, col) / matrixA(i, i)
  Next i
Next col

CleanExit:
  Exit Sub
ErrHandler:
  mLogger.msgErrorFromErr , "luSolverInPlace"
End Sub

' ========== LU DECOMPOSITION ==========
Function luSolver(matrixA() As Double, matrixB() As Double) As Double()
Dim matrixLower() As Double, matrixUpper() As Double
Dim result() As Double: result = matrixB

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"
If Not arrNumberOfRows(matrixA) = arrNumberOfRows(matrixB) Then Err.Raise 5, "AlgeLin.bas", "Incompatible matrix dimensions"

ReDim matrixLower(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))
ReDim matrixUpper(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

luDecompose matrixA, matrixLower, matrixUpper
fowardSub matrixLower, result
backwardSub matrixUpper, result

luSolver = result

CleanExit:
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "luSolver"
End Function

Function luSolverHomo(matrixA() As Double) As Double()
Dim matrixB() As Double

ReDim matrixB(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To LBound(matrixA, 2))

luSolverHomo = luSolver(matrixA, matrixB)
End Function

Sub luDecompose(matrixA() As Double, ByRef matrixLower() As Double, ByRef matrixUpper() As Double)
Dim i As Integer, j As Integer, k As Integer

ReDim matrixLower(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))
ReDim matrixUpper(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)

    If i <= j Then
      matrixUpper(i, j) = matrixA(i, j)
      If i = j Then matrixLower(i, j) = 1

      For k = LBound(matrixA, 1) To i - 1
        matrixUpper(i, j) = matrixUpper(i, j) - matrixLower(i, k) * matrixUpper(k, j)
      Next
    Else
      matrixLower(i, j) = matrixA(i, j)
      For k = LBound(matrixA, 1) To j - 1
        matrixLower(i, j) = matrixLower(i, j) - matrixLower(i, k) * matrixUpper(k, j)
      Next
      matrixLower(i, j) = matrixLower(i, j) / matrixUpper(j, j)
    End If
  Next
Next
End Sub

Private Sub fowardSub(lowerTri() As Double, matrixB() As Double)
Dim i As Integer, j As Integer, k As Integer

For k = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = LBound(lowerTri, 1) To UBound(lowerTri, 1)
    matrixB(i, k) = matrixB(i, k)
    For j = LBound(lowerTri, 2) To i - 1
      matrixB(i, k) = matrixB(i, k) - lowerTri(i, j) * matrixB(j, k)
    Next
    matrixB(i, k) = matrixB(i, k) / lowerTri(i, i)
  Next
Next
End Sub

Private Sub backwardSub(upperTri() As Double, matrixB() As Double)
Dim i As Integer, j As Integer, k As Integer

For k = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = UBound(upperTri, 1) To LBound(upperTri, 1) Step -1
    matrixB(i, k) = matrixB(i, k)
    For j = i + 1 To UBound(upperTri, 2)
      matrixB(i, k) = matrixB(i, k) - upperTri(i, j) * matrixB(j, k)
    Next
    matrixB(i, k) = matrixB(i, k) / upperTri(i, i)
  Next
Next
End Sub

' ========== LU DECOMPOSITION WITH PARTIAL PIVOTING ==========
Private Function pivot(ByRef matrixA() As Double) As Integer()
Dim pvt() As Integer: ReDim pvt(LBound(matrixA, 1) To UBound(matrixA, 1))

Dim i As Integer, j As Integer
Dim maxRow As Integer, tmp As Integer
Dim maxVal As Double

' --- Initialize permutation vector ---
For i = LBound(matrixA, 1) To UBound(matrixA, 1): pvt(i) = i: Next i

' --- Partial pivoting (row pivoting) ---
For j = LBound(matrixA, 2) To IIf(UBound(matrixA, 2) < UBound(matrixA, 1) - 1, UBound(matrixA, 2), UBound(matrixA, 1) - 1)
  maxRow = j
  maxVal = Abs(matrixA(pvt(j), j))

  For i = j + 1 To UBound(matrixA, 1)
    If Abs(matrixA(pvt(i), j)) > maxVal Then
      maxVal = Abs(matrixA(pvt(i), j))
      maxRow = i
    End If
  Next i

  If maxRow <> j Then
    tmp = pvt(j)
    pvt(j) = pvt(maxRow)
    pvt(maxRow) = tmp
  End If
Next j

pivot = pvt
End Function

Function luSolverWithPivoting(matrixA() As Double, matrixB() As Double) As Double()
Dim matrixLower() As Double, matrixUpper() As Double
Dim result() As Double: result = matrixB
Dim pvt() As Integer: pvt = pivot(matrixA)

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"
If Not arrNumberOfRows(matrixA) = arrNumberOfRows(matrixB) Then Err.Raise 5, "AlgeLin.bas", "Incompatible matrix dimensions"

luDecomposeWithPivoting matrixA, matrixLower, matrixUpper, pvt
fowardSubWithPivoting matrixLower, result, pvt
backwardSubWithPivoting matrixUpper, result

luSolverWithPivoting = result

CleanExit:
  Exit Function
ErrHandler:
  mLogger.msgErrorFromErr , "luSolverWithPivoting"
End Function

Sub luDecomposeWithPivoting(matrixA() As Double, ByRef matrixLower() As Double, ByRef matrixUpper() As Double, pvt() As Integer)
Dim i As Integer, j As Integer, k As Integer

ReDim matrixLower(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))
ReDim matrixUpper(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)

    If i <= j Then
      matrixUpper(i, j) = matrixA(pvt(i), j)
      If i = j Then matrixLower(i, j) = 1

      For k = LBound(matrixA, 1) To i - 1
        matrixUpper(i, j) = matrixUpper(i, j) - matrixLower(i, k) * matrixUpper(k, j)
      Next
    Else
      matrixLower(i, j) = matrixA(pvt(i), j)
      For k = LBound(matrixA, 1) To j - 1
        matrixLower(i, j) = matrixLower(i, j) - matrixLower(i, k) * matrixUpper(k, j)
      Next
      matrixLower(i, j) = matrixLower(i, j) / matrixUpper(j, j)
    End If
  Next
Next
End Sub

Private Sub fowardSubWithPivoting(matrixLower() As Double, matrixB() As Double, pvt() As Integer)
Dim i As Integer, j As Integer, k As Integer
Dim matrixY() As Double

ReDim matrixY(LBound(matrixB, 1) To UBound(matrixB, 1), LBound(matrixB, 2) To UBound(matrixB, 2))

For k = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = LBound(matrixLower, 1) To UBound(matrixLower, 1)
    matrixY(i, k) = matrixB(pvt(i), k)
    For j = LBound(matrixLower, 2) To i - 1
      matrixY(i, k) = matrixY(i, k) - matrixLower(i, j) * matrixY(j, k)
    Next
    'matrixY(i, k) = matrixY(i, k) / matrixLower(i, i)
  Next
Next

matrixB = matrixY
End Sub

Private Sub backwardSubWithPivoting(matrixUpper() As Double, matrixB() As Double)
Dim i As Integer, j As Integer, k As Integer
Dim maxU As Double: maxU = 0

On Error GoTo ErrHandler

For i = LBound(matrixUpper, 1) To UBound(matrixUpper, 1)
  If Abs(matrixUpper(i, i)) > maxU Then maxU = Abs(matrixUpper(i, i))
Next

For k = LBound(matrixB, 2) To UBound(matrixB, 2)
  For i = UBound(matrixUpper, 1) To LBound(matrixUpper, 1) Step -1
    'matrixB(i, k) = matrixB(i, k)
    For j = i + 1 To UBound(matrixUpper, 2)
      matrixB(i, k) = matrixB(i, k) - matrixUpper(i, j) * matrixB(j, k)
    Next
    If Abs(matrixUpper(i, i)) < 10 ^ -12 * maxU Then Err.Raise 5, "AlgeLin.bas", "Singular or ill-conditioned matrix"
    matrixB(i, k) = matrixB(i, k) / matrixUpper(i, i)
  Next
Next

CleanExit:
  Exit Sub
ErrHandler:
  mLogger.msgErrorFromErr , "backwardSubWithPivoting"
End Sub

'=================================================
'********* Eigenvectors & Eigenvalues ************
'=================================================
Function arrPolynom(matrixA() As Double) As Double()
Dim i As Integer
Dim matrixC() As Double: ReDim matrixC(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To LBound(matrixA, 2))
Dim matrixB() As Double: matrixB = matrixA

matrixC(LBound(matrixA, 1), LBound(matrixA, 2)) = -arrTrace(matrixA)

Dim row As Integer: row = LBound(matrixA, 1)

For i = 2 To arrNumberOfRows(matrixA)
  addScaledIdentity matrixB, matrixC(row, LBound(matrixA, 2))
  matrixB = arrMult(matrixA, matrixB)

  row = row + 1: matrixC(row, LBound(matrixA, 2)) = -arrTrace(matrixB) / i
Next

arrPolynom = matrixC
End Function

' ========== EIGENVALUES USING LU DECOMPOSITION ==========
Function arrEigenValues(matrixA() As Double) As Double()
Dim i As Integer, j As Integer
Dim matrixLower() As Double, matrixUpper() As Double, result() As Double
Dim matrixB() As Double: matrixB = matrixA

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"

ReDim result(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To LBound(matrixA, 2))

Dim mu As Double
Const tol As Double = 10 ^ -6
Const MAXIter As Integer = 10 ^ 4

For i = 1 To MAXIter
  ' ---- Shift ----
  mu = matrixB(UBound(matrixB, 1), UBound(matrixB, 2))
  addScaledIdentity matrixB, -mu

  ' ---- QR via LU ----
  luDecompose matrixB, matrixLower, matrixUpper
  matrixB = arrMult(matrixUpper, matrixLower)

  ' ---- UNDO SHIFT ----
  addScaledIdentity matrixB, mu

  If maxAbsInLower(matrixB) < tol Then
    For j = LBound(matrixA, 1) To UBound(matrixA, 1): result(j, LBound(matrixA, 2)) = matrixB(j, j): Next
    arrEigenValues = result
    Exit Function
  End If
Next

Err.Raise 5, "AlgeLin.bas", "LU decomposition did not converge"

ErrHandler:
  mLogger.msgErrorFromErr , "arrEigenValues"
End Function

Private Function maxAbsInLower(matrixA() As Double) As Double
Dim i As Integer, j As Integer

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To i - 1
    If maxAbsInLower < Abs(matrixA(i, j)) Then maxAbsInLower = Abs(matrixA(i, j))
  Next
Next
End Function

' ========== EIGENVECTORS & EIGENVALUES USING QR DECOMPOSITION ==========
Sub arrEigenVectorsAndValues(matrixA() As Double, ByRef eigVectors() As Double, ByRef eigValues() As Double)
Dim i As Integer, j As Integer
Dim matrixQ() As Double, matrixR() As Double, Qtot() As Double, matrixB() As Double

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"
If Not arrIsSymmetric(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-symmetric matrix"

ReDim eigVectors(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))
ReDim eigValues(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To LBound(matrixA, 2))
ReDim Qtot(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

Dim mu As Double
Const tol As Double = 10 ^ -6
Const MAX_ITER As Integer = 10 ^ 4

arrToIdentity Qtot
matrixB = matrixA

For i = 1 To MAX_ITER
  mu = wilkinsonShift(matrixB)
  addScaledIdentity matrixB, -mu

  qrDecompose matrixB, matrixQ, matrixR
  matrixB = arrMult(matrixR, matrixQ)

  addScaledIdentity matrixB, mu

  Qtot = arrMult(Qtot, matrixQ)

  If maxAbsInLower2(matrixB) < tol Then
    For j = LBound(matrixB, 1) To UBound(matrixB, 1)
      eigValues(j - LBound(matrixB, 1) + LBound(eigValues, 1), LBound(eigValues, 2)) = matrixB(j, j)
    Next

    eigVectors = Qtot
    Exit Sub
  End If
Next

Err.Raise 5, "AlgeLin.bas", "QR decomposition did not converge"

ErrHandler:
  mLogger.msgErrorFromErr , "arrEigenVectorsAndValues"
End Sub

Sub qrDecompose(matrixA() As Double, ByRef matrixQ() As Double, ByRef matrixR() As Double)
Dim i As Long, j As Long, k As Long
Dim matrixV() As Double, beta As Double, S As Double, dot As Double

ReDim matrixQ(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))
arrToIdentity matrixQ

matrixR = matrixA

For k = LBound(matrixA, 1) To Application.min(UBound(matrixA, 1) - 1, UBound(matrixA, 2))

  ' --- Build vector x ---
  ReDim matrixV(k To UBound(matrixA, 1))
  For i = k To UBound(matrixA, 1)
    matrixV(i) = matrixR(i, k)
  Next i

  ' --- Compute norm ---
  S = arrNorm2(matrixV)
  If S = 0 Then GoTo NextK

  If matrixV(k) >= 0 Then S = -S
  matrixV(k) = matrixV(k) - S

  ' --- Normalize matrixV ---
  beta = arrNorm2(matrixV)
  If beta = 0 Then GoTo NextK

  For i = k To UBound(matrixA, 1)
    matrixV(i) = matrixV(i) / beta
  Next i

  ' --- Apply reflector to matrixR ---
  For j = k To UBound(matrixA, 2)
    dot = 0
    For i = k To UBound(matrixA, 1)
      dot = dot + matrixV(i) * matrixR(i, j)
    Next i
    For i = k To UBound(matrixA, 1)
      matrixR(i, j) = matrixR(i, j) - 2 * matrixV(i) * dot
    Next i
  Next j

  ' --- Apply reflector to matrixQ ---
  For j = LBound(matrixA, 1) To UBound(matrixA, 1)
    dot = 0
    For i = k To UBound(matrixA, 1)
      dot = dot + matrixV(i) * matrixQ(i, j)
    Next i
    For i = k To UBound(matrixA, 1)
      matrixQ(i, j) = matrixQ(i, j) - 2 * matrixV(i) * dot
    Next i
  Next j

NextK:
Next k

' matrixQ was built as product of reflectors on left and need to be transposed
matrixQ = arrTranspose(matrixQ)
End Sub

Private Function wilkinsonShift(matrixA() As Double) As Double
Dim N As Integer
Dim ann As Double, ann1 As Double, matrixB As Double
Dim DELTA As Double, signDelta As Double

ann = matrixA(UBound(matrixA, 1), UBound(matrixA, 1))
ann1 = matrixA(UBound(matrixA, 1) - 1, UBound(matrixA, 1) - 1)
matrixB = matrixA(UBound(matrixA, 1), UBound(matrixA, 1) - 1)

DELTA = (ann1 - ann) / 2

If DELTA >= 0 Then
  signDelta = 1
Else
  signDelta = -1
End If

wilkinsonShift = ann - signDelta * matrixB * matrixB / _
                 (Abs(DELTA) + Sqr(DELTA * DELTA + matrixB * matrixB))

End Function

Private Function maxAbsInLower2(matrixA() As Double) As Double
Dim i As Integer, j As Integer

For i = LBound(matrixA, 1) + 1 To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To i - 1
    If Abs(matrixA(i, j)) > maxAbsInLower2 Then maxAbsInLower2 = Abs(matrixA(i, j))
  Next j
Next i
End Function

Function arrNorm2(matrixA() As Double) As Double
Dim i As Long

For i = LBound(matrixA, 1) To UBound(matrixA, 1): arrNorm2 = arrNorm2 + matrixA(i) * matrixA(i): Next i

arrNorm2 = Sqr(arrNorm2)
End Function

' ========== EIGENVECTORS SOLVING HOMOGENEOUS LINEAR SYSTEM ==========
Function arrEigenVectors(matrixA() As Double) As Double()
Dim i As Integer, j As Integer
Dim eigVectors() As Double, eigValues() As Double, result() As Double

On Error GoTo ErrHandler

If Not arrIsSquare(matrixA) Then Err.Raise 5, "AlgeLin.bas", "Non-square matrix"

ReDim result(LBound(matrixA, 1) To UBound(matrixA, 1), LBound(matrixA, 2) To UBound(matrixA, 2))

eigValues = arrEigenValues(matrixA)

For i = LBound(eigValues, 1) To UBound(eigValues, 1)
  eigVectors = matrixA
  addScaledIdentity eigVectors, -eigValues(i, LBound(eigValues, 2))
  eigVectors = luSolverHomo(eigVectors)

  For j = LBound(result, 1) To UBound(result, 1)
    result(j, i - LBound(eigValues, 1) + LBound(result, 2)) = eigVectors(j, LBound(eigVectors, 2))
  Next
Next

CleanExit:
  arrEigenVectors = result
ErrHandler:
  mLogger.msgErrorFromErr , "arrEigenVectors"
End Function

'=================================================
'******************* Printing ********************
'=================================================
Sub arrDump(matrixA() As Double, ws As Worksheet, Optional row As Integer = 2, Optional col As Integer = 2)
If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("XMatrix")

Dim i As Integer, j As Integer

Dim R0 As Integer: R0 = LBound(matrixA, 1)
Dim c0 As Integer: c0 = LBound(matrixA, 2)

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    ws.Cells(i - R0 + row, j - c0 + col).value = matrixA(i, j)
  Next j
Next i
End Sub

Sub arrDump1D(matrixA() As Double, ws As Worksheet, Optional row As Integer = 2, Optional col As Integer = 2)
If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("XMatrix")

Dim i As Integer

Dim R0 As Integer: R0 = LBound(matrixA, 1)

For i = LBound(matrixA, 1) To UBound(matrixA, 1)
   ws.Cells(i - R0 + row, col).value = matrixA(i)
Next i
End Sub

Function arrToString(matrixA() As Double, Optional numFormat As eNumberFormat = NumberFormat_automatic) As String
Dim i As Integer, j As Integer
Dim maxIntLen As Integer, maxIntLenFirstCol As Integer, numSpaces As Integer
Dim sInt As String, sNum As String
Dim pos As String

Dim decSep As String
decSep = Application.International(xlDecimalSeparator)

'--- Find max integer lengths ---
For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    sNum = numberToString(matrixA(i, j), numFormat)

    pos = InStr(1, sNum, decSep)

    If pos > 0 Then
      sInt = Left(sNum, pos - 1)
    Else
      sInt = sNum
    End If

    If Len(sInt) > maxIntLen Then maxIntLen = Len(sInt)

    If j = LBound(matrixA, 2) And Len(sInt) > maxIntLenFirstCol Then maxIntLenFirstCol = Len(sInt)
  Next
Next

'--- Build output ---
For i = LBound(matrixA, 1) To UBound(matrixA, 1)
  For j = LBound(matrixA, 2) To UBound(matrixA, 2)
    sNum = numberToString(matrixA(i, j), numFormat)

    pos = InStr(1, sNum, decSep)
    If pos > 0 Then
      sInt = Left(sNum, pos - 1)
    Else
      sInt = sNum
    End If

    If j = LBound(matrixA, 2) Then
      numSpaces = maxIntLenFirstCol - Len(sInt)
    Else
      numSpaces = maxIntLen - Len(sInt) + 1
    End If

    arrToString = arrToString & String(numSpaces, " ") & sNum
  Next
  arrToString = arrToString & vbNewLine
Next
End Function


