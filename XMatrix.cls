VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject.Utilities")
'==========================================================================================
' XMatrix.cls
' A class for matrix operations using XDouble objects.
' Provides allocation, arithmetic, randomization, LU decomposition, eigenvalue calculation,
' and utility functions.
'
'References:
'  FRANCO, Neide Maria Bertoldi. Cálculo numérico. 1. ed. São Paulo: Pearson Universidades, 2006.
'==========================================================================================
Option Explicit

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.
'   Stores the matrix elements.
' Description:
'   Private storage for the matrix as a 2D array of XDouble.
' Dependencies:
'   m_Mij.
'   XDouble class.
'==========================================================================================
Private m_Mij() As XDouble
Private m_recalculateEigen As Boolean

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.
'   Allocates the matrix.
' Description:
'   Allocates the matrix with the given upper bounds for rows and columns.
'   Initializes each element as a new XDouble.
' Parameters:
'   rowUBound - upper bound for rows
'   colUBound - upper bound for columns
' Return:
'   None.
' Dependencies:
'   m_Mij.
'   XDouble class.
'==========================================================================================
Sub Allocate(rowUBound As Integer, colUBound As Integer)
Dim i As Integer, j As Integer

' Allocate the 2D array for the matrix
ReDim m_Mij(rowUBound, colUBound)

' Initialize each element as a new XDouble object
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    Set m_Mij(i, j) = New XDouble
  Next
Next
End Sub

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.
' Description:
'   Returns the internal 2D array of XDouble.
' Parameters:
'   None.
' Return:
'   2D array of XDouble.
' Dependencies:
'   m_Mij.
'==========================================================================================
Property Get M() As XDouble()
' Return the internal matrix array
M = m_Mij
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.XDouble at a position.
' Description:
'   Returns the XDouble at the specified row and column.
' Parameters:
'   row - Row index; col - Column index.
' Return:
'   XDouble at (row, col).
' Dependencies:
'   m_Mij.
'==========================================================================================
'@DefaultMember
Property Get Mij(ByVal row As Integer, ByVal col As Integer) As XDouble
Attribute Mij.VB_UserMemId = 0
' Return the XDouble at the given position
Set Mij = m_Mij(row, col)
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Gets number of rows.
' Description:
'   Returns the number of rows in the matrix.
' Parameters:
'   None.
' Return:
'   Integer, number of rows.
' Dependencies:
'   m_Mij.
'==========================================================================================
Property Get Rows() As Integer
' Calculate number of rows
Rows = UBound(m_Mij, 1) - LBound(m_Mij, 1) + 1
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Gets number of columns.
' Description:
'   Returns the number of columns in the matrix.
' Parameters:
'   None.
' Return:
'   Integer, number of columns.
' Dependencies:
'   m_Mij.
'==========================================================================================
Property Get Cols() As Integer
' Calculate number of columns
Cols = UBound(m_Mij, 2) - LBound(m_Mij, 2) + 1
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Sets the matrix from an array.
' Description:
'   Sets the matrix values from a 1D or 2D array of Double.
'   Allocates the matrix to match the dimensions of the input array and copies
'   the values from the input array into the matrix. If the input is a 1D array,
'   the matrix will have one column; if 2D, the matrix will match both dimensions.
' Parameters:
'   MArray - 1D or 2D array of Double containing the values to set.
' Return:
'   None.
' Dependencies:
'   m_Mij.Allocate, m_Mij.
'==========================================================================================
Property Let MFromArray(MArray() As Double)
Dim i As Integer, j As Integer

If ArrayDimensions(MArray) = 1 Then

  ' Allocate the matrix to match the input array size
  Allocate UBound(MArray, 1), 0

  ' Copy values from the input array to the matrix
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    m_Mij(i, 0) = MArray(i)
  Next

Else

  ' Allocate the matrix to match the input array size
  Allocate UBound(MArray, 1), UBound(MArray, 2)

  ' Copy values from the input array to the matrix
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      m_Mij(i, j) = MArray(i, j)
    Next
  Next
End If
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Determines the number of dimensions of an array.
' Description:
'   Returns the number of dimensions of the given array by attempting to access
'   the upper bound for increasing dimension indices until an error occurs.
' Parameters:

'   arr - array of Double whose dimensions are to be determined.
' Return:
'   Integer, the number of dimensions of the array.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Function ArrayDimensions(arr() As Double) As Integer
Dim dimCount As Integer
On Error GoTo EndLoop
Do
  dimCount = dimCount + 1
  Dim temp As Long
  temp = UBound(arr, dimCount)
Loop
EndLoop:
ArrayDimensions = dimCount - 1
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Converts the matrix to an array.
' Description:
'   Converts the matrix to a 2D array of Double.
' Parameters:
'   None.
' Return:
'   2D array of Double.
' Dependencies:
'   m_Mij.
'==========================================================================================
Function MToArray() As Double()
Dim i As Integer, j As Integer
Dim A() As Double

' Allocate output array
ReDim A(UBound(m_Mij, 1), UBound(m_Mij, 2))

' Copy values from the matrix to the output array
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    A(i, j) = m_Mij(i, j)
  Next
Next

MToArray = A
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Generates random Double matrix.
' Description:
'   Generates a matrix with random Double values between min and max.
'   Optionally the returned matrix can be symmetric.
'   The returned matrix can override the current matrix or be a new one.
' Parameters:
'   min, max - range;
'   sym - symmetric;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.RndBtwnDouble, Allocate, Mij.
'==========================================================================================
Property Get MRandDouble(min As Double, max As Double, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MRandDouble = New XMatrix
  MRandDouble.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandDouble = Me
End If

' Only allow symmetry if matrix is square
If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  ' Fill upper triangle and mirror to lower
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        MRandDouble.Mij(i, j) = RndBtwnDouble(min, max)
      Else
        MRandDouble.Mij(i, j) = MRandDouble.Mij(j, i)
      End If
    Next
  Next
Else
  ' Fill all elements randomly
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MRandDouble.Mij(i, j) = RndBtwnDouble(min, max)
    Next
  Next
End If
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Generates random Integer matrix.
' Description:
'   Generates a matrix with random Integer values between min and max. Optionally
'   the returned matrix can be symmetric.
'   The returned matrix can override the current matrix or be a new one.
' Parameters:
'   min, max - range;
'   sym - symmetric;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.RndBtwnInteger, Allocate, Mij.
'==========================================================================================
Property Get MRandInteger(min As Integer, max As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MRandInteger = New XMatrix
  MRandInteger.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandInteger = Me
End If

' Only allow symmetry if matrix is square
If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  ' Fill upper triangle and mirror to lower
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        MRandInteger.Mij(i, j) = RndBtwnInteger(min, max)
      Else
        MRandInteger.Mij(i, j) = MRandInteger.Mij(j, i)
      End If
    Next
  Next
Else
  ' Fill all elements randomly
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MRandInteger.Mij(i, j) = RndBtwnInteger(min, max)
    Next
  Next
End If
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Generates banded random Double matrix.
' Description:
'   Generates a banded matrix with random Double values within a band
'   of width tri. Optionally the returned matrix can be symmetric.
'   The returned matrix can override the current matrix or be a new one.
' Parameters:
'   min, max - range;
'   tri - band width;
'   sym - symmetric;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.RndBtwnDouble, Allocate, Mij.
'==========================================================================================
Property Get MRandTriDouble(min As Double, max As Double, tri As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MRandTriDouble = New XMatrix
  MRandTriDouble.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandTriDouble = Me
End If

' Only allow symmetry if matrix is square
If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  ' Fill upper triangle within band and mirror to lower
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        If j > i - tri - 1 And j <= i + tri Then
          MRandTriDouble.Mij(i, j) = RndBtwnDouble(min, max)
        Else
          MRandTriDouble.Mij(i, j) = 0
        End If
      Else
        MRandTriDouble.Mij(i, j) = MRandTriDouble.Mij(j, i)
      End If
    Next
  Next
Else
  ' Fill all elements within band randomly
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If j >= i - tri And j <= i + tri Then
        MRandTriDouble.Mij(i, j) = RndBtwnDouble(min, max)
      Else
        MRandTriDouble.Mij(i, j) = 0
      End If
    Next
  Next
End If
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Generates banded random Integer matrix.
' Description:
'   Generates a banded matrix with random Integer values within a band
'   of width tri. Optionally the returned matrix can be symmetric.
'   The returned matrix can override the current matrix or be a new one.
' Parameters:
'   min, max - range;
'   tri - band width;
'   sym - symmetric;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.RndBtwnInteger, Allocate, Mij.
'==========================================================================================
Property Get MRandTriInteger(min As Integer, max As Integer, tri As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MRandTriInteger = New XMatrix
  MRandTriInteger.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandTriInteger = Me
End If

' Only allow symmetry if matrix is square
If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  ' Fill upper triangle within band and mirror to lower
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        If j > i - tri - 1 And j <= i + tri Then
          MRandTriInteger.Mij(i, j) = RndBtwnInteger(min, max)
        Else
          MRandTriInteger.Mij(i, j) = 0
        End If
      Else
        MRandTriInteger.Mij(i, j) = MRandTriInteger.Mij(j, i)
      End If
    Next
  Next
Else
  ' Fill all elements within band randomly
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If j >= i - tri And j <= i + tri Then
        MRandTriInteger.Mij(i, j) = RndBtwnInteger(min, max)
      Else
        MRandTriInteger.Mij(i, j) = 0
      End If
    Next
  Next
End If
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.identity matrix.
' Description:
'   Returns an identity matrix. Optionally the returned matrix can be scaled
'   by x. The returned matrix can override the current matrix or be a new one.
' Parameters:
'   x - scale;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, Mij.
'==========================================================================================
Property Get MXidentity(Optional x As Double = 1, Optional newMat As Boolean = True) As XMatrix
Dim i As Integer

' Create new or reuse current matrix
If newMat Then
  Set MXidentity = New XMatrix
  MXidentity.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
  Set MXidentity = Me
End If

' Set diagonal elements to 1*x
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  MXidentity.Mij(i, i) = 1 * x
Next
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.race of matrix.
' Description:
'   Returns the trace (sum of diagonal elements) of the matrix.
' Parameters:
'   None.
' Return:
'   Double, trace value.
' Dependencies:
'   m_Mij.
'==========================================================================================
Function MTrace() As Double
Dim i As Integer

' Sum diagonal elements
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  MTrace = MTrace + m_Mij(i, i)
Next
End Function

'=========================================================================================
' Purpose:
'   Performs QR Decomposition of a square matrix A using the Gram-Schmidt process.
' Description:
'   Decomposes the input matrix A into an orthogonal matrix Q and an upper triangular
'   matrix R such that A = Q * R. This implementation uses the classical Gram-Schmidt
'   process.
' Parameters:
'   A - Input square matrix (2D Variant array, size n x n)
'   Q - Output orthogonal matrix (2D Variant array, size n x n)
'   R - Output upper triangular matrix (2D Variant array, size n x n)
' Return:
'   None. Q and R are modified in place.
' Dependencies:
'   None.
'=========================================================================================
Private Sub QRDecomposition(A() As Double, Q() As Double, R() As Double)
Dim n As Integer, i As Integer, j As Integer, k As Integer
n = UBound(A, 1)
Dim v() As Double, qj() As Double
ReDim Q(0 To n, 0 To n)
ReDim R(0 To n, 0 To n)

For j = 0 To n
  ReDim v(0 To n)
  For i = 0 To n
    v(i) = A(i, j)
  Next i

  For k = 0 To j - 1
    ' Only allocate qj if needed
    ReDim qj(0 To n)
    For i = 0 To n
      qj(i) = Q(i, k)
    Next i
    R(k, j) = 0
    For i = 0 To n
      R(k, j) = R(k, j) + Q(i, k) * A(i, j)
    Next i
    For i = 0 To n
      v(i) = v(i) - R(k, j) * Q(i, k)
    Next i
  Next k

  R(j, j) = 0
  For i = 0 To n
    R(j, j) = R(j, j) + v(i) * v(i)
  Next i

  R(j, j) = Sqr(R(j, j))
  For i = 0 To n
    Q(i, j) = v(i) / R(j, j)
  Next i
Next j
End Sub

'=========================================================================================
' Purpose:
'   Computes the eigenvalues and eigenvectors of a real square matrix using the QR
'   algorithm.
' Description:
'   Iteratively applies QR decomposition to the input matrix A to approximate its
'   eigenvalues (as the diagonal elements of the resulting matrix) and eigenvectors
'   (as the columns of the accumulated orthogonal transformation). Convergence is
'   determined by the sum of the absolute values of the sub-diagonal elements.
' Parameters:
'   A           - Input square matrix (2D Variant array, size n x n)
'   eigenvalues - Output array of eigenvalues (1D Variant array, size n)
'   eigenvectors- Output matrix of eigenvectors (2D Variant array, size n x n)
' Return:
'   None. eigenvalues and eigenvectors are modified in place.
' Dependencies:
'   QRDecomposition
'=========================================================================================
Private Sub QREigen(A() As Double, eigenValues() As Double, eigenVectors() As Double)
Const tol As Double = 0.000001
Const MaxIter As Integer = 100

Dim n As Integer, i As Integer, j As Integer, k As Integer, iter As Integer
n = UBound(A, 1)
Dim Ak() As Double, Q() As Double, R() As Double, Qprod() As Double, temp() As Double
ReDim Ak(0 To n, 0 To n)
ReDim Q(0 To n, 0 To n)
ReDim R(0 To n, 0 To n)
ReDim Qprod(0 To n, 0 To n)

' Copy A to Ak
For i = 0 To n
  For j = 0 To n
    Ak(i, j) = A(i, j)
    Qprod(i, j) = IIf(i = j, 1, 0)
  Next j
Next i

For iter = 1 To MaxIter
  Call QRDecomposition(Ak, Q, R)
  ' Ak = R * Q
  ReDim temp(0 To n, 0 To n)
  For i = 0 To n
    For j = 0 To n
      temp(i, j) = 0
      For k = 0 To n
        temp(i, j) = temp(i, j) + R(i, k) * Q(k, j)
      Next k
    Next j
  Next i
  For i = 0 To n
    For j = 0 To n
      Ak(i, j) = temp(i, j)
    Next j
  Next i
  ' Qprod = Qprod * Q
  ReDim temp(0 To n, 0 To n)
  For i = 0 To n
    For j = 0 To n
      temp(i, j) = 0
      For k = 0 To n
        temp(i, j) = temp(i, j) + Qprod(i, k) * Q(k, j)
      Next k
    Next j
  Next i
  For i = 0 To n
    For j = 0 To n
      Qprod(i, j) = temp(i, j)
    Next j
  Next i
  ' Check for convergence
  Dim offDiag As Double
  offDiag = 0
  For i = 1 To n
    For j = 0 To i - 1
      offDiag = offDiag + Abs(Ak(i, j))
    Next j
  Next i
  If offDiag < tol Then Exit For
Next iter

' Eigenvalues are the diagonal of Ak
ReDim eigenValues(0 To n)
For i = 0 To n
  eigenValues(i) = Ak(i, i)
Next i

' Eigenvectors are columns of Qprod
ReDim eigenVectors(0 To n, 0 To n)
For i = 0 To n
  For j = 0 To n
    eigenVectors(i, j) = Qprod(i, j)
  Next j
Next i
End Sub

Public Function MEigenVectorsQR() As XMatrix
Set MEigenVectorsQR = New XMatrix
MEigenVectorsQR.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)

Dim eigenVectors() As Double: Dim eigenValues() As Double
ReDim eigenVectors(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim eigenValues(UBound(m_Mij, 1), 0)

Call QREigen(MToArray, eigenValues, eigenVectors)
MEigenVectorsQR.MFromArray = eigenVectors
End Function

Public Function MEigenValuesQR() As XMatrix
Set MEigenValuesQR = New XMatrix
MEigenValuesQR.Allocate UBound(m_Mij, 1), 0

Dim eigenVectors() As Double: Dim eigenValues() As Double
ReDim eigenVectors(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim eigenValues(UBound(m_Mij, 1), 0)

Call QREigen(MToArray, eigenValues, eigenVectors)
MEigenValuesQR.MFromArray = eigenValues
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Characteristic polynomial coefficients.
' Description:
'   Computes the coefficients of the characteristic polynomial of the matrix.
' Parameters:
'   None.
' Return:
'   XMatrix (coefficients).
' Dependencies:
'   m_Mij.MTrace, MMult, MAdd, MXidentity, Clone.
'==========================================================================================
Function MPolynom() As XMatrix
Dim i As Integer
Dim eValues() As Double
Dim A As XMatrix

' Allocate output for coefficients
Set MPolynom = New XMatrix
ReDim eValues(UBound(m_Mij, 1), 0)
Set A = Clone

' First coefficient is -trace
eValues(0, 0) = -MTrace

' Recursively compute coefficients
For i = LBound(m_Mij, 1) + 1 To UBound(m_Mij, 1)
  Set A = MMult(A.MAdd(A.MXidentity(eValues(i - 1, 0))), True)
  eValues(i, 0) = -1 / (i + 1) * A.MTrace
Next

' Store coefficients in output matrix
MPolynom.MFromArray = eValues
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Approximates eigenvalues.
' Description:
'   Approximates the eigenvalues of the matrix using the LU decomposition method.
' Parameters:
'   None.
' Return:
'   XMatrix (eigenvalues).
' Dependencies:
'   m_Mij.MToArray, LUDecompArr, MMultArr, MaxAbsValueInLower, Mij.
'==========================================================================================
Function MEigenValues() As XMatrix
Set MEigenValues = New XMatrix
MEigenValues.Allocate UBound(m_Mij, 1), 0

Const Error As Double = 0.000001
Const MaxIter As Double = 10000#

Dim i As Integer
Dim L() As Double, U() As Double, A() As Double
Dim T As New XMatrix
ReDim L(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim U(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim A(UBound(m_Mij, 1), UBound(m_Mij, 2))

' Copy matrix to array
A = MToArray

' Iteratively apply LU decomposition and multiply U*L
Do
  LUDecompArr A, L, U
  A = MMultArr(U, L)
  i = i + 1
Loop Until MaxAbsValueInLower(A) < Error Or i > MaxIter

' If not converged, exit
If MaxAbsValueInLower(A) > Error Then Exit Function

' Diagonal elements approximate eigenvalues
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  MEigenValues.Mij(i, 0) = A(i, i)
Next
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Max abs value in lower triangle.
' Description:
'   Returns the maximum absolute value in the lower triangle of a matrix.
' Parameters:
'   - 2D array of Double.
' Return:
'   Double.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Private Function MaxAbsValueInLower(A() As Double) As Double
Dim i As Integer, j As Integer

' Scan lower triangle for largest absolute value
For i = LBound(A, 1) To UBound(A, 1)
  For j = LBound(A, 2) To i - 1
    If MaxAbsValueInLower < Abs(A(i, j)) Then MaxAbsValueInLower = Abs(A(i, j))
  Next
Next
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.ranspose.
' Description:
'   Returns the transpose of the matrix. The returned matrix can
'   override the current matrix or be a new one.
' Parameters:
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, Mij, Copy.
'==========================================================================================
Property Get MTran(Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Allocate new matrix with swapped dimensions
Set MTran = New XMatrix
MTran.Allocate UBound(m_Mij, 2), UBound(m_Mij, 1)

' Copy transposed elements
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MTran.Mij(j, i) = m_Mij(i, j)
  Next
Next

' Optionally copy result to current matrix
If Not newMat Then Set Copy = MTran
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Adds matrices.
' Description:
'   Adds another matrix to this matrix. The returned matrix can
'   override the current matrix or be a new one.
' Parameters:
'   mat - matrix to add;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, Mij.
'==========================================================================================
Property Get MAdd(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MAdd = New XMatrix
  MAdd.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MAdd = Me
End If

' Elementwise addition
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MAdd.Mij(i, j) = Mij(i, j) + mat.Mij(i, j)
  Next
Next
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Subtracts matrices.
' Description:
'   Subtracts another matrix from this matrix. The returned matrix can
'   override the current matrix or be a new one.
' Parameters:
'   mat - matrix to subtract;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, Mij.
'==========================================================================================
Property Get MSub(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

' Create a new matrix if requested, otherwise use current
If newMat Then
  Set MSub = New XMatrix
  MSub.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MSub = Me
End If

' Elementwise subtraction
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MSub.Mij(i, j) = Mij(i, j) - mat.Mij(i, j)
  Next
Next
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Multiplies matrices.
' Description:
'   Multiplies this matrix by another matrix. The returned matrix can
'   override the current matrix or be a new one.
' Parameters:
'   mat - matrix to multiply;
'   newMat - return new matrix.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, Mij, Copy.
'==========================================================================================
Property Get MMult(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer, k As Integer

' Allocate result matrix
Set MMult = New XMatrix
MMult.Allocate UBound(m_Mij, 1), UBound(mat.M, 2)

' Matrix multiplication: sum over j
For k = LBound(mat.M, 2) To UBound(mat.M, 2)
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MMult.Mij(i, k) = MMult.Mij(i, k) + Mij(i, j) * mat.Mij(j, k)
    Next
  Next
Next

' Optionally copy result to current matrix
If Not newMat Then Set Copy = MMult
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Multiplies arrays as matrices. This is only used internally.
' Description:
'   Multiplies two 2D arrays as matrices.
' Parameters:
'   a, b - 2D arrays of Double.
' Return:
'   2D array of Double.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Private Function MMultArr(A() As Double, b() As Double) As Double()
Dim i As Integer, j As Integer, k As Integer
Dim c() As Double

' Allocate result array
ReDim c(UBound(A, 1), UBound(b, 2))

' Matrix multiplication for arrays
For k = LBound(b, 2) To UBound(b, 2)
  For i = LBound(A, 1) To UBound(A, 1)
    For j = LBound(A, 2) To UBound(A, 2)
      c(i, k) = c(i, k) + A(i, j) * b(j, k)
    Next
  Next
Next

MMultArr = c
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Checks symmetry.
' Description:
'   Checks if the matrix is symmetric.
' Parameters:
'   None.
' Return:
'   Boolean.
' Dependencies:
'   m_Mij.
'==========================================================================================
Public Function IsSymmetric() As Boolean
Dim i As Integer, j As Integer

' Assume symmetric unless proven otherwise
IsSymmetric = True

' Matrix must be square
If Rows <> Cols Then
  IsSymmetric = False
  Exit Function
End If

' Check if all (i,j) == (j,i)
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To i - 1
    If m_Mij(i, j).Value <> m_Mij(j, i).Value Then
      IsSymmetric = False
      Exit Function
    End If
  Next
Next
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Computes pivot vector.
' Description:
'   Computes the pivot vector for partial pivoting in LU decomposition.
' Parameters:
'   None.
' Return:
'   Integer array (pivot vector).
' Dependencies:
'   m_Mij.
'==========================================================================================
Private Function Pivot() As Integer()
Dim i As Integer, j As Integer, k As Integer, L As Integer
Dim doAgain As Boolean: doAgain = False
Dim P() As Integer: ReDim P(UBound(m_Mij, 1))

ReDim Pivot(UBound(m_Mij, 1))

' Initialize pivot vector
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1):  P(i) = i: Next

' Pivoting loop
Do
  doAgain = False: L = L + 1
  ' Forward pass
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    For i = j + 1 To UBound(m_Mij, 1)
      If Abs(m_Mij(P(j), j)) < Abs(m_Mij(P(i), j)) Then
        ' Swap in pivot vector
        k = P(j):      P(j) = P(i):      P(i) = k: doAgain = True
      End If
    Next
  Next

  ' Backward pass
  For j = UBound(m_Mij, 2) To LBound(m_Mij, 2) Step -1
    For i = j - 1 To 0 Step -1
      If Abs(m_Mij(P(j), j)) < Abs(m_Mij(P(i), j)) Then
        ' Swap in pivot vector
        k = P(j):      P(j) = P(i):      P(i) = k: doAgain = True
      End If
    Next
  Next
Loop While L < 50 And doAgain

Pivot = P
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Permutation matrix from pivot.
' Description:
'   Returns the permutation matrix corresponding to the pivot vector.
' Parameters:
'   None.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Pivot, Allocate, Mij.
'==========================================================================================
Public Function MPivot() As XMatrix
Set MPivot = New XMatrix
Dim i As Integer

Dim P() As Integer: P = Pivot: MPivot.Allocate UBound(P, 1), UBound(P, 1)

' Build permutation matrix from pivot vector
For i = LBound(P, 1) To UBound(P, 1)
  MPivot.Mij(i, P(i)) = 1
Next
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Solves Ax=b with pivoting.
' Description:
'   Solves the linear system Ax = b using LU decomposition with pivoting.
' Parameters:
'   b - right-hand side matrix.
' Return:
'   XMatrix (solution).
' Dependencies:
'   m_Mij.MPivot, MMult, LUSolver.
'==========================================================================================
Function LUSolverWithPivoting(b As XMatrix) As XMatrix
Dim L() As Double, U() As Double
Dim P As XMatrix: Set P = MPivot
Dim matPivoted As XMatrix
Dim bPivoted As XMatrix

ReDim L(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim U(UBound(m_Mij, 1), UBound(m_Mij, 2))

' Apply permutation, then solve using LU
Set matPivoted = P.MMult(Me, True)
Set bPivoted = P.MMult(b, True)

Set LUSolverWithPivoting = matPivoted.LUSolver(bPivoted)
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Solves Ax=b (no pivot).
' Description:
'   Solves the linear system Ax = b using LU decomposition (no pivoting).
' Parameters:
'   b - right-hand side matrix.
' Return:
'   XMatrix (solution).
' Dependencies:
'   m_Mij.LUDecomp, FowardSub, BackwardSub.
'==========================================================================================
Function LUSolver(b As XMatrix) As XMatrix
Dim L() As Double, U() As Double

ReDim L(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim U(UBound(m_Mij, 1), UBound(m_Mij, 2))

' LU decomposition, then forward and backward substitution
LUDecomp L, U
FowardSub L, b
BackwardSub U, b

Set LUSolver = b
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Forward substitution.
' Description:
'   Performs forward substitution for solving lower-triangular systems.
' Parameters:
'   lowerTri - lower triangular array;
'   b - right-hand side matrix.
' Return:
'   None.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Private Sub FowardSub(lowerTri() As Double, b As XMatrix)
Dim i As Integer, j As Integer, k As Integer

' Forward substitution algorithm
For k = LBound(b.M, 2) To UBound(b.M, 2)
  For i = LBound(lowerTri, 1) To UBound(lowerTri, 1)
    b.Mij(i, k) = b.Mij(i, k)
    For j = LBound(lowerTri, 2) To i - 1
      b.Mij(i, k) = b.Mij(i, k) - lowerTri(i, j) * b.Mij(j, k)
    Next
    b.Mij(i, k) = b.Mij(i, k) / lowerTri(i, i)
  Next
Next
End Sub

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Backward substitution.
' Description:
'   Performs backward substitution for solving upper-triangular systems.
' Parameters:
'   upperTri - upper triangular array;
'   b - right-hand side matrix.
' Return:
'   None.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Private Sub BackwardSub(upperTri() As Double, b As XMatrix)
Dim i As Integer, j As Integer, k As Integer

' Backward substitution algorithm
For k = LBound(b.M, 2) To UBound(b.M, 2)
  For i = UBound(upperTri, 1) To LBound(upperTri, 1) Step -1
    b.Mij(i, k) = b.Mij(i, k)
    For j = i + 1 To UBound(upperTri, 2)
      b.Mij(i, k) = b.Mij(i, k) - upperTri(i, j) * b.Mij(j, k)
    Next
    b.Mij(i, k) = b.Mij(i, k) / upperTri(i, i)
  Next
Next
End Sub

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.LU decomposition.
' Description:
'   Performs LU decomposition of the matrix into lower and upper triangular
'   matrices.
' Parameters:
'   lower, upper - output arrays.
' Return:
'   None.
' Dependencies:
'   m_Mij.
'==========================================================================================
Private Sub LUDecomp(lower() As Double, upper() As Double)
Dim i As Integer, j As Integer, k As Integer

' LU decomposition algorithm
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)

    If i <= j Then
      upper(i, j) = m_Mij(i, j)
      If i = j Then lower(i, j) = 1

      For k = LBound(m_Mij, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, k) * upper(k, j)
      Next
    Else
      lower(i, j) = m_Mij(i, j)
      For k = LBound(m_Mij, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, k) * upper(k, j)
      Next
      lower(i, j) = lower(i, j) / upper(j, j)
    End If
  Next
Next
End Sub

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.LU decomposition for arrays.
' Description:
'   Performs LU decomposition of a 2D array into lower and upper triangular
'   matrices.
' Parameters:
'   a - input array;
'   lower, upper - output arrays.
' Return:
'   None.
' Dependencies:
'   m_Mij.None.
'==========================================================================================
Private Sub LUDecompArr(A() As Double, lower() As Double, upper() As Double)
Dim i As Integer, j As Integer, k As Integer

' LU decomposition for arrays
For i = LBound(A, 1) To UBound(A, 1)
  For j = LBound(A, 2) To UBound(A, 2)

    If i <= j Then
      upper(i, j) = A(i, j)
      If i = j Then lower(i, j) = 1

      For k = LBound(A, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, k) * upper(k, j)
      Next
    Else
      lower(i, j) = A(i, j)
      For k = LBound(A, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, k) * upper(k, j)
      Next
      lower(i, j) = lower(i, j) / upper(j, j)
    End If
  Next
Next
End Sub

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.LU matrix.
' Description:
'   Returns the lower or upper triangular matrix from LU decomposition.
' Parameters:
'   returnUpper - if True, returns upper; else lower.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.LUDecomp, Allocate, MFromArray.
'==========================================================================================
Function LU(Optional returnUpper As Boolean = False) As XMatrix
Set LU = New XMatrix
LU.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)

Dim i As Integer, j As Integer, k As Integer
Dim upper() As Double, lower() As Double

ReDim lower(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim upper(UBound(m_Mij, 1), UBound(m_Mij, 2))

' LU decomposition and return lower or upper
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)

    If i <= j Then
      upper(i, j) = m_Mij(i, j)
      If i = j Then lower(i, j) = 1

      For k = LBound(m_Mij, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, k) * upper(k, j)
      Next
    Else
      lower(i, j) = m_Mij(i, j)
      For k = LBound(m_Mij, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, k) * upper(k, j)
      Next
      lower(i, j) = lower(i, j) / upper(j, j)
    End If
  Next
Next

If returnUpper Then
  LU.MFromArray = upper
Else
  LU.MFromArray = lower
End If
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.String representation.
' Description:
'   Returns a string representation of the matrix, formatted with the given
'   number format.
' Parameters:
'   numFormat - format string.
' Return:
'   String.
' Dependencies:
'   m_Mij.
'==========================================================================================
Function ToString(Optional numFormat As String = "#,##0.00") As String
Dim i As Integer, j As Integer
Dim maxIntLen As Integer, maxIntLen1Col As Integer, numSpaces As Integer
Dim strInteger As String, strNumber As String

' Determine maximum integer length for formatting
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    strNumber = Format(m_Mij(i, j), numFormat)

    strInteger = Split(strNumber, IIf(CInt("0,1") = 0, ",", "."))(0)
    If Len(strInteger) > maxIntLen Then maxIntLen = Len(strInteger)

    If j = LBound(m_Mij, 2) And Len(strInteger) > maxIntLen1Col Then maxIntLen1Col = Len(strInteger)
  Next
Next

' Format each element and align columns for pretty output
For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    strNumber = Format(m_Mij(i, j), numFormat)

    strInteger = Split(strNumber, IIf(CInt("0,1") = 0, ",", "."))(0)

    If j = LBound(m_Mij, 2) Then
      numSpaces = maxIntLen1Col - Len(strInteger)
    Else
      numSpaces = maxIntLen - Len(strInteger) + 1
    End If

    ToString = ToString & String(numSpaces, " ") & strNumber
  Next
  ToString = ToString & vbNewLine
Next
End Function

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Copies matrix contents.
' Description:
'   Copies the contents of another matrix into this matrix.
' Parameters:
'   mat - matrix to copy from.
' Return:
'   None.
' Dependencies:
'   m_Mij.Allocate, m_Mij.
'==========================================================================================
Property Set Copy(mat As XMatrix)
Dim i As Integer, j As Integer

' Allocate and copy all elements
Allocate UBound(mat.M, 1), UBound(mat.M, 2)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    m_Mij(i, j) = mat.Mij(i, j)
  Next
Next
End Property

'==========================================================================================
' Purpose:
'   Returns the internal 2D array.Deep copy of matrix.
' Description:
'   Returns a deep copy of the matrix as a new XMatrix object.
' Parameters:
'   None.
' Return:
'   XMatrix.
' Dependencies:
'   m_Mij.Allocate, m_Mij.
'==========================================================================================
Function Clone() As XMatrix
Set Clone = New XMatrix

Dim i As Integer, j As Integer

' Allocate and copy all elements to new matrix
Clone.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    Clone.Mij(i, j) = m_Mij(i, j)
  Next
Next
End Function
