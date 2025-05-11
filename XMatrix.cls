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
Option Explicit

Private m_Mij() As XDouble

Sub Allocate(rowUBound As Integer, colUBound As Integer)
Dim i As Integer, j As Integer

ReDim m_Mij(rowUBound, colUBound)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    Set m_Mij(i, j) = New XDouble
  Next
Next
End Sub

Property Get M() As XDouble(): M = m_Mij: End Property

'@DefaultMember
Property Get Mij(ByVal row As Integer, ByVal col As Integer) As XDouble: Set Mij = m_Mij(row, col): End Property
Attribute Mij.VB_UserMemId = 0

Property Get Rows() As Integer: Rows = UBound(m_Mij, 1) - LBound(m_Mij, 1) + 1: End Property

Property Get Cols() As Integer: Cols = UBound(m_Mij, 2) - LBound(m_Mij, 2) + 1: End Property

Property Let MFromArray(MArray() As Double)
Dim i As Integer, j As Integer

Allocate UBound(MArray, 1), UBound(MArray, 2)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    m_Mij(i, j) = MArray(i, j)
  Next
Next
End Property

Function MToArray() As Double()
Dim i As Integer, j As Integer
Dim a() As Double

ReDim a(UBound(m_Mij, 1), UBound(m_Mij, 2))

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    a(i, j) = m_Mij(i, j)
  Next
Next

MToArray = a
End Function

Property Get MRandDouble(min As Double, max As Double, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MRandDouble = New XMatrix
  MRandDouble.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandDouble = Me
End If

If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
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
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MRandDouble.Mij(i, j) = RndBtwnDouble(min, max)
    Next
  Next
End If
End Property

Property Get MRandInteger(min As Integer, max As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MRandInteger = New XMatrix
  MRandInteger.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandInteger = Me
End If

If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
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
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MRandInteger.Mij(i, j) = RndBtwnInteger(min, max)
    Next
  Next
End If
End Property

Property Get MRandTriDouble(min As Double, max As Double, tri As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MRandTriDouble = New XMatrix
  MRandTriDouble.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandTriDouble = Me
End If

If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        If j > i - tri And j < i + tri Then
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
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If j > i - tri And j < i + tri Then
        MRandTriDouble.Mij(i, j) = RndBtwnDouble(min, max)
      Else
        MRandTriDouble.Mij(i, j) = 0
      End If
    Next
  Next
End If
End Property

Property Get MRandTriInteger(min As Integer, max As Integer, tri As Integer, Optional sym As Boolean = True, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MRandTriInteger = New XMatrix
  MRandTriInteger.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MRandTriInteger = Me
End If

If (UBound(m_Mij, 1) - LBound(m_Mij, 1)) <> (UBound(m_Mij, 2) - LBound(m_Mij, 2)) Then sym = False

If sym Then
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If i <= j Then
        If j > i - tri And j < i + tri Then
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
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      If j > i - tri And j < i + tri Then
        MRandTriInteger.Mij(i, j) = RndBtwnInteger(min, max)
      Else
        MRandTriInteger.Mij(i, j) = 0
      End If
    Next
  Next
End If
End Property

Property Get MXidentity(Optional x As Double = 1, Optional newMat As Boolean = True) As XMatrix
Dim i As Integer

If newMat Then
  Set MXidentity = New XMatrix
  MXidentity.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
  Set MXidentity = Me
End If

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  MXidentity.Mij(i, i) = 1 * x
Next
End Property

Function MTrace() As Double
Dim i As Integer

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  MTrace = MTrace + m_Mij(i, i)
Next
End Function

Function MPolynom() As XMatrix
Dim i As Integer
Dim eValues() As Double
Dim a As XMatrix

Set MPolynom = New XMatrix
ReDim eValues(UBound(m_Mij, 1), 0)
Set a = Clone

eValues(0, 0) = -MTrace

For i = LBound(m_Mij, 1) + 1 To UBound(m_Mij, 1)
  Set a = MMult(a.MAdd(a.MXidentity(eValues(i - 1, 0))), True)
  eValues(i, 0) = -1 / (i + 1) * a.MTrace
Next

MPolynom.MFromArray = eValues
End Function

Function MEigenValues() As XMatrix
Set MEigenValues = New XMatrix
MEigenValues.Allocate UBound(m_Mij, 1), 0

Const Error As Double = 0.000001
Const MaxIter As Double = 10000#

Dim i As Integer
Dim L() As Double, U() As Double, a() As Double
Dim T As New XMatrix
ReDim L(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim U(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim a(UBound(m_Mij, 1), UBound(m_Mij, 2))

a = MToArray

Do
  LUDecompArr a, L, U
  a = MMultArr(U, L)
  i = i + 1
Loop Until MaxAbsValueInLower(a) < Error Or i > MaxIter

If MaxAbsValueInLower(a) > Error Then Exit Function

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    MEigenValues.Mij(i, 0) = a(i, i)
Next
End Function

Private Function MaxAbsValueInLower(a() As Double) As Double
Dim i As Integer, j As Integer

For i = LBound(a, 1) To UBound(a, 1)
  For j = LBound(a, 2) To i - 1
    If MaxAbsValueInLower < Abs(a(i, j)) Then MaxAbsValueInLower = Abs(a(i, j))
  Next
Next
End Function

Property Get MTran(Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

Set MTran = New XMatrix
MTran.Allocate UBound(m_Mij, 2), UBound(m_Mij, 1)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MTran.Mij(j, i) = m_Mij(i, j)
  Next
Next

If Not newMat Then Set Copy = MTran
End Property

Property Get MAdd(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MAdd = New XMatrix
  MAdd.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MAdd = Me
End If

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MAdd.Mij(i, j) = Mij(i, j) + mat.Mij(i, j)
  Next
Next
End Property

Property Get MSub(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer

If newMat Then
  Set MSub = New XMatrix
  MSub.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)
Else
  Set MSub = Me
End If

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    MSub.Mij(i, j) = Mij(i, j) - mat.Mij(i, j)
  Next
Next
End Property

Property Get MMult(mat As XMatrix, Optional newMat As Boolean = False) As XMatrix
Dim i As Integer, j As Integer, K As Integer

Set MMult = New XMatrix
MMult.Allocate UBound(m_Mij, 1), UBound(mat.M, 2)

For K = LBound(mat.M, 2) To UBound(mat.M, 2)
  For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
    For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
      MMult.Mij(i, K) = MMult.Mij(i, K) + Mij(i, j) * mat.Mij(j, K)
    Next
  Next
Next

If Not newMat Then Set Copy = MMult
End Property

Private Function MMultArr(a() As Double, b() As Double) As Double()
Dim i As Integer, j As Integer, K As Integer
Dim c() As Double

ReDim c(UBound(a, 1), UBound(b, 2))

For K = LBound(b, 2) To UBound(b, 2)
  For i = LBound(a, 1) To UBound(a, 1)
    For j = LBound(a, 2) To UBound(a, 2)
      c(i, K) = c(i, K) + a(i, j) * b(j, K)
    Next
  Next
Next

MMultArr = c
End Function

Public Function IsSymmetric() As Boolean
Dim i As Integer, j As Integer

IsSymmetric = True

If Rows <> Cols Then
  IsSymmetric = False
  Exit Function
End If

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To i - 1
    If m_Mij(i, j).Value <> m_Mij(j, i).Value Then
      IsSymmetric = False
      Exit Function
    End If
  Next
Next
End Function

Function LUSolver(b As XMatrix) As XMatrix
Dim L() As Double, U() As Double

ReDim L(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim U(UBound(m_Mij, 1), UBound(m_Mij, 2))

LUDecomp L, U
FowardSub L, b
BackwardSub U, b

Set LUSolver = b
End Function

Private Sub FowardSub(lowerTri() As Double, b As XMatrix)
Dim i As Integer, j As Integer, K As Integer

For K = LBound(b.M, 2) To UBound(b.M, 2)
  For i = LBound(lowerTri, 1) To UBound(lowerTri, 1)
    b.Mij(i, K) = b.Mij(i, K)
    For j = LBound(lowerTri, 2) To i - 1
      b.Mij(i, K) = b.Mij(i, K) - lowerTri(i, j) * b.Mij(j, K)
    Next
    b.Mij(i, K) = b.Mij(i, K) / lowerTri(i, i)
  Next
Next
End Sub

Private Sub BackwardSub(upperTri() As Double, b As XMatrix)
Dim i As Integer, j As Integer, K As Integer

For K = LBound(b.M, 2) To UBound(b.M, 2)
  For i = UBound(upperTri, 1) To LBound(upperTri, 1) Step -1
    b.Mij(i, K) = b.Mij(i, K)
    For j = i + 1 To UBound(upperTri, 2)
      b.Mij(i, K) = b.Mij(i, K) - upperTri(i, j) * b.Mij(j, K)
    Next
    b.Mij(i, K) = b.Mij(i, K) / upperTri(i, i)
  Next
Next
End Sub

Private Sub LUDecomp(lower() As Double, upper() As Double)
Dim i As Integer, j As Integer, K As Integer

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)

    If i <= j Then
      upper(i, j) = m_Mij(i, j)
      If i = j Then lower(i, j) = 1

      For K = LBound(m_Mij, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, K) * upper(K, j)
      Next
    Else
      lower(i, j) = m_Mij(i, j)
      For K = LBound(m_Mij, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, K) * upper(K, j)
      Next
      lower(i, j) = lower(i, j) / upper(j, j)
    End If
  Next
Next
End Sub

Private Sub LUDecompArr(a() As Double, lower() As Double, upper() As Double)
Dim i As Integer, j As Integer, K As Integer

For i = LBound(a, 1) To UBound(a, 1)
  For j = LBound(a, 2) To UBound(a, 2)

    If i <= j Then
      upper(i, j) = a(i, j)
      If i = j Then lower(i, j) = 1

      For K = LBound(a, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, K) * upper(K, j)
      Next
    Else
      lower(i, j) = a(i, j)
      For K = LBound(a, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, K) * upper(K, j)
      Next
      lower(i, j) = lower(i, j) / upper(j, j)
    End If
  Next
Next
End Sub

Function LU(Optional returnUpper As Boolean = False) As XMatrix
Set LU = New XMatrix
LU.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)

Dim i As Integer, j As Integer, K As Integer
Dim upper() As Double, lower() As Double

ReDim lower(UBound(m_Mij, 1), UBound(m_Mij, 2))
ReDim upper(UBound(m_Mij, 1), UBound(m_Mij, 2))

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)

    If i <= j Then
      upper(i, j) = m_Mij(i, j)
      If i = j Then lower(i, j) = 1

      For K = LBound(m_Mij, 1) To i - 1
        upper(i, j) = upper(i, j) - lower(i, K) * upper(K, j)
      Next
    Else
      lower(i, j) = m_Mij(i, j)
      For K = LBound(m_Mij, 1) To j - 1
        lower(i, j) = lower(i, j) - lower(i, K) * upper(K, j)
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

Function ToString(Optional numFormat As String = "#,##0.00") As String
Dim i As Integer, j As Integer
Dim maxIntLen As Integer, maxIntLen1Col As Integer, numSpaces As Integer
Dim strInteger As String, strNumber As String

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    strNumber = Format(m_Mij(i, j), numFormat)

    strInteger = Split(strNumber, IIf(CInt("0,1") = 0, ",", "."))(0)
    If Len(strInteger) > maxIntLen Then maxIntLen = Len(strInteger)

    If j = LBound(m_Mij, 2) And Len(strInteger) > maxIntLen1Col Then maxIntLen1Col = Len(strInteger)
  Next
Next

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

Property Set Copy(mat As XMatrix)
Dim i As Integer, j As Integer

Allocate UBound(mat.M, 1), UBound(mat.M, 2)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    m_Mij(i, j) = mat.Mij(i, j)
  Next
Next
End Property

Function Clone() As XMatrix
Set Clone = New XMatrix

Dim i As Integer, j As Integer

Clone.Allocate UBound(m_Mij, 1), UBound(m_Mij, 2)

For i = LBound(m_Mij, 1) To UBound(m_Mij, 1)
  For j = LBound(m_Mij, 2) To UBound(m_Mij, 2)
    Clone.Mij(i, j) = m_Mij(i, j)
  Next
Next
End Function
