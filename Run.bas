Attribute VB_Name = "Run"
'@Folder("VBAProject")
'===============================================================================
' Purpose: Test harness for XMatrix class operations.
' Description: Demonstrates allocation, randomization, cloning, and various matrix
'              operations such as eigenvalue computation using the XMatrix class.
' Parameters: None
' Return: None
' Dependencies: Requires XMatrix class with methods Allocate, MRandDouble, Clone,
'               ToString, MEigenValues, etc.
'===============================================================================

Option Explicit

'===============================================================================
' Purpose: Entry point for matrix operation tests.
' Description: Allocates matrices, fills them with random values, clones, and
'              prints results of operations.
' Parameters: None
' Return: None
' Dependencies: XMatrix class and its methods.
'===============================================================================

Sub XDoubleUsageExamples()
Dim xd As XDouble
Set xd = New XDouble

' Set and get the name property
xd.name = "Pressure"
Debug.Print "Name: " & xd.name

' Set and get the unit property
xd.Unit = "Pa"
Debug.Print "Unit: " & xd.Unit

' Set and get the format property
xd.Frmt = "#,##0.00"
Debug.Print "Format: " & xd.Frmt

' Set and get the value property
xd.Value = 101325
Debug.Print "Value: " & xd.Value

' Get formatted string representation
Debug.Print "ToString: " & xd.ToString

' Clone the object
Dim xdClone As XDouble
Set xdClone = xd.Clone

' Output cloned object's formatted value
Debug.Print xdClone.ToStringWithPropName

' Event handling (declare WithEvents XDouble in a class module)
' Example:

' Private WithEvents myXD As XDouble
' Private Sub myXD_ChangedValue(ByVal oldValue As Double, ByVal newValue As Double)
'     Debug.Print "Value changed from " & oldValue & " to " & newValue
' End Sub
End Sub

Sub XMatrixUsageExamples()
Dim matA As XMatrix, matB As XMatrix
Dim arr() As Double: ReDim arr(2, 2)

' Create new matrices
Set matA = New XMatrix
Set matB = New XMatrix

' Fill array with values
arr(0, 0) = 1: arr(0, 1) = 2: arr(0, 2) = 3
arr(1, 0) = 4: arr(1, 1) = 5: arr(1, 2) = 6
arr(2, 0) = 7: arr(2, 1) = 8: arr(2, 2) = 9

' Initialize matA from array
matA.MFromArray = arr

' Access internal matrix array
Dim internalArr As Variant
internalArr = matA.MToArray

' Access single element
Debug.Print "matA(1,1) value: "; matA.Mij(1, 1).Value & vbNewLine

' Get number of rows and columns
Debug.Print "matA Rows: "; matA.Rows & vbNewLine
Debug.Print "matA Cols: "; matA.Cols & vbNewLine

' Convert matrix to array
Dim arrOut As Variant
arrOut = matA.MToArray
Debug.Print "Array(1,1): "; arrOut(1, 1) & vbNewLine

' Allocate a 6x6 matrix for variable matA. It overwrites existing matA matrix.
matA.Allocate 5, 5

' Generate a symmetric random matrix with double values between 1 and 10. Matrix size of the new matrix is the same as matA and overwrite the existing matA.
Debug.Print "Random Symmetric Double Matrix:"
Debug.Print matA.MRandDouble(1, 10, True, False).ToString

' Generate a new non-symmetric random matrix with integer values between 1 and 10. Matrix size of the new matrix is the same as matA.
Set matB = matA.MRandInteger(1, 10, False, True)
Debug.Print "Random Non-Symmetric Integer Matrix:"
Debug.Print matB.ToString

' Generate a new 5-diagonal banded (2 diagonals above and 2 diagonals below the main diagonal) symmetric random matrix with double values between 1 and 10. Matrix size of the new matrix is the same as matA.
Debug.Print "5-Diagonal Banded Random Double Matrix:"
Debug.Print matA.MRandTriDouble(1, 10, 2, True, True).ToString

' Generate a new 7-banded (3 diagonals above and 3 diagonals below the main diagonal) non-symmetric random matrix with integer values between 1 and 10. Matrix size of the new matrix is the same as matA.
Debug.Print "7-Diagonal Banded Random Integer Matrix:"
Debug.Print matA.MRandTriInteger(1, 10, 3, False, True).ToString

' Create a new Identity matrix scaled by 1.4.
Debug.Print "Identity Matrix * 1.4:"
Debug.Print matA.MXidentity(1.4, True).ToString

' Trace of matB
Debug.Print "Trace of matB: "; matB.MTrace & vbNewLine

' Characteristic polynomial coefficients of matA
Debug.Print "Characteristic Polynomial Coefficients of matA:"
Debug.Print matA.MPolynom.ToString

' Eigenvalues of matA
Debug.Print "Eigenvalues of matA (QR-Algorithm):"
Debug.Print matA.MEigenValuesQR.ToString

' Eigenvalues of matA
Debug.Print "Eigenvalues of matA (LU-Algorithm):"
Debug.Print matA.MEigenValues.ToString

' Eigenvalues of matA
Debug.Print "Eigenvectors of matA:"
Debug.Print matA.MEigenVectorsQR.ToString

' Transpose of matB
Debug.Print "Transpose of matB:"
Debug.Print matB.MTran(True).ToString

' Addition of matA and matB. Overwrite matA with the result.
Debug.Print "Addition of matA and matB:"
Debug.Print matA.MAdd(matB, False).ToString

' Subtraction of matA and matB. Overwrite matA with the result.
Debug.Print "Subtraction of matA and matB:"
Debug.Print matA.MSub(matB, False).ToString

' Multiplication of matA and matB.
Debug.Print "Multiplication of matA and matB:"
Debug.Print matA.MMult(matB, True).ToString

' Symmetry check
Debug.Print "Is matA symmetric? "; matA.IsSymmetric & vbNewLine
Debug.Print "Is matB symmetric? "; matB.IsSymmetric & vbNewLine

' Permutation matrix for pivoting
Debug.Print "Permutation Matrix of matA:"
Debug.Print matA.MPivot.ToString

' LU solver with pivoting. Overwrite matB with the result.
Debug.Print "LU Solver (matA*X=matB, for each column of matB) With Pivoting:"
Debug.Print matA.LUSolverWithPivoting(matB).ToString

' LU solver without pivoting. Overwrite matB with the result.
Debug.Print "LU Solver (matA*X=matB, for each column of matB) Without Pivoting:"
Debug.Print matA.LUSolver(matB).ToString

' LU decomposition (lower and upper)
Debug.Print "LU Lower:"
Debug.Print matA.LU(False).ToString
Debug.Print "LU Upper:"
Debug.Print matA.LU(True).ToString

' String representation with 5 decimal places
Debug.Print "String Representation with 5 Decimal Places:"
Debug.Print matA.ToString("#,##0.00000")

' Copy matrix contents
Set matA.Copy = matB
Debug.Print "matA after Copy from matB:"
Debug.Print matA.ToString

' Deep clone
Dim matClone As XMatrix
Set matClone = matA.Clone
Debug.Print "Cloned Matrix:"
Debug.Print matClone.ToString
End Sub
