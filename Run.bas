Attribute VB_Name = "Run"
'@Folder("VBAProject")
Option Explicit

Sub TestMatrix()
Dim a As New XMatrix
Dim x As New XMatrix
Dim b As New XMatrix

a.Allocate 9, 9
x.Allocate 2, 0

Set a = a.MRandDouble(-100#, 100#, True)
Set x = x.MRandDouble(0, 1000#)
Set b = x.Clone

'A.Mij(0, 0) = 2: A.Mij(0, 1) = 1: A.Mij(0, 2) = 3
'A.Mij(1, 0) = 5: A.Mij(1, 1) = 5: A.Mij(1, 2) = 8
'A.Mij(2, 0) = 4: A.Mij(2, 1) = 2: A.Mij(2, 2) = 7

Debug.Print a.ToString("#,##0.000")
'Debug.Print x.ToString
'Debug.Print A.LU.MMult(A.LU(True)).ToString
'Debug.Print A.LU.ToString
'Debug.Print A.LU(True).ToString
'Debug.Print A.LUSolver(x).ToString
'Debug.Print A.MMult(x).MSub(b).ToString
Debug.Print a.MEigenValues.ToString("#,##0.0000")
End Sub
