Attribute VB_Name = "ModVector"
'modVector : Vector module
'Author : C.Dutoit
'Mail   : dutoitc@hotmail.com
'Web    : http://www.dutoitc.fr.st or http://dutoitc.tsx.org
'ICQ    : UIN = 9657082
'MSN    : dutoitc@hotmail.com

Option Explicit

'A Vector
Type TVector
    X As Single
    Y As Single
End Type 'TVector


'Vector addition
Function Add(L As TVector, R As TVector) As TVector
    Dim Vec As TVector  'Attention aux Effets de bords
    Vec.X = L.X + R.X
    Vec.Y = L.Y + R.Y
    Add = Vec
End Function 'Add


'Scalar-Vector multiplication
Function Mul(S As Single, V As TVector) As TVector
    Dim Vec As TVector
    Vec.X = V.X * S
    Vec.Y = V.Y * S
    Mul = Vec
End Function 'Mul


'Return a vector
Function Vec(X As Single, Y As Single) As TVector
    Vec.X = X
    Vec.Y = Y
End Function 'Vec
