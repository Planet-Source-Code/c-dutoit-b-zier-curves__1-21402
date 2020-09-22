Attribute VB_Name = "modBezier"
'modBezier
'Author : C.Dutoit
'Mail   : dutoitc@hotmail.com
'Web    : http://www.dutoitc.fr.st or http://dutoitc.tsx.org
'ICQ    : UIN = 9657082
'MSN    : dutoitc@hotmail.com

Option Explicit

'Draw a Bezier curve on pct with the specified control points. Depth is Recursive depth for Bezier calc.
Sub Draw(pct As PictureBox, P0 As TVector, P1 As TVector, P2 As TVector, P3 As TVector, Depth As Byte)
    Dim nP0 As TVector
    Dim nP1 As TVector
    Dim nP2 As TVector
    Dim nP3 As TVector
    
    'Depth :
    If Depth > 0 Then
        'left
        nP0 = P0
        nP1 = Add(Mul(1 / 2, P0), Mul(1 / 2, P1))
        nP2 = Add(Add(Mul(1 / 4, P0), Mul(1 / 2, P1)), Mul(1 / 4, P2))
        nP3 = Add(Add(Add(Mul(1 / 8, P0), Mul(3 / 8, P1)), Mul(3 / 8, P2)), Mul(1 / 8, P3))
        Draw pct, nP0, nP1, nP2, nP3, Depth - 1
        
        'right
        nP0 = P3
        nP1 = Add(Mul(1 / 2, P3), Mul(1 / 2, P2))
        nP2 = Add(Add(Mul(1 / 4, P3), Mul(1 / 2, P2)), Mul(1 / 4, P1))
        nP3 = Add(Add(Add(Mul(1 / 8, P3), Mul(3 / 8, P2)), Mul(3 / 8, P1)), Mul(1 / 8, P0))
        Draw pct, nP0, nP1, nP2, nP3, Depth - 1
    Else
        pct.Line (P0.X, P0.Y)-(P1.X, P1.Y)
        pct.Line -(P2.X, P2.Y)
        pct.Line -(P3.X, P3.Y)
    End If
End Sub 'Draw
