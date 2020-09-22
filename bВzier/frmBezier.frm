VERSION 5.00
Begin VB.Form frmBezier 
   Caption         =   "Bezier - C.Dutoit - Feb 2001 - http://www.dutoitc.fr.st"
   ClientHeight    =   6675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDraw 
      Caption         =   "&Draw !"
      Height          =   495
      Left            =   7920
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox pct 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H000040C0&
      ForeColor       =   &H000040C0&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   397
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   1
      Top             =   240
      Width           =   7695
   End
   Begin VB.HScrollBar scrStep 
      Height          =   255
      LargeChange     =   4
      Left            =   120
      Max             =   10
      Min             =   1
      TabIndex        =   0
      Top             =   6360
      Value           =   1
      Width           =   7095
   End
   Begin VB.Label Label2 
      Caption         =   "Segments"
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblNbSeg 
      Alignment       =   1  'Right Justify
      Caption         =   "n"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   6360
      Width           =   495
   End
End
Attribute VB_Name = "frmBezier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name   : frmBezier
'Author : C.Dutoit
'Mail   : dutoitc@hotmail.com
'Web    : http://www.dutoitc.fr.st or http://dutoitc.tsx.org
'ICQ    : UIN = 9657082
'MSN    : dutoitc@hotmail.com
Option Explicit

Dim P(0 To 3) As TVector    'Base Vertices (Points in french)
Dim NbSeg As Byte           'Nb of segments=depth for Bezier calculation

Dim DockedPoint As Integer  'Docked point by mouse

'draw all
Private Sub cmdDraw_Click()
    Redraw
End Sub 'cmdDraw_Click


'Init
Private Sub Form_Load()
    'Init main vertices
    P(0) = Vec(pct.ScaleWidth / 10, pct.ScaleHeight * 9 / 10)
    P(1) = Vec(pct.ScaleWidth / 3, pct.ScaleHeight / 4)
    P(2) = Vec(pct.ScaleWidth * 2 / 3, pct.ScaleHeight / 4)
    P(3) = Vec(pct.ScaleWidth * 9 / 10, pct.ScaleHeight * 9 / 10)
    
    'Nb segments
    NbSeg = 4
    scrStep.Value = NbSeg
    
    'mouse docked point
    DockedPoint = -1
    
    'redraw
    Show
    Redraw
End Sub 'Form_Load


'Dock point with mouse
Private Sub pct_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Dist As Long, TmpDist As Long
    Dim i As Byte
    
    'Search nearest point
    DockedPoint = 0
    Dist = Sqr((X - P(0).X) ^ 2 + (Y - P(0).Y) ^ 2)
    For i = 1 To 3
        TmpDist = Sqr((X - P(i).X) ^ 2 + (Y - P(i).Y) ^ 2)
        If TmpDist < Dist Then
            Dist = TmpDist
            DockedPoint = i
        End If
    Next i
    pct_MouseMove Button, Shift, X, Y
End Sub 'pct_MouseDown


'Move ctrl points
Private Sub pct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If DockedPoint > -1 Then
        P(DockedPoint).X = X
        P(DockedPoint).Y = Y
        Redraw
    End If
End Sub 'pct_MouseMove


'Remove docking
Private Sub pct_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DockedPoint = -1
End Sub 'pct_mouseup


'Change Caption and redraw
Private Sub scrStep_Change()
    NbSeg = scrStep.Value                   'Change actual value
    lblNbSeg.Caption = 3 * 2 ^ NbSeg        'Change caption for user
    Redraw                                  'Redraw the Bezier curve
End Sub 'scrStep_Change


'draw all
Sub Redraw()
    Dim i As Byte
    
    pct.Cls
    
    'Draw ctrl points
    For i = 0 To 3
        pct.Circle (P(i).X, P(i).Y), 4, RGB(0, 255, 255)
    Next i
    
    'draw the Bezier curve
    Draw pct, P(0), P(1), P(2), P(3), NbSeg
End Sub 'Redraw
