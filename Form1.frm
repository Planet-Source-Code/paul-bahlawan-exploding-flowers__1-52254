VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7695
   FillColor       =   &H0000FF00&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' EXPLODING FLOWERS (for screen saver)
''' By Paul Bahlawan
''' March 8, 2004
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Declare Function Polygon Lib "gdi32" _
  (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Private Elem(5) As Integer
Private cx(5) As Long
Private cy(5) As Long
Private dx(5) As Long
Private dy(5) As Long
Private ca(5) As Long
Private da(5) As Long
Private l1(5) As Long
Private l2(5) As Long
Private w(5) As Long
Private bnc(5) As Single
Private bnr(5) As Single
Private r(5) As Integer
Private g(5) As Integer
Private b(5) As Integer
Private dr(5) As Integer
Private dg(5) As Integer
Private db(5) As Integer

Private Px As Long
Private Py As Long

Const pi = 3.14159265

Private Sub Form_Load()
BorderStyle = 0
Caption = "Exploding Flowers"
WindowState = 2
FillStyle = 0
FillColor = 0
BackColor = 0
ScaleMode = vbPixels
Show
Randomize
Timer1.Interval = 10
End Sub

Private Sub Timer1_Timer()
Dim cnt As Integer
Dim angle As Integer
Dim Polyset(3) As POINTAPI
Dim i As Integer
'Timer1.Enabled = False
'Do

'Create Flower
i = Int(Rnd(1) * 100)
If i < 6 Then
    If Elem(i) = 0 Then
        Elem(i) = Int(Rnd(1) * 13) + 3     'Number of petals
        cx(i) = Int(Rnd(1) * ScaleWidth)   'Center of flower
        cy(i) = Int(Rnd(1) * ScaleHeight)
        Do
            dx(i) = Int(Rnd(1) * 9) - 4    'direction
            dy(i) = Int(Rnd(1) * 9) - 4
        Loop While dx(i) = 0 And dy(i) = 0
        ca(i) = Int(Rnd(1) * 360)
        da(i) = Int(Rnd(1) * 7) - 3        'spin
        l1(i) = Int(Rnd(1) * 400) + 20     'petal length
        l2(i) = Int(Rnd(1) * (l1(i) * 0.8))
        bnc(i) = Rnd(1)                    'bounce
        bnr(i) = Rnd(1) / 20               'bounce rate
        w(i) = Int(Rnd(1) * 25) + 1        'petalwidth
        r(i) = Int(Rnd(1) * 253) + 2       'colours
        g(i) = Int(Rnd(1) * 253) + 2
        b(i) = Int(Rnd(1) * 253) + 2
        dr(i) = Int(Rnd(1) * 5) - 2
        dg(i) = Int(Rnd(1) * 5) - 2
        db(i) = Int(Rnd(1) * 5) - 2
    End If
End If


'Draw Flower(s)
For i = 0 To 5
    If Elem(i) > 0 Then
        For cnt = 0 To Elem(i) - 1
            angle = ca(i) + (360 / Elem(i) * cnt)
        
            Polyset(0).X = cx(i)
            Polyset(0).Y = cy(i)
            
            Polar l2(i) * bnc(i), angle + w(i)
            'Line (cx, cy)-(cx + Px, cy + Py), colr
            Polyset(1).X = cx(i) + Px
            Polyset(1).Y = cy(i) + Py
            
            Polar l1(i) * bnc(i), angle
            'Line (CurrentX, CurrentY)-(cx + Px, cy + Py), colr
            Polyset(2).X = cx(i) + Px
            Polyset(2).Y = cy(i) + Py
            
            Polar l2(i) * bnc(i), angle - w(i)
            'Line (CurrentX, CurrentY)-(cx + Px, cy + Py), colr
            'Line (CurrentX, CurrentY)-(cx, cy), colr
            Polyset(3).X = cx(i) + Px
            Polyset(3).Y = cy(i) + Py
            
            ForeColor = RGB(r(i), g(i), b(i))
            Polygon Me.hdc, Polyset(0), 4
        Next

'Move Flowers
        cx(i) = cx(i) + dx(i)
        cy(i) = cy(i) + dy(i)
        ca(i) = ca(i) + da(i)
        'rotate
        If ca(i) < 0 Then ca(i) = 360
        If ca(i) > 360 Then ca(i) = 0
        'bounce
        bnc(i) = bnc(i) + bnr(i)
        If bnc(i) > 1 Or bnc(i) < 0.05 Then bnr(i) = -bnr(i)
        'colours
        r(i) = r(i) + dr(i)
        If r(i) > 254 Or r(i) < 2 Then dr(i) = -dr(i)
        g(i) = g(i) + dg(i)
        If g(i) > 254 Or g(i) < 2 Then dg(i) = -dg(i)
        b(i) = b(i) + db(i)
        If b(i) > 254 Or b(i) < 2 Then db(i) = -db(i)

'Terminate Flowers
        If cx(i) < -50 Or cx(i) > ScaleWidth + 50 Then Elem(i) = 0
        If cy(i) < -50 Or cy(i) > ScaleHeight + 50 Then Elem(i) = 0
    
    End If
Next

'DoEvents
'Loop
End Sub

Private Sub Polar(ByVal r, ByVal t)
t = (t Mod 360) * pi / 180
Px = r * Cos(t)
Py = r * Sin(t)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Unload Me
End
End Sub

