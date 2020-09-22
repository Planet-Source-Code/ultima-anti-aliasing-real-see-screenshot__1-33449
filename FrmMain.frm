VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Area 
      BackColor       =   &H00000000&
      Height          =   3495
      Left            =   0
      ScaleHeight     =   229
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3960
         Top             =   2640
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This part of the Program is badly written, i think :),
'the module is better to read (and more important ;))
'because this is just a demo of what you can do with AA (Anti
'Aliasing). There are tons of things you can do with AA.
'(I am not sure if what i did is real AA but i hope so :))
'
'
'Just for your info... the sample is badly coded because i didnt work
'on it long (i am sleepy thats y :)) but i worked hard on the AA
'(because i wasnt sleepy then ;D )
'
'
'If you ever should use this in your programs, you DON'T need to ask
'me... (its PSC right? its obvious that you will use it :D)
'You DON'T even have to give me credit... i am honored enough if you
'actualy use it (would be more of a wonder)


Dim points(300) As Point2D  'I am quite sleepy so i wont
Dim Temp1 As Single         'comment the sample as much
Dim Xx As Single            'as i did the Module...
Dim Yy As Single            'Sorry...
Dim PNum As Integer
Dim isMin As Boolean

Private Sub Form_Load()
DoEvents
Me.Show
End Sub
Private Function MakeP(X As Single, Y As Single, R As Byte, G As Byte, B As Byte) As Point2D
With MakeP
    .X = X
    .Y = Y
    .col.R = R              'Helper function
    .col.G = G
    .col.B = B
End With
End Function
Private Sub Alias(inte As Integer)
    Call AntiA(points(inte), Area, 10, 1, 1, 1, -1, 0)
    Call AntiA(points(inte), Area, 10, 1, 1, 1, 1, 0) 'I was lazy
    Call AntiA(points(inte), Area, 10, 1, 1, 1, 0, 1) '(its 3:40 am)
    Call AntiA(points(inte), Area, 10, 1, 1, 1, 0, -1)
End Sub
Private Sub ReDraw()
PNum = 0
Xx = 0                           'redraws the "Bolt"
Yy = 100
While PNum <= UBound(points)
    Randomize
    Temp1 = Rnd() * 20
    Temp1 = (Temp1 - (Temp1 Mod 1)) - 10
    If Temp1 < 0 Then isMin = True Else isMin = False
    
    For k% = 0 To Sqr(Temp1 * Temp1)
        
        PNum = PNum + 1
        If PNum <= UBound(points) Then
            If isMin = False Then
                Yy = Yy + 1
                points(k%) = MakeP(Xx, Yy, 100, 100, 200)
                Call SetPixelV(Area.hdc, points(k%).X, points(k%).Y, RGB(points(k%).col.R, points(k%).col.G, points(k%).col.B))
                Alias (k%)
            Else
                Yy = Yy - 1
                points(k%) = MakeP(Xx, Yy, 100, 100, 200)
                Call SetPixelV(Area.hdc, points(k%).X, points(k%).Y, RGB(points(k%).col.R, points(k%).col.G, points(k%).col.B))
                Alias (k%)
            End If
        End If
        Xx = Xx + 1
    Next k%
Wend

End Sub

Private Sub Timer1_Timer()
Area.Cls
DoEvents            'Draws 2 bolts at once...
ReDraw
ReDraw
End Sub
' GOOOOOOOOOOOOOOOOD NIGHT!!!!!!!!!!!!!!! zzz...z....zz...z
