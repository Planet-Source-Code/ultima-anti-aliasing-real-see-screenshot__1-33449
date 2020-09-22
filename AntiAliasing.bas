Attribute VB_Name = "AntiAliasing"
'***********************************'
'* By: Michael Kissner (<<ULTIMA>>)*'
'* Anti-Aliasing                   *'
'* This is very simple to under-   *'
'* stand. All it does is, it takes *'
'* a point and a neighbouring one  *'
'* calculates the average and      *'
'* applies it.                     *'
'***********************************'

'I dont understand why People Copyright stuff they post on
'PSC...(Except for complex things ofcourse,like A full game+3D engine...)

'Please tell me if you know any way to make this faster :D
'(cause this is the part thats killing my Framerate in my GL)

' Two helper function.
'   -SetPixelV = same as PSet only faster
'   -GetPixel  = returns a color from a certain location (X,Y)
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long


Public Type DRGB                'My custom Color.
    R As Byte
    G As Byte
    B As Byte
End Type

Public Type Point2D             'My custom Point in 2D.
    X As Single
    Y As Single
    col As DRGB                 'Note that this uses RGB not Long
End Type                        'as color
Public Function LongToRGB(col As Long) As DRGB
With LongToRGB                  'This function Makes a long
    .R = col And 255            'into a RGB code.
    .G = (col And 65280) \ 256&
    .B = (col And 16711680) \ 65535
End With
End Function

Public Function getAvColor(Col1 As DRGB, Col2 As DRGB) As DRGB
Dim R, G, B As Double           'This function returns an
With getAvColor                 'average of 2 colors
    R = (CDbl(Col1.R) + CDbl(Col2.R)) / 2
    G = (CDbl(Col1.G) + CDbl(Col2.G)) / 2
    B = (CDbl(Col1.B) + CDbl(Col2.B)) / 2
    .R = R - (R Mod 1)          'Since the colors are in Byte
    .G = G - (G Mod 1)          'They don't have a dot
    .B = B - (B Mod 1)          'like 0.2, so 0.2 for example
End With                        'would be 0 or 100.9 = 100
End Function
Private Function AA(point As Point2D, picbox As PictureBox, intensity As Integer, DirX As Single, DirY As Single)
    Dim tempColor As DRGB       'Takes a Point and a neighbouring
                                'Point at location
                                'OriginalPoint.X + DirX
                                'OriginalPoint.Y + DirY
                                'then calculate the Average color
                                'and applies it to the new point
    tempColor = LongToRGB(GetPixel(picbox.hdc, point.X + DirX, point.Y + DirY))
    For k% = 1 To 10 Step intensity
        tempColor = getAvColor(tempColor, point.col)
    Next k%
    Call SetPixelV(picbox.hdc, point.X + DirX, point.Y + DirY, RGB(tempColor.R, tempColor.G, tempColor.B))

End Function
Public Function AntiA(point As Point2D, picbox As PictureBox, intensity As Integer, Rdif As Integer, Gdif As Integer, Bdif As Integer, DirX As Single, DirY As Single)
    Dim TRDif As Single         'This function Calls the AA
    Dim TGDif As Single         'funciton (the previewes one)
    Dim TBDif As Single         'As long as the diference between
    Dim tempColor As DRGB       'the 2 points is smaller then
    Dim tempColor1 As DRGB      'Rdif and Gdif and Bdif
    Dim tempColor2 As DRGB      'Each standing for a color (RGB)
    Dim curP As Point2D         'This function also has a certain
    curP = point                'direction (DirX and DirY).
                                'It can also go diagonal and other
    tempColor1 = curP.col       'special stuff
    tempColor2 = LongToRGB(GetPixel(picbox.hdc, curP.X + DirX, curP.Y + DirY))
                                
    If tempColor2.R >= tempColor1.R Then    'The following calculate
        TRDif = tempColor2.R - tempColor1.R 'The differences of the RGBs
    Else
        TRDif = tempColor1.R - tempColor2.R
    End If
    If TRDif < 0 Then TRDif = -TRDif
    If tempColor2.G >= tempColor1.G Then
        TGDif = tempColor2.G - tempColor1.G
    Else
        TGDif = tempColor1.G - tempColor2.G
    End If
    If TGDif < 0 Then TGDif = -TGDif
    If tempColor2.B >= tempColor1.B Then
        TBDif = tempColor2.B - tempColor1.B
    Else
        TBDif = tempColor1.B - tempColor2.B
    End If
    If TBDif < 0 Then TBDif = -TBDif
                                
    tempColor2 = LongToRGB(GetPixel(picbox.hdc, curP.X, curP.Y))
                        'While the Color diferences are bigger
                        'than the once specified, the following
                        'is executed
    While TRDif > Rdif And TGDif > Gdif And TBDif > Bdif
            
        Call AA(curP, picbox, intensity, DirX, DirY) 'calls AA Func
        tempColor1 = tempColor2 'gets the colors of the next
                                'pair of points (the previous one
                                'and
                                'Previous.X+Dirx.....
                                'This point (Prev.X + Dirx) will next
                                'round be the "previous" one.
        tempColor2 = LongToRGB(GetPixel(picbox.hdc, curP.X + DirX, curP.Y + DirY))
                        
        curP.col = tempColor2
        
        If tempColor2.R >= tempColor1.R Then    'Same as the beginning
            TRDif = tempColor2.R - tempColor1.R 'Calculates the diferences
        Else
            TRDif = tempColor1.R - tempColor2.R
        End If
        If TRDif < 0 Then TRDif = -TRDif
        If tempColor2.G >= tempColor1.G Then
            TGDif = tempColor2.G - tempColor1.G
        Else
            TGDif = tempColor1.G - tempColor2.G
        End If
        If TGDif < 0 Then TGDif = -TGDif
        If tempColor2.B >= tempColor1.B Then
            TBDif = tempColor2.B - tempColor1.B
        Else
            TBDif = tempColor1.B - tempColor2.B
        End If
        If TBDif < 0 Then TBDif = -TBDif
        
        curP.X = curP.X + DirX     'This just sets the coords
        curP.Y = curP.Y + DirY     'for the next point to be tested
    Wend
End Function

