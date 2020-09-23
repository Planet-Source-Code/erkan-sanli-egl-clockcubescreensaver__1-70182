Attribute VB_Name = "modFilter"
Option Explicit

Public Function Bilinear(Data() As Long, U As Single, V As Single) As Long

    Dim IntX As Integer, IntY As Integer
    Dim FracX As Single, FracY As Single
    Dim SubFracX As Single, SubFracY As Single
    Dim C0 As COLORRGB
    Dim C1 As COLORRGB
    Dim C2 As COLORRGB
    Dim C3 As COLORRGB
    Dim C4 As COLORRGB
    
    On Error Resume Next
    
    IntX = Fix(U)
    IntY = Fix(V)
    FracX = U - IntX
    FracY = V - IntY
    SubFracX = 1 - FracX
    SubFracY = 1 - FracY
        
    C1 = ColorLongToRGB(Data(IntX, IntY))
    If IntX < 225 Then
        C2 = ColorLongToRGB(Data(IntX + 1, IntY))
    Else
        C2 = C1
    End If
    
    If IntY < 225 Then
        C3 = ColorLongToRGB(Data(IntX, IntY + 1))
    Else
        C3 = C2
    End If
    
    If IntX < 225 And IntY < 225 Then
        C4 = ColorLongToRGB(Data(IntX + 1, IntY + 1))
    Else
        C4 = C3
    End If
    
    C0.R = (FracX * (FracY * C4.R + SubFracY * C2.R)) + (SubFracX * (FracY * C3.R + SubFracY * C1.R))
    C0.G = (FracX * (FracY * C4.G + SubFracY * C2.G)) + (SubFracX * (FracY * C3.G + SubFracY * C1.G))
    C0.B = (FracX * (FracY * C4.B + SubFracY * C2.B)) + (SubFracX * (FracY * C3.B + SubFracY * C1.B))
    
    Bilinear = ColorRGBToLong(C0)

End Function

Public Function ColorLongToRGB(lColor As Long) As COLORRGB

    ColorLongToRGB.R = (lColor And &HFF&)
    ColorLongToRGB.G = (lColor And &HFF00&) / &H100&
    ColorLongToRGB.B = (lColor And &HFF0000) / &H10000

End Function

Public Function ColorRGBToLong(C1 As COLORRGB) As Long

    ColorRGBToLong = RGB(C1.R, C1.G, C1.B)

End Function

