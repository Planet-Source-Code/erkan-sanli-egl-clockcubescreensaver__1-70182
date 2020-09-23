Attribute VB_Name = "modVisualisation"
Option Explicit

Private Type TEXEL
    Y1      As Long
    U1      As Long
    V1      As Long
    Y2      As Long
    U2      As Long
    V2      As Long
    Used    As Boolean
End Type

Dim Texels() As TEXEL

Public Sub Render()
    
    Dim idx As Integer, iV As Integer
    Dim minX As Long, maxX As Long
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim X3 As Long, Y3 As Long
    Dim X4 As Long, Y4 As Long
    
    With Mesh1
        For idx = 0 To UBound(.FaceV)
'Get points values
            iV = .FaceV(idx).iVisible
            X1 = .Screen(.Faces(iV).a).X
            Y1 = .Screen(.Faces(iV).a).Y
            X2 = .Screen(.Faces(iV).B).X
            Y2 = .Screen(.Faces(iV).B).Y
            X3 = .Screen(.Faces(iV).C).X
            Y3 = .Screen(.Faces(iV).C).Y
            X4 = .Screen(.Faces(iV).D).X
            Y4 = .Screen(.Faces(iV).D).Y
'Redim Texels
            minX = CanLeft + 1
            maxX = CanLeft + CanWidth - 1
            ReDim Texels(minX To maxX)
'Line Interpolation
            AffineTexLine X1, Y1, X2, Y2, 0&, 0&, 225&, 0&
            AffineTexLine X2, Y2, X3, Y3, 225&, 0&, 225&, 225&
            AffineTexLine X3, Y3, X4, Y4, 225&, 225&, 0&, 225&
            AffineTexLine X4, Y4, X1, Y1, 0&, 225&, 0&, 0&
'Limitation
            If minX < 0 Then minX = 0
            If maxX > CanvasWidth Then maxX = CanvasWidth
'Fill
            For minX = minX To maxX
                FillTexLine minX, .Faces(iV).TexArray
            Next
        Next
    End With
End Sub

Private Sub AffineTexLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, U1 As Single, V1 As Single, U2 As Single, V2 As Single)
    
    Dim DeltaX      As Long
    Dim StepY       As Single
    Dim StepU       As Single
    Dim StepV       As Single
    Dim StepAllY    As Single
    Dim StepAllU    As Single
    Dim StepAllV    As Single
    
    If X1 < X2 Then
        DeltaX = X2 - X1
        StepY = Ratio(Y2 - Y1, DeltaX)
        StepU = Ratio(U2 - U1, DeltaX)
        StepV = Ratio(V2 - V1, DeltaX)
        StepAllY = CSng(Y1)
        StepAllU = CSng(U1)
        StepAllV = CSng(V1)
        For X1 = X1 To X2
            With Texels(X1)
                If .Used Then
                    If .Y1 < Y1 Then .Y1 = Y1: .U1 = U1:  .V1 = V1
                    If .Y2 > Y1 Then .Y2 = Y1: .U2 = U1:  .V2 = V1
                Else
                    .Y1 = Y1: .U1 = U1: .V1 = V1
                    .Y2 = Y1: .U2 = U1: .V2 = V1
                    .Used = True
                End If
            End With
            StepAllY = StepAllY + StepY
            StepAllU = StepAllU + StepU
            StepAllV = StepAllV + StepV
            Y1 = CLng(StepAllY)
            U1 = CLng(StepAllU)
            V1 = CLng(StepAllV)
        Next
    Else
        DeltaX = X1 - X2
        StepY = Ratio(Y1 - Y2, DeltaX)
        StepU = Ratio(U1 - U2, DeltaX)
        StepV = Ratio(V1 - V2, DeltaX)
        StepAllY = CSng(Y2)
        StepAllU = CSng(U2)
        StepAllV = CSng(V2)
        For X2 = X2 To X1
            With Texels(X2)
                If .Used Then
                    If .Y1 < Y2 Then .Y1 = Y2: .U1 = U2:  .V1 = V2
                    If .Y2 > Y2 Then .Y2 = Y2: .U2 = U2:  .V2 = V2
                Else
                    .Y1 = Y2: .U1 = U2: .V1 = V2
                    .Y2 = Y2: .U2 = U2: .V2 = V2
                    .Used = True
                End If
            End With
            StepAllY = StepAllY + StepY
            StepAllU = StepAllU + StepU
            StepAllV = StepAllV + StepV
            Y2 = CLng(StepAllY)
            U2 = CLng(StepAllU)
            V2 = CLng(StepAllV)
        Next
   End If
   
End Sub

Private Sub FillTexLine(X As Long, Data() As Long)
    
    Dim DeltaY      As Long
    Dim StepU       As Single
    Dim StepV       As Single
    Dim StepAllU    As Single
    Dim StepAllV    As Single
    Dim minY As Long, maxY As Long
On Error Resume Next
    With Texels(X)
        DeltaY = .Y1 - .Y2
        StepU = Ratio(.U1 - .U2, DeltaY)
        StepV = Ratio(.V1 - .V2, DeltaY)
        StepAllU = CSng(.U2)
        StepAllV = CSng(.V2)
        
        If .Y2 < 0 Then
            minY = 0
            StepAllU = StepAllU + (StepU * Abs(.Y2))
            StepAllV = StepAllV + (StepV * Abs(.Y2))
        Else
            minY = .Y2
        End If
        maxY = IIf(.Y1 > CanvasHeight, CanvasHeight, .Y1)
      
        If Params.Mask Then
            If Params.Opacity = 0 Then
                For minY = minY To maxY
                    If Data(.U2, .V2) <> Params.MaskColor Then CanArray(X, minY) = Bilinear(Data, StepAllU, StepAllV)
                    StepAllU = StepAllU + StepU
                    StepAllV = StepAllV + StepV
                    .U2 = CLng(StepAllU)
                    .V2 = CLng(StepAllV)
                Next
            Else
                For minY = minY To maxY
                    If Data(.U2, .V2) <> Params.MaskColor Then
                        CanArray(X, minY) = ColorBlend( _
                                                CanArray(X, minY), _
                                                Bilinear(Data, StepAllU, StepAllV), _
                                                Params.Opacity)
                    End If
                    StepAllU = StepAllU + StepU
                    StepAllV = StepAllV + StepV
                    .U2 = CLng(StepAllU)
                    .V2 = CLng(StepAllV)
                Next
            End If
        Else
            If Params.Opacity = 0 Then
                For minY = minY To maxY
                    'CanArray(X, minY) = Data(.U2, .V2)
                    CanArray(X, minY) = Bilinear(Data, StepAllU, StepAllV)
                    StepAllU = StepAllU + StepU
                    StepAllV = StepAllV + StepV
                    .U2 = CLng(StepAllU)
                    .V2 = CLng(StepAllV)
                Next
            Else
                For minY = minY To maxY
                    CanArray(X, minY) = ColorBlend( _
                                            CanArray(X, minY), _
                                            Bilinear(Data, StepAllU, StepAllV), _
                                            Params.Opacity)
                    StepAllU = StepAllU + StepU
                    StepAllV = StepAllV + StepV
                    .U2 = CLng(StepAllU)
                    .V2 = CLng(StepAllV)
                Next
            End If
        End If
    End With

End Sub

Private Function Ratio(ByVal R1 As Long, ByVal R2 As Long) As Single
    
    If (R2) Then Ratio = CSng(R1 / R2)

End Function

Private Function ColorBlend(C1 As Long, C2 As Long, Rate As Byte) As Long
    
    Dim R1 As Byte, G1 As Byte, B1 As Byte
    Dim R2 As Byte, G2 As Byte, B2 As Byte
    Dim R3 As Integer, G3 As Integer, B3 As Integer
    Dim Rate1 As Single, Rate2 As Single
    
    Rate1 = Rate / 100
    Rate2 = (100 - Rate) / 100
    
    R1 = (C1 And &HFF&)
    G1 = (C1 And &HFF00&) / &H100&
    B1 = (C1 And &HFF0000) / &H10000
    
    R2 = (C2 And &HFF&)
    G2 = (C2 And &HFF00&) / &H100&
    B2 = (C2 And &HFF0000) / &H10000

    R3 = (R1 * Rate1) + (R2 * Rate2)
    G3 = (G1 * Rate1) + (G2 * Rate2)
    B3 = (B1 * Rate1) + (B2 * Rate2)

    ColorBlend = RGB(R3, G3, B3)

End Function
