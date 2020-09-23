Attribute VB_Name = "modMatrix"
Option Explicit
    
Public Const sPIDiv180  As Single = 0.0174533

Public Type MATRIX
    rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
    rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
    rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
End Type

Public Function MatrixMultVector(M As MATRIX, V As VECTOR3) As VECTOR3

    MatrixMultVector.X = M.rc11 * V.X + M.rc12 * V.Y + M.rc13 * V.Z + M.rc14
    MatrixMultVector.Y = M.rc21 * V.X + M.rc22 * V.Y + M.rc23 * V.Z + M.rc24
    MatrixMultVector.Z = M.rc31 * V.X + M.rc32 * V.Y + M.rc33 * V.Z + M.rc34

End Function

Public Function MatrixWorld() As MATRIX
    
    Dim CosX As Single
    Dim SinX As Single
    Dim CosY As Single
    Dim SinY As Single
    Dim CosZ As Single
    Dim SinZ As Single
    
    With Mesh1
        With .Rotation
            CosX = Cos(.X * sPIDiv180)
            SinX = Sin(.X * sPIDiv180)
            CosY = Cos(.Y * sPIDiv180)
            SinY = Sin(.Y * sPIDiv180)
            CosZ = Cos(.Z * sPIDiv180)
            SinZ = Sin(.Z * sPIDiv180)
        End With
        MatrixWorld.rc11 = .Scale.X * CosY * CosZ
        MatrixWorld.rc12 = .Scale.Y * (SinX * SinY * CosZ + CosX * -SinZ)
        MatrixWorld.rc13 = .Scale.Z * (CosX * SinY * CosZ + SinX * SinZ)
        MatrixWorld.rc14 = .Translation.X
        MatrixWorld.rc21 = .Scale.X * CosY * SinZ
        MatrixWorld.rc22 = .Scale.Y * (SinX * SinY * SinZ + CosX * CosZ)
        MatrixWorld.rc23 = .Scale.Z * (CosX * SinY * SinZ + -SinX * CosZ)
        MatrixWorld.rc24 = .Translation.Y
        MatrixWorld.rc31 = .Scale.X * -SinY
        MatrixWorld.rc32 = .Scale.Y * SinX * CosY
        MatrixWorld.rc33 = .Scale.Z * CosX * CosY
        MatrixWorld.rc34 = .Translation.Z
    End With
    
End Function
