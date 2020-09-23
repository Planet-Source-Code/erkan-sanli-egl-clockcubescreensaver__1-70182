Attribute VB_Name = "modVector"
Option Explicit

Public Function VectorSet(X As Single, Y As Single, Z As Single) As VECTOR3

    VectorSet.X = X
    VectorSet.Y = Y
    VectorSet.Z = Z

End Function

Public Function VectorSub(V1 As VECTOR3, V2 As VECTOR3) As VECTOR3

    VectorSub.X = V1.X - V2.X
    VectorSub.Y = V1.Y - V2.Y
    VectorSub.Z = V1.Z - V2.Z
'    VectorSub.W = 1

End Function

Public Function VectorAdd(V1 As VECTOR3, V2 As VECTOR3) As VECTOR3

    VectorAdd.X = V1.X + V2.X
    VectorAdd.Y = V1.Y + V2.Y
    VectorAdd.Z = V1.Z + V2.Z
'    VectorAdd.W = 1

End Function

'Public Function VectorScale(V As VECTOR3, S As Single) As VECTOR3
'
'    VectorScale.X = V.X * S
'    VectorScale.Y = V.Y * S
'    VectorScale.Z = V.Z * S
'    VectorScale.W = 1
'
'End Function

Public Function CrossProduct(V1 As VECTOR3, V2 As VECTOR3) As VECTOR3
     
    CrossProduct.X = (V1.Y * V2.Z) - (V1.Z * V2.Y)
    CrossProduct.Y = (V1.Z * V2.X) - (V1.X * V2.Z)
    CrossProduct.Z = (V1.X * V2.Y) - (V1.Y * V2.X)
'    CrossProduct.W = 1

End Function

Public Function VectorNormalize(V As VECTOR3) As VECTOR3

    Dim VLength As Single
    
    VLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))
    If VLength = 0 Then VLength = 1
    VectorNormalize.X = V.X / VLength
    VectorNormalize.Y = V.Y / VLength
    VectorNormalize.Z = V.Z / VLength
    'VectorNormalize.W = 1

End Function

'Public Function VectorLength(V As VECTOR3) As Single
'
'    VectorLength = Sqr((V.X * V.X) + (V.Y * V.Y) + (V.Z * V.Z))
'
'End Function

Public Function DotProduct(V1 As VECTOR3, V2 As VECTOR3) As Single

    DotProduct = (V1.X * V2.X) + (V1.Y * V2.Y) + (V1.Z * V2.Z)

End Function


