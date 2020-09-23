Attribute VB_Name = "modPublics"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Public Type VECTOR3
    X                   As Single
    Y                   As Single
    Z                   As Single
End Type

Public Type FACE
    a                   As Integer
    B                   As Integer
    C                   As Integer
    D                   As Integer
    Normal              As VECTOR3
    NormalT             As VECTOR3
    TexArray()          As Long
End Type

Public Type ORDER
    ZValue              As Single
    iVisible            As Integer
End Type

Public Type POINTAPI
    X                   As Long
    Y                   As Long
End Type

Public Type MESH
    NumVector           As Integer
    NumFaces            As Integer
    Vertices()          As VECTOR3
    VerticesT()         As VECTOR3
    Screen()            As POINTAPI
    Faces()             As FACE
    FaceV()             As ORDER
    Scale               As VECTOR3
    Rotation            As VECTOR3
    Translation         As VECTOR3
    HalfDiagonalLenght  As Single
End Type

Public Type COLORRGB
    R               As Integer
    G               As Integer
    B               As Integer
End Type

Public Mesh1            As MESH
Public CubeLen          As Single

Public CanvasDIB        As New clsDIB
Public BackDIB          As New clsDIB
Public TexDIB(5)        As New clsDIB

Public CanArray()       As Long
Public BackArray()      As Long

Public CanvasWidth      As Long
Public CanvasHeight     As Long
Public OriginX          As Long
Public OriginY          As Long
Public CanLeft          As Long
Public CanTop           As Long
Public CanWidth         As Long
Public CanHeight        As Long

Public Sub CreateCube()
    
    Dim idx As Integer
    
    With Mesh1
        .NumVector = 7
        .NumFaces = 5
        ReDim .Vertices(.NumVector)
        ReDim .VerticesT(.NumVector)
        ReDim .Screen(.NumVector)
        ReDim .Faces(.NumFaces)

        .Vertices(0).X = -0.5:   .Vertices(0).Y = -0.5:   .Vertices(0).Z = 0.5
        .Vertices(1).X = 0.5:    .Vertices(1).Y = -0.5:   .Vertices(1).Z = 0.5
        .Vertices(2).X = 0.5:    .Vertices(2).Y = 0.5:    .Vertices(2).Z = 0.5
        .Vertices(3).X = -0.5:   .Vertices(3).Y = 0.5:    .Vertices(3).Z = 0.5
        .Vertices(4).X = -0.5:   .Vertices(4).Y = -0.5:   .Vertices(4).Z = -0.5
        .Vertices(5).X = 0.5:    .Vertices(5).Y = -0.5:   .Vertices(5).Z = -0.5
        .Vertices(6).X = 0.5:    .Vertices(6).Y = 0.5:    .Vertices(6).Z = -0.5
        .Vertices(7).X = -0.5:   .Vertices(7).Y = 0.5:    .Vertices(7).Z = -0.5
        
        .Faces(0).a = 0:        .Faces(0).B = 1:        .Faces(0).C = 2:          .Faces(0).D = 3
        .Faces(1).a = 1:        .Faces(1).B = 5:        .Faces(1).C = 6:          .Faces(1).D = 2
        .Faces(2).a = 5:        .Faces(2).B = 4:        .Faces(2).C = 7:          .Faces(2).D = 6
        .Faces(3).a = 4:        .Faces(3).B = 0:        .Faces(3).C = 3:          .Faces(3).D = 7
        .Faces(4).a = 3:        .Faces(4).B = 2:        .Faces(4).C = 6:          .Faces(4).D = 7
        .Faces(5).a = 1:        .Faces(5).B = 0:        .Faces(5).C = 4:          .Faces(5).D = 5
        
        For idx = 0 To .NumFaces
            .Faces(idx).Normal = VectorNormalize(CrossProduct( _
                                 VectorSub(.Vertices(.Faces(idx).C), .Vertices(.Faces(idx).B)), _
                                 VectorSub(.Vertices(.Faces(idx).a), .Vertices(.Faces(idx).B))))
        Next
        .Scale = VectorSet(CubeLen, CubeLen, CubeLen)
        .HalfDiagonalLenght = (Fix(Sqr(.Scale.X ^ 2 + .Scale.Y ^ 2 + .Scale.Y ^ 2)) / 2) - (CubeLen / 3)
    End With
    
End Sub

Public Sub CalculateLimits()
    
    Dim idx         As Integer
    Dim MinVector   As POINTAPI
    Dim MaxVector   As POINTAPI

    With Mesh1
        MinVector = .Screen(0)
        MaxVector = .Screen(0)
        For idx = 1 To .NumVector
            If .Screen(idx).X < MinVector.X Then MinVector.X = .Screen(idx).X
            If .Screen(idx).Y < MinVector.Y Then MinVector.Y = .Screen(idx).Y
            If .Screen(idx).X > MaxVector.X Then MaxVector.X = .Screen(idx).X
            If .Screen(idx).Y > MaxVector.Y Then MaxVector.Y = .Screen(idx).Y
        Next
        CanLeft = (MinVector.X) - 1
        CanTop = (MinVector.Y) - 1
        CanWidth = (MaxVector.X - MinVector.X) + 3
        CanHeight = (MaxVector.Y - MinVector.Y) + 3
    End With

End Sub

Public Function RotatePoint(Angle As Integer, LLenght As Integer) As POINTAPI

    Const CCenter = 113
    Dim Radian As Single
    
    Radian = Angle * sPIDiv180
    RotatePoint.X = Cos(Radian) * LLenght + CCenter
    RotatePoint.Y = Sin(Radian) * LLenght + CCenter

End Function

