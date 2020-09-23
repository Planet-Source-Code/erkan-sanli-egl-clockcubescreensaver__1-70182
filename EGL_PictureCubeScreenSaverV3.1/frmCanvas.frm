VERSION 5.00
Begin VB.Form frmCanvas 
   BorderStyle     =   0  'None
   Caption         =   "EGL_PictureCube"
   ClientHeight    =   8670
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   ClipControls    =   0   'False
   Icon            =   "frmCanvas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   578
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4320
      Top             =   3960
   End
   Begin VB.PictureBox picClock 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3390
      Left            =   240
      ScaleHeight     =   226
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   226
      TabIndex        =   2
      Top             =   4200
      Visible         =   0   'False
      Width           =   3390
   End
   Begin VB.PictureBox picLoad 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   4200
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.PictureBox PicNoPic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   120
      ScaleHeight     =   253
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   4440
   End
End
Attribute VB_Name = "frmCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tX           As Single
Dim tY           As Single
Dim CometColorOp As Long

Private Declare Function PaintDesktop Lib "user32" (ByVal hDC As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long


Private Sub Form_Load()
    
    Dim ScrRes() As String
    Dim TmpRes As String
    Me.Hide
    Timer1.Enabled = False
    Call RegRead
    Call GetDisplay
    If Thumbnail = False Then
        TmpRes = Left(Params.ScreenResolution, Len(Params.ScreenResolution) - 1)
        If TmpRes <> "No change" Then
            With TempDisplay    'Get current display setting (width,height and bits)
                ScrRes = Split(Params.ScreenResolution, " x ")
                .Width = CLng(ScrRes(0))
                .Height = CLng(ScrRes(1))
                .Depth = CInt(ScrRes(2))
                Call SetDisplay(.Width, .Height, .Depth)
            End With
        Else
            TempDisplay = CurrentDisplay
        End If
        Me.WindowState = vbMaximized
    Else
        TempDisplay.Width = 152
        TempDisplay.Height = 112
    End If
    Me.ScaleMode = vbPixels
    Me.WindowState = vbMaximized
    tX = 2
    tY = 1
    CanvasWidth = TempDisplay.Width
    CanvasHeight = TempDisplay.Height
    OriginX = CanvasWidth / 2
    OriginY = CanvasHeight / 2
    CanLeft = 0
    CanTop = 0
    CanWidth = CanvasWidth
    CanHeight = CanvasHeight
    Select Case Params.CubeSize
        Case 1: CubeLen = TempDisplay.Height / 10
        Case 2: CubeLen = TempDisplay.Height / 8
        Case 3: CubeLen = TempDisplay.Height / 6
        Case 4: CubeLen = TempDisplay.Height / 4
        Case 5: CubeLen = TempDisplay.Height / 2
    End Select
    Call CreateCube
    Select Case Params.Interval
        Case 1: Timer1.Interval = 200
        Case 2: Timer1.Interval = 150
        Case 3: Timer1.Interval = 100
        Case 4: Timer1.Interval = 50
        Case 5: Timer1.Interval = 1
    End Select
    
    If Thumbnail = True Then Timer1.Interval = 400
    
    Call LoadPics
    Timer1.Enabled = True
    Me.Show

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Params.MouseMove Then Exit Sub
    If Thumbnail = False Then Unload Me

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    
    Dim X1 As Single
    Static X2 As Single
    
    If Params.MouseMove Then Exit Sub
    
    X1 = X
    If X2 = 0 Then X2 = X1
    If X1 <> X2 And Thumbnail = False Then Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   
    If Thumbnail = False Then Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim idx As Integer
    
    If Params.ScreenResolution <> "No change" Then
        Call SetDisplay(CurrentDisplay.Width, CurrentDisplay.Height, CurrentDisplay.Depth)
    End If
    Call CanvasDIB.Delete(CanArray): Set CanvasDIB = Nothing
    Call BackDIB.Delete(BackArray): Set BackDIB = Nothing
    For idx = 0 To 5
        Call TexDIB(idx).Delete(Mesh1.Faces(idx).TexArray)
        Set TexDIB(idx) = Nothing
    Next
    Call EndScreenSaver
    End

End Sub

Private Sub LoadPics()
        
    Dim idx As Integer
'Canvas
    Call CanvasDIB.Create(CanvasWidth + 1, CanvasHeight + 1, CanArray)

'Background: Picture or Screen
    If Params.BackGroundOption = 1 Then
            If FileExists(Params.FacePath(6)) Then
                Call BackDIB.CreateFromFile(Params.FacePath(6), CanvasWidth + 1, CanvasHeight + 1, BackArray)
            Else
                Call BackDIB.CreateArrayFromPictureBox(PicNoPic, CanvasWidth + 1, CanvasHeight + 1, BackArray)
            End If
    ElseIf Params.BackGroundOption = 2 Then
            frmCanvas.WindowState = vbMaximized
            picLoad.Move 0, 0, CurrentDisplay.Width, CurrentDisplay.Height
            PaintDesktop (picLoad.hDC)
            Call BackDIB.Create(CanvasWidth, CanvasHeight, BackArray)
            SetStretchBltMode BackDIB.hDC, vbPaletteModeNone
            StretchBlt BackDIB.hDC, 0, 0, CanvasWidth, CanvasHeight, picLoad.hDC, 0, 0, CurrentDisplay.Width, CurrentDisplay.Height, vbSrcCopy
    End If
    
'Faces
    If Params.CubeType = 1 Then
        picClock.Picture = LoadResPicture(Params.ClockFaceID, vbResBitmap)
        For idx = 0 To 5
            Call TexDIB(idx).Delete(Mesh1.Faces(idx).TexArray)
            Call TexDIB(idx).CreateArrayFromPictureBox(picClock, 226, 226, Mesh1.Faces(idx).TexArray)
        Next
        Timer2.Enabled = True
    Else
        Timer2.Enabled = False
        For idx = 0 To 5
            If FileExists(Params.FacePath(idx)) Then
                Call TexDIB(idx).CreateFromFile(Params.FacePath(idx), 226, 226, Mesh1.Faces(idx).TexArray)
            Else
                Call TexDIB(idx).CreateArrayFromPictureBox(PicNoPic, 226, 226, Mesh1.Faces(idx).TexArray)
            End If
        Next
    End If
    
End Sub

Private Sub Timer1_Timer()

    Dim idx As Long
    Dim matWorld As MATRIX

    DoEvents
    With Mesh1
'Action Control
        If CanLeft + CanWidth > CanvasWidth Then
            tX = -1
        ElseIf CanLeft <= 0 Then
            tX = 2
        End If
        If CanTop + CanHeight > CanvasHeight Then
            tY = -3
        ElseIf CanTop <= 0 Then
            tY = 4
        End If

'Update position
        .Rotation.X = .Rotation.X - 1: .Rotation.X = .Rotation.X Mod 360
        .Rotation.Y = .Rotation.Y - 2: .Rotation.Y = .Rotation.Y Mod 360
        .Rotation.Z = .Rotation.Z + 1: .Rotation.Z = .Rotation.Z Mod 360
        .Translation.X = .Translation.X + tX
        .Translation.Y = .Translation.Y + tY
        matWorld = MatrixWorld
        For idx = 0 To .NumVector
            .VerticesT(idx) = MatrixMultVector(matWorld, .Vertices(idx))
            .Screen(idx).X = .VerticesT(idx).X + OriginX
            .Screen(idx).Y = .VerticesT(idx).Y + OriginY
        Next idx

'Faces normal
        For idx = 0 To .NumFaces
            .Faces(idx).NormalT = MatrixMultVector(matWorld, .Faces(idx).Normal)
        Next

'Rendering
        Select Case Params.EffectOption
            Case "Single"
                If Params.BackGroundOption = 0 Then
                    Call CanvasDIB.Clear(CanArray)
                Else
                    BitBlt CanvasDIB.hDC, CanLeft, CanTop, CanWidth, CanHeight, BackDIB.hDC, CanLeft, CanTop, vbSrcCopy
                End If
            Case "Track"
        End Select
        Call CalculateLimits
        If VisibleFaces > -1 Then Call Render
        StretchBlt Me.hDC, 0, 0, CanvasWidth, CanvasHeight, CanvasDIB.hDC, 1, 1, CanvasWidth - 1, CanvasHeight - 1, vbSrcCopy
    End With
End Sub


Private Sub Timer2_Timer()

    Dim idx As Integer
    Dim p1 As POINTAPI
    Dim p2 As POINTAPI
    Dim Angle As Integer
    Dim col As Long
    
    Select Case Params.ClockFaceID
        Case 101: col = RGB(155, 155, 155)
        Case 102: col = RGB(15, 15, 15)
        Case 103: col = RGB(155, 25, 155)
        Case 104: col = RGB(250, 250, 250)
    End Select
    
    picClock.Picture = LoadResPicture(Params.ClockFaceID, vbResBitmap)
'hour
    Angle = (30 * (Hour(Now) + (Minute(Now) / 60))) - 90
    p1 = RotatePoint(Angle, -10)
    p2 = RotatePoint(Angle, 40)
    picClock.DrawWidth = 9
    picClock.Line (p1.X, p1.Y)-(p2.X, p2.Y), col
'second
    Angle = 6 * Minute(Now) - 90
    p1 = RotatePoint(Angle, -10)
    p2 = RotatePoint(Angle, 60)
    picClock.DrawWidth = 7
    picClock.Line (p1.X, p1.Y)-(p2.X, p2.Y), col
'minute
    Angle = 6 * Second(Now) - 90
    p1 = RotatePoint(Angle, -10)
    p2 = RotatePoint(Angle, 70)
    picClock.DrawWidth = 5
    picClock.Line (p1.X, p1.Y)-(p2.X, p2.Y), RGB(95, 35, 35)
    
    picClock.Picture = picClock.Image
    For idx = 0 To 5
        Call TexDIB(idx).Delete(Mesh1.Faces(idx).TexArray)
        Call TexDIB(idx).CreateArrayFromPictureBox(picClock, 226, 226, Mesh1.Faces(idx).TexArray)
    Next
End Sub


