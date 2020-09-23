VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture Cube Settings"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMouseMove 
      Caption         =   "Discard mouse events"
      Height          =   375
      Left            =   7440
      TabIndex        =   45
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Frame Frame10 
      Caption         =   "Clock Picture"
      Height          =   2655
      Left            =   7320
      TabIndex        =   43
      Top             =   120
      Width           =   2055
      Begin ComctlLib.Slider sldClockFace 
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Max             =   3
         TickStyle       =   1
      End
      Begin VB.Image imgClockFace 
         Height          =   1335
         Left            =   360
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Cube Type"
      Height          =   1215
      Left            =   7320
      TabIndex        =   40
      Top             =   2880
      Width           =   2055
      Begin VB.OptionButton optType 
         Caption         =   "Clock"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optType 
         Caption         =   "Picture"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Opacity"
      Height          =   1095
      Left            =   4920
      TabIndex        =   36
      Top             =   5520
      Width           =   2295
      Begin ComctlLib.Slider sldOpacity 
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   10
         SmallChange     =   10
         Max             =   100
         SelStart        =   50
         TickStyle       =   1
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label Label6 
         Caption         =   "100%"
         Height          =   255
         Left            =   1680
         TabIndex        =   39
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "0%"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Mask Color"
      Height          =   855
      Left            =   4920
      TabIndex        =   33
      Top             =   4560
      Width           =   2295
      Begin VB.CheckBox chkMask 
         Caption         =   "Use Mask Color"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblMaskColor 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1680
         TabIndex        =   35
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Size"
      Height          =   1095
      Left            =   2520
      TabIndex        =   28
      Top             =   5520
      Width           =   2295
      Begin ComctlLib.Slider sldCubeSize 
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         Max             =   5
         SelStart        =   3
         TickStyle       =   1
         Value           =   3
      End
      Begin VB.Label Label4 
         Caption         =   "min"
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "max"
         Height          =   255
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Speed"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   5520
      Width           =   2295
      Begin ComctlLib.Slider sldInterval 
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   327682
         LargeChange     =   1
         Min             =   1
         Max             =   5
         SelStart        =   3
         TickStyle       =   1
         Value           =   3
      End
      Begin VB.Label Label2 
         Caption         =   "max"
         Height          =   255
         Left            =   1680
         TabIndex        =   26
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "min"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Resolution"
      Height          =   855
      Left            =   2520
      TabIndex        =   22
      Top             =   4560
      Width           =   2295
      Begin VB.ComboBox cmbResolutions 
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Text            =   "cmbResolutions"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Effects"
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   2295
      Begin VB.ComboBox cmbEffects 
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Text            =   "cmbEffects"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Faces Pictures"
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   4920
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   4920
         TabIndex        =   18
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   4920
         TabIndex        =   16
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   4920
         TabIndex        =   14
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   4920
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   21
         Top             =   2160
         Width           =   4575
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1440
         Width           =   4575
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Image imgFace 
         Height          =   1335
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   4575
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   8400
      TabIndex        =   5
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   6120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Background Picture"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   7095
      Begin VB.CommandButton cmdPicture 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   4800
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.OptionButton optBackground 
         Caption         =   "Screen"
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   360
         Width           =   1000
      End
      Begin VB.OptionButton optBackground 
         Caption         =   "Picture"
         Height          =   300
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1000
      End
      Begin VB.OptionButton optBackground 
         Caption         =   "Blank"
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1000
      End
      Begin VB.Image imgBack 
         Height          =   975
         Left            =   5520
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblPicture 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No picture"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------
' Get MyPictures Folder Path
Private Type ITEMID
    cb      As Long
    abID    As Integer
End Type

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As ITEMID) As Long
'----------
Dim cdiLoad As clsCommonDialog
Dim MyPicFold As String

Private Sub chkMask_Click()
    
    Params.Mask = chkMask.Value

End Sub

Private Sub chkMouseMove_Click()
    
    Params.MouseMove = chkMouseMove.Value

End Sub

Private Sub cmbEffects_Change()
    
    Call cmbEffects_Click

End Sub

Private Sub cmbEffects_Click()
    
    Dim idx As Integer
    
    If cmbEffects.Text = "Single" Then
        For idx = 0 To 2
            optBackground(idx).Enabled = True
        Next
    Else
        For idx = 0 To 2
            optBackground(idx).Enabled = False
        Next
    End If
    
End Sub

Private Sub Form_Load()
    
    Dim idx As Integer
    
    Set cdiLoad = New clsCommonDialog
    Call FillEffectsCombo
    Call FillResolutionsCombo
    Call RegRead
    With Params
'Faces pictures + background picture
        For idx = 0 To 6 'Mesh1.NumFaces + 1
            If Len(.FacePath(idx)) > 43 Then
                lblPicture(idx).Caption = "..." & Right(.FacePath(idx), 43)
            Else
                lblPicture(idx).Caption = .FacePath(idx)
            End If
        Next
        
        If lblPicture(6).Caption = "No picture" Then imgBack.Picture = LoadResPicture(105, vbResBitmap)
'Background options
        optBackground(.BackGroundOption).Value = True
        Call lblPicture_Click(6)
'Effects options
        cmbEffects.Text = .EffectOption
'Resolution options
        cmbResolutions.Text = .ScreenResolution
'Set Speed slider value
        sldInterval.Value = .Interval
'Set Cubesize slider value
        sldCubeSize.Value = .CubeSize
'Set Mask
        chkMask.Value = .Mask
        lblMaskColor.BackColor = .MaskColor
'Set Opacity slider value
        sldOpacity.Value = .Opacity
'Set optType
        optType(.CubeType).Value = True
'Set Clock Face Index on slider
        Select Case .ClockFaceID
            Case 101: sldClockFace.Value = 0
            Case 102: sldClockFace.Value = 1
            Case 103: sldClockFace.Value = 2
            Case 104: sldClockFace.Value = 3
        End Select
        imgClockFace.Picture = LoadResPicture(.ClockFaceID, vbResBitmap)
'Set check mouse move
        chkMouseMove.Value = .MouseMove
    End With
    
    MyPicFold = GetMyPicturesFolder
    
End Sub

Private Sub cmdPicture_Click(Index As Integer)
      
    With cdiLoad
        .Filter = "All picture files|*.bmp;*.dib;*.jpg;*.gif"
        .InitDir = MyPicFold
        .CancelError = True
        .ShowOpen
        If Len(.FileName) <> 0 Then
            Params.FacePath(Index) = .FileName
            If Len(.FileName) > 43 Then
                lblPicture(Index).Caption = "..." & Right(.FileName, 43)
            Else
                lblPicture(Index).Caption = .FileName
            End If
            Call lblPicture_Click(Index)
        End If
    End With

End Sub

Private Sub lblMaskColor_Click()
    
    cdiLoad.ShowColor
    lblMaskColor.BackColor = cdiLoad.Color
    Params.MaskColor = lblMaskColor.BackColor

End Sub

Private Sub lblPicture_Click(Index As Integer)

    If lblPicture(Index).Caption = "No picture" Then
        imgFace.Picture = LoadResPicture(105, vbResBitmap)
    Else
        If FileExists(Params.FacePath(Index)) Then
            If Index = 6 Then
                imgBack.Picture = LoadPicture(Params.FacePath(Index))
            Else
                imgFace.Picture = LoadPicture(Params.FacePath(Index))
            End If
        Else
            If Index = 6 Then
                imgBack.Picture = LoadPicture("")
            Else
                imgFace.Picture = LoadPicture("")
            End If
        End If
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set cdiLoad = Nothing

End Sub

Private Sub lblPicture_DblClick(Index As Integer)
    
    Call cmdPicture_Click(Index)

End Sub

Private Sub optBackground_Click(Index As Integer)
    
    Params.BackGroundOption = Index

End Sub

Private Sub cmdApply_Click()

    Call RegWrite
    Unload Me
    
End Sub

Private Sub cmdCancel_Click()
    
    If Params.First = False Then RegWrite
    Unload Me

End Sub

Private Sub cmbEffects_Validate(Cancel As Boolean)
    
    Params.EffectOption = cmbEffects.Text

End Sub

Private Sub cmbResolutions_Validate(Cancel As Boolean)
    
    Params.ScreenResolution = cmbResolutions.Text

End Sub

Private Sub FillEffectsCombo()
    
    With cmbEffects
        .AddItem "Single"
        .AddItem "Track"
    End With

End Sub

Private Sub FillResolutionsCombo()
    
    With cmbResolutions
        .AddItem "No change"
        .AddItem "640 x 480 x 16"
        .AddItem "640 x 480 x 32"
        .AddItem "800 x 600 x 16"
        .AddItem "800 x 600 x 32"
        .AddItem "1024 x 768 x 16"
        .AddItem "1024 x 768 x 32"
    End With

End Sub

Private Sub optType_Click(Index As Integer)
    
    Params.CubeType = Index

End Sub

Private Sub sldClockFace_Click()
    
    Select Case sldClockFace.Value
        Case 0: Params.ClockFaceID = 101
        Case 1: Params.ClockFaceID = 102
        Case 2: Params.ClockFaceID = 103
        Case 3: Params.ClockFaceID = 104
    End Select
    imgClockFace.Picture = LoadResPicture(Params.ClockFaceID, vbResBitmap)
    

End Sub

Private Sub sldCubeSize_Change()
    
    Params.CubeSize = sldCubeSize.Value

End Sub

Private Sub sldInterval_Change()
    
    Params.Interval = sldInterval.Value

End Sub

Private Sub sldOpacity_Change()
    
    Params.Opacity = sldOpacity.Value

End Sub

Private Function GetMyPicturesFolder() As String
    
    Dim RetVal As Long
    Dim tiID As ITEMID
    Dim Folder As String
    
    Folder = Space$(260)
    RetVal = SHGetSpecialFolderLocation(hWnd, &H27, tiID) 'MyPictures= &H27
    RetVal = SHGetPathFromIDList(ByVal tiID.cb, ByVal Folder)
    If RetVal Then GetMyPicturesFolder = Left$(Folder, InStr(1, Folder, Chr$(0)) - 1) & "\"

End Function

