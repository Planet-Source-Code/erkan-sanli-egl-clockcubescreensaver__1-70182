Attribute VB_Name = "modDisplaySettings"
Option Explicit

Private Const ENUM_CURRENT_SETTINGS = &HFFFF - 1
Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
   dmdevicename                   As String * 32
   dmspecversion                  As Integer
   dmdriverversion                As Integer
   dmsize                         As Integer
   dmdriverextra                  As Integer
   dmfields                       As Long
   dmorientation                  As Integer
   dmpapersize                    As Integer
   dmpaperlength                  As Integer
   dmpaperwidth                   As Integer
   dmscale                        As Integer
   dmcopies                       As Integer
   dmdefaultsource                As Integer
   dmprintquality                 As Integer
   dmcolor                        As Integer
   dmduplex                       As Integer
   dmyresolution                  As Integer
   dmttoption                     As Integer
   dmcollate                      As Integer
   dmformname                     As String * 32
   dmunusedpadding                As Integer
   dmbitsperpel                   As Long
   dmpelswidth                    As Long
   dmpelsheight                   As Long
   dmdisplayflags                 As Long
   dmdisplayfrequency             As Long
End Type

Public Type DISPLAYSETTINGS
    Width As Long
    Height As Long
    Depth As Integer
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal a As Long, ByVal B As Long, wef As DEVMODE) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByRef wef As Any, ByVal i As Long) As Long

Public CurrentDisplay           As DISPLAYSETTINGS
Public TempDisplay              As DISPLAYSETTINGS

Public Sub SetDisplay(ByVal Width As Long, ByVal Height As Long, ByVal Depth As Long)
    
    Dim DMode    As DEVMODE
    Dim RetVal As Long
    Dim iAns    As Integer

    RetVal = EnumDisplaySettings(0, 0, DMode)

    With DMode
        .dmfields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
        .dmpelswidth = Width
        .dmpelsheight = Height
        .dmbitsperpel = Depth
    End With
    RetVal = ChangeDisplaySettings(DMode, 2)
    Select Case RetVal
        Case 0
            Call ChangeDisplaySettings(DMode, 4)
        Case Else
            MsgBox "Mode not supported", vbSystemModal, "Error"
            End
    End Select
    
End Sub

Public Sub GetDisplay()

    Dim DMode As DEVMODE
    Dim RetVal As Long
    
    RetVal = EnumDisplaySettings(0&, ENUM_CURRENT_SETTINGS, DMode)
    
    If RetVal = 0 Then
        MsgBox "Error evaluating the current screen resolution!"
    Else
        CurrentDisplay.Width = Format(DMode.dmpelswidth, "@@@@")
        CurrentDisplay.Height = Format(DMode.dmpelsheight, "@@@@")
        CurrentDisplay.Depth = Format(DMode.dmbitsperpel, "@@")
    End If
    
End Sub
