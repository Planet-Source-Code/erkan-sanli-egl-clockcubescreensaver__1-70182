VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    RGBResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
 
Public Enum CdlgExt_Flags
' Open & Save Dialog
    cdlOFNAllowMultiselect = &H200
    cdlOFNCreatePrompt = &H2000
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400
    cdlOFNFileMustExist = &H1000
    cdlOFNHelpButton = &H10
    cdlOFNHideReadOnly = &H4
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000
    cdlOFNNoValidate = &H100
    cdlOFNOverwritePrompt = &H2
    cdlOFNPathMustExist = &H800
    cdlOFNReadOnly = &H1
    cdlOFNShareAware = &H4000
'Color Dialog
    cdlCCANYCOLOR = &H100
    cdlCCFullOpen = &H2
    cdlCCHelpButton = &H8
    cdlCCPreventFullOpen = &H4
    cdlCCRGBInit = &H1
End Enum

Const MAX_PATH = 260                    'General
Private RetValue        As Long         'General
Private OFN             As OPENFILENAME ' Open & Save Dialog

Private mFileName As String
Private mFileTitle As String
Private mhOwner As Long
Private mDialogTitle As String
Private mFilter As String
Private mInitDir As String
Private mSelDir As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mHelpFile As String
Private mHelpKey As Long
Private mRGBResult As Long
Private mDialogPrompt As String
Private mFlags As CdlgExt_Flags
Private mCancelError As Boolean
Private mhIcon As Long
Private mAppName As String
Private i As Integer

Public Property Let CancelError(ByVal vData As Boolean)
   mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
  CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property

Public Property Get flags() As CdlgExt_Flags
    flags = mFlags
End Property

Public Property Let flags(ByVal New_Flags As CdlgExt_Flags)
    mFlags = New_Flags
End Property

Public Property Get DialogTitle() As String
   DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
   mDialogTitle = sTitle
End Property

Public Property Get DialogPrompt() As String
    DialogPrompt = mDialogPrompt
End Property

Public Property Let DialogPrompt(ByVal New_Prompt As String)
    mDialogPrompt = New_Prompt
End Property

Public Property Get AppName() As String
    AppName = mAppName
End Property

Public Property Let AppName(ByVal New_AppName As String)
    mAppName = New_AppName
End Property

Public Property Let hIcon(ByVal vData As Long)
    mhIcon = vData
End Property

Public Property Get hIcon() As Long
   hIcon = mhIcon
End Property

' Open , Save
Public Property Get DefaultExt() As String
   DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(sDefExt As String)
   mDefaultExt = sDefExt
End Property

Public Property Get FileName() As String
   FileName = mFileName
End Property

Public Property Let FileName(sFileName As String)
   mFileName = sFileName
End Property

Public Property Get FileTitle() As String
   FileTitle = mFileTitle
End Property

Public Property Let FileTitle(sTitle As String)
   mFileTitle = sTitle
End Property

Public Property Get Filter() As String
   Filter = mFilter
End Property

Public Property Let Filter(sFilter As String)
   mFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
   FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(lIndex As Long)
    mFilterIndex = lIndex
End Property

Public Property Get InitDir() As String
   InitDir = mInitDir
End Property

Public Property Let InitDir(sDir As String)
    mInitDir = sDir
End Property

Public Property Get SelDir() As String
   SelDir = mSelDir
End Property

Public Property Let SelDir(sDir As String)
    mSelDir = sDir
End Property

'Color Dialog
Public Property Get Color() As Long
   Color = mRGBResult
End Property

Public Property Let Color(lValue As Long)
   mRGBResult = lValue
End Property

Public Sub ShowOpen()
Dim iDelim As Integer
  InitOFN
  RetValue = GetOpenFileName(OFN)
    If RetValue > 0 Then
       iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
       If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
       iDelim = InStr(OFN.lpstrFile, vbNullChar)
       If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
    Else
       If mCancelError Then mFileName = "" 'Err.Raise 0
    End If
End Sub

Private Sub InitOFN()
  Dim sTemp As String, i As Integer
  Dim uFlag As Long
  uFlag = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H10 Or &H100 Or &H200 Or &H400 Or &H800 Or &H1000 Or &H2000 Or &H4000 Or &H8000 Or &H40000 Or &H80000 Or &H100000 Or &H200000)
  With OFN
       .lStructSize = Len(OFN)
       .hWndOwner = mhOwner
       .flags = uFlag
       .lpstrDefExt = mDefaultExt
       sTemp = mInitDir
       If sTemp = "" Then sTemp = App.Path
       .lpstrInitialDir = sTemp
       sTemp = mFileName
       .lpstrFile = sTemp & String$(255 - Len(sTemp), 0)
       .nMaxFile = 255
       .lpstrFileTitle = String$(255, 0)
       .nMaxFileTitle = 255
        sTemp = mFilter
            For i = 1 To Len(sTemp)
                If Mid(sTemp, i, 1) = "|" Then
                   Mid(sTemp, i, 1) = vbNullChar
                End If
            Next
        sTemp = sTemp & String$(2, 0)
        .lpstrFilter = sTemp
        .nFilterIndex = mFilterIndex
        .lpstrTitle = mDialogTitle
        .hInstance = App.hInstance
 End With
End Sub

Public Sub ShowColor()
Dim cc As CHOOSECOLOR
Dim a As Long
    cc.lStructSize = Len(cc)
    cc.hWndOwner = mhOwner
    cc.hInstance = App.hInstance
    cc.flags = mFlags And (&H1 Or &H2 Or &H4 Or &H8 Or &H100)
    cc.lpCustColors = GetSetting(App.EXEName, "Settings", "CustomColors", String$(16 * 4, 0)) 'load custom colors
    a = ChooseColorAPI(cc)
        If (a) Then
            mRGBResult = cc.RGBResult
            SaveSetting App.EXEName, "Settings", "CustomColors", cc.lpCustColors 'save custom colors
        Else
            If mCancelError Then Err.Raise (RetValue)
        End If
End Sub

