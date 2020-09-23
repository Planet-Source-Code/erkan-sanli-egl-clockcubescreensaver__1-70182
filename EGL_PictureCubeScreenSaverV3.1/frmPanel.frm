VERSION 5.00
Begin VB.Form frmPanel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture Cube"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Settings"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Screensaver"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    Unload Me

End Sub

Private Sub Command1_Click()
    
    Call StartScreenSaver

End Sub

Private Sub Command2_Click()
    
    frmSetup.Show vbModal

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End

End Sub
