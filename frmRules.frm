VERSION 5.00
Begin VB.Form frmRules 
   Caption         =   "Form1"
   ClientHeight    =   11265
   ClientLeft      =   7215
   ClientTop       =   2310
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   ScaleHeight     =   11265
   ScaleWidth      =   13725
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7680
      TabIndex        =   3
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4320
      TabIndex        =   2
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmRules.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Label lblRules 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   0
      Top             =   960
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   0
      Picture         =   "frmRules.frx":01C5
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdMain_Click()

    frmRules.Hide 'hides the rules form
    frmMain.Show 'shows the main menu form

End Sub

Private Sub cmdPlay_Click()

    frmRules.Hide 'hides the rules form
    frmGame.Show 'shows the game form

End Sub

Private Sub Form_Load()

    rules = True 'sets a boolean value to true to represent that the user has read the rules

End Sub
