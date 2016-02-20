VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   12540
   ClientLeft      =   7320
   ClientTop       =   1665
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   12540
   ScaleWidth      =   13260
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4980
      TabIndex        =   3
      Top             =   6960
      Width           =   3015
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "Rules"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5040
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4980
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Guess the number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   10575
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -120
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

    End 'ends the program

End Sub

Private Sub cmdPlay_Click()

    If rules = False Then 'checks to see if the user has read the rules
    
        MsgBox ("Please read the rules before starting the game") 'if they havent it tells you to read the rules
        frmMain.Hide 'hides the main form
        frmRules.Show 'shows the rules form
    
    ElseIf rules = True Then 'if the user has read the rules this runs

        frmMain.Hide 'hides the main form
        frmGame.Show 'starts the game

    End If



End Sub

Private Sub cmdRules_Click()

    frmMain.Hide 'hides the main form
    frmRules.Show 'shows the rules form

End Sub
