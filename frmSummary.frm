VERSION 5.00
Begin VB.Form frmSummary 
   Caption         =   "Form1"
   ClientHeight    =   12360
   ClientLeft      =   6795
   ClientTop       =   1485
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   ScaleHeight     =   12360
   ScaleWidth      =   12120
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
      Height          =   735
      Left            =   6480
      TabIndex        =   12
      Top             =   10320
      Width           =   1935
   End
   Begin VB.CommandButton cmdMain 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   11
      Top             =   10320
      Width           =   1935
   End
   Begin VB.Label lblComputer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   10
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Computer Generated #:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   9
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label lblOutcome 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   8
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Outcome:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   7
      Top             =   9000
      Width           =   2055
   End
   Begin VB.Label lblSum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Sum:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   5
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblGuess 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   4
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Guessed Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label lblDice 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Dice Roll:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   1
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label lblSummary 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   0
      Picture         =   "frmSummary.frx":0000
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "frmSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Integer 'represents the sum of the rolled number and the inputed number

Private Sub cmdExit_Click()

    End 'ends the program

End Sub

Private Sub cmdMain_Click()

    frmMain.Show 'shows the main form
    frmSummary.Hide 'hides the summary form
    rules = False 'makes rules false so the next person to play reads them too

End Sub

Private Sub Form_Activate()

    lblDice.Caption = roll 'outputs what you rolled to the dice caption box
    lblGuess.Caption = guess 'outputs what you guessed to the guess caption box
    
    total = roll + guess 'calculates the sum of the
    lblSum.Caption = total 'outputs the sum of the roll and guess to the sum caption box
    lblComputer.Caption = num

    If num Mod total = 0 Then 'this takes the computer generated number, divides by the sum and checks to see if the remainder is 0. If the remainder is 0 its either a factor or the actual number
    
        lblOutcome.Caption = "WIN" 'tells you that you won
    
    Else 'if the remainder is anything other than 0 then it runs this

        lblOutcome.Caption = "You lost" 'tells you that you lost

    End If


End Sub
