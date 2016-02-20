VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   7845
   ClientTop       =   2520
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   11085
   ScaleWidth      =   11205
   Begin VB.TextBox txtInput 
      Alignment       =   2  'Center
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdRoll 
      Caption         =   "Roll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   3
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2400
      TabIndex        =   2
      Top             =   6360
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Enter a number to the right"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2520
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblRoll 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5280
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label lblRule 
      Alignment       =   2  'Center
      Caption         =   "Enter a number in the space provided above"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   2040
      TabIndex        =   1
      Top             =   8640
      Width           =   6135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Time to Play!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   0
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   0
      Picture         =   "frmGame.frx":0000
      Top             =   0
      Width           =   24000
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerate_Click()

    Randomize 'initializes the randomize function
    num = Int(Rnd * 25 + 1) 'generates a random number between 1 and 25 and stores it in a variable
    lblRule.Caption = "The number has been generated, click submit to continue" ''tells you the next steps
    cmdGenerate.Enabled = False 'stops you from generating another number
    cmdSubmit.Enabled = True 'allows you to click submit to continue

End Sub

Private Sub cmdRoll_Click()

    txtInput.Enabled = False 'Stops you from enetering a new number after rolling
    
    Randomize 'Initializes the randomize function
    roll = Int(Rnd * 6 + 1) 'generates a random number between 1 and 6 (the dice roll) and stores it in a variable
    
    lblRoll.Caption = roll 'Prints the outcome of the roll to the roll caption box
    cmdRoll.Enabled = False 'Stops you from clicking roll again
    cmdGenerate.Enabled = True 'Allows you to have the computer generate a number
    
    
End Sub

Private Sub cmdSubmit_Click()

    frmGame.Hide 'hides the game form
    frmSummary.Show 'shows the summary form

End Sub

Private Sub Form_Activate()

    txtInput.Enabled = True 'allows you to input a value
    txtInput.Text = "0" 'sets the default value to 0 in the textbox
    lblRoll.Caption = "" 'clears the roll caption box

    cmdRoll.Enabled = False 'disables the roll button
    cmdGenerate.Enabled = False 'disables the generate button
    cmdSubmit.Enabled = False 'disables the submit button
    
End Sub

Private Sub txtInput_Change()
    
    If IsNumeric(txtInput.Text) And Len(txtInput.Text) > 0 Then 'makes sure the user enters a number and checks to see if the space isnt blank
    
        If txtInput.Text > 19 Or txtInput.Text < 1 Then 'if the user inputed value is greater than 19 or less than 1 then this runs
            
            lblRule.Caption = "Enter a number between 1 and 19" 'Tells you to enter a number between 1 and 19
            cmdRoll.Enabled = False 'disables the roll button so you cant start with an invalid entry
        
        ElseIf txtInput.Text <= 19 And txtInput.Text >= 1 Then 'if the user enters a value between 1 and 19 inclusive this runs
            
            cmdRoll.Enabled = True 'enables the roll button so you can continue
            lblRule.Caption = "Click roll to roll the dice" 'Tells you next steps
            guess = txtInput.Text 'stores your inputed value in a variable
    
    End If
    
    
    Else
        
        MsgBox ("Enter a number") 'if they dont enter a number it tells them to enter a number
        cmdRoll.Enabled = False 'disables you from rolling without selecting a number first
    
    End If
    

    

        
    
End Sub
