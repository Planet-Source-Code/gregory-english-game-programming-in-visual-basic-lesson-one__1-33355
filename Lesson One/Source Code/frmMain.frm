VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GP in VB - Lesson one"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit TTT"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "New Game"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblP2Score 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   14
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblP1Score 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblP2 
      Alignment       =   2  'Center
      Caption         =   "Player Two"
      Height          =   255
      Left            =   1920
      TabIndex        =   12
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblP1 
      Alignment       =   2  'Center
      Caption         =   "Player One"
      Height          =   255
      Left            =   1920
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.Line lnBoard 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   3
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line lnBoard 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   2
      X1              =   600
      X2              =   600
      Y1              =   0
      Y2              =   2160
   End
   Begin VB.Line lnBoard 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   1
      X1              =   0
      X2              =   1800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line lnBoard 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Index           =   0
      X1              =   0
      X2              =   1800
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   8
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   7
      Left            =   720
      TabIndex        =   7
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblSquare 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*************************
'*Project: prjLessonOne  *
'*Desc: Tic-Tac-Toe      *
'*Created: 4/02/02       *
'*************************

'//Simple Declarations
Dim bytPlayerTurn As Byte '0 = No Game, 1 = Player One, 2 = Player Two
Dim intP1Score As Integer
Dim intP2Score As Integer

Private Sub cmdNewGame_Click()
'//Start the game!
StartGame
End Sub

Private Sub Form_Load()
'//Initialize the variables
bytPlayerTurn = 0 'No Game started
intP1Score = 0
intP2Score = 0
End Sub

Private Sub StartGame()
'//This sub starts our tic-tac-toe game
Dim X As Long
'//clear board just in case of previous game
For X = 0 To 8
    lblSquare(X).Caption = ""
Next X
'//player turn = 1
bytPlayerTurn = 1
End Sub

Private Sub lblSquare_Click(Index As Integer)
'//Before applying an X or an O check if there is a game
If bytPlayerTurn = 0 Then
    '//no game
    MsgBox "Sorry, no game is in progress at this time!"
    Exit Sub '//get out of this sub
End If

'//Whos turn is it?
If bytPlayerTurn = 1 Then
    '//Player One
    '//check if spot is already taken
    If lblSquare(Index).Caption <> "" Then
        MsgBox "Sorry, this spot is already taken!"
        Exit Sub
    End If
    
    '//Add the X to the spot
    lblSquare(Index).Caption = "X"
        
    '//Check to see if they win the game
    If CheckWin("X") = True Then
        '//Player One wins!
        MsgBox "Player One is the winner!"
        intP1Score = intP1Score + 1
        lblP1Score.Caption = intP1Score
        bytPlayerTurn = 0 '//game over so no game running
        Exit Sub
    End If
    '//if doesnt win, go to player 2
    bytPlayerTurn = 2
ElseIf bytPlayerTurn = 2 Then
    '//Player Two
    '//check if spot is already taken
    If lblSquare(Index).Caption <> "" Then
        MsgBox "Sorry, this spot is already taken!"
        Exit Sub
    End If
    
    '//Add the O to the spot
    lblSquare(Index).Caption = "O"
        
    '//Check to see if they win the game
    If CheckWin("O") = True Then
        '//Player Two wins!
        MsgBox "Player Two is the winner!"
        intP2Score = intP2Score + 1
        lblP2Score.Caption = intP2Score
        bytPlayerTurn = 0 '//game over so no game running
        Exit Sub
    End If
    '//if doesnt win, go to player 1
    bytPlayerTurn = 1
End If
End Sub

Private Function CheckWin(strLetter As String) As Boolean
'//Remember we have to check every spot....
'//strLetter holds what letter we are looking for X or O
'//Top Horizontal
If lblSquare(0).Caption = strLetter And lblSquare(1).Caption = strLetter And lblSquare(2).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//Mid Horizontal
If lblSquare(3).Caption = strLetter And lblSquare(4).Caption = strLetter And lblSquare(5).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//Bottom horizontal
If lblSquare(6).Caption = strLetter And lblSquare(7).Caption = strLetter And lblSquare(8).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//Left vertical
If lblSquare(0).Caption = strLetter And lblSquare(3).Caption = strLetter And lblSquare(6).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//mid vertical
If lblSquare(1).Caption = strLetter And lblSquare(4).Caption = strLetter And lblSquare(7).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//right vertical
If lblSquare(2).Caption = strLetter And lblSquare(5).Caption = strLetter And lblSquare(8).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//Diagnol top left to bottom right
If lblSquare(0).Caption = strLetter And lblSquare(4).Caption = strLetter And lblSquare(8).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//Diagnol top right to bottom left
If lblSquare(2).Caption = strLetter And lblSquare(4).Caption = strLetter And lblSquare(6).Caption = strLetter Then
    '//Win!!
    CheckWin = True
    Exit Function
End If

'//If we get this far, no win
CheckWin = False
End Function
