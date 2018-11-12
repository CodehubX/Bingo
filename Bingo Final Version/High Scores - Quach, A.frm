VERSION 5.00
Begin VB.Form frmHighScores 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C000&
   Caption         =   "High Scores"
   ClientHeight    =   4845
   ClientLeft      =   5880
   ClientTop       =   1740
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "High Scores - Quach, A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   5085
   Begin VB.CommandButton cmdExitHighS 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   4575
   End
   Begin VB.PictureBox picData 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      ScaleHeight     =   3795
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Alex Quach
'Date: June 7/2016
'Purpose: To create a form that displays the high scores

Private Sub cmdExitHighS_Click()
    
    'Closes the high score form once the user is done
    Unload frmHighScores
    'Unpauses the game
    If GameStatus Then
    
        frmMain.tmrTimer.Enabled = True
        
    End If
    
End Sub

Private Sub Form_Load()
    
    'Displays the appropriate headings
    picData.Print "High scores!"; Tab(24); "Scores"
    picData.Print
    
    'Pauses the game
    frmMain.tmrTimer.Enabled = False
        
    'Reads the high score file and displays the current
    'high scores on the file
    ReadFile
    DisplayFile

End Sub
