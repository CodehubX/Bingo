VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H0080C0FF&
   Caption         =   "About"
   ClientHeight    =   3825
   ClientLeft      =   1890
   ClientTop       =   1905
   ClientWidth     =   7500
   Icon            =   "About - Quach, A.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7500
   Begin VB.CommandButton cmdExitAbout 
      BackColor       =   &H00FF00FF&
      Caption         =   "O&k"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00C000C0&
      TabIndex        =   1
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "© Alex Quach 2016, ICS 3U Riverdale Bingo Version 2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Alex Quach
'Date: May 24/2016
'Purpose: To create an about page for the user

Private Sub cmdExitAbout_Click()

    'Closes the about form when the user
    'is done
    Unload frmAbout
    
End Sub
