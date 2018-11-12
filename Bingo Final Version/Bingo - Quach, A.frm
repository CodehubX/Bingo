VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFF00&
   Caption         =   "Bingo"
   ClientHeight    =   8715
   ClientLeft      =   2820
   ClientTop       =   1245
   ClientWidth     =   11355
   Icon            =   "Bingo - Quach, A.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   11355
   Begin VB.Timer tmrTimer 
      Interval        =   500
      Left            =   4320
      Top             =   3600
   End
   Begin VB.Frame fraBoard2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Board 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4215
      Left            =   7320
      TabIndex        =   2
      Top             =   120
      Width           =   3855
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   25
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   24
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   23
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   22
         Left            =   840
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   21
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   20
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   19
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   18
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   17
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   16
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   15
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   14
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   13
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   12
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   11
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   10
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   9
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   8
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   7
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   6
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   4
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   3
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover3 
         Height          =   735
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   25
         Left            =   3000
         TabIndex        =   129
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   24
         Left            =   2280
         TabIndex        =   128
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   1560
         TabIndex        =   127
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   840
         TabIndex        =   126
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   120
         TabIndex        =   125
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   3000
         TabIndex        =   124
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   2280
         TabIndex        =   123
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   1560
         TabIndex        =   122
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   840
         TabIndex        =   121
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   120
         TabIndex        =   120
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   3000
         TabIndex        =   119
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   2280
         TabIndex        =   118
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   1560
         TabIndex        =   117
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   840
         TabIndex        =   116
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   120
         TabIndex        =   115
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   3000
         TabIndex        =   114
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   2280
         TabIndex        =   113
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   1560
         TabIndex        =   112
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   840
         TabIndex        =   111
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   110
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3000
         TabIndex        =   109
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2280
         TabIndex        =   108
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1560
         TabIndex        =   107
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   840
         TabIndex        =   106
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   105
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraBoard1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Board 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3855
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   25
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   24
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   23
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   22
         Left            =   840
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   21
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   20
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   19
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   18
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   17
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   16
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   15
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   14
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   13
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   12
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   11
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   10
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   9
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   8
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   7
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   6
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   4
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   3
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover 
         Height          =   735
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   25
         Left            =   3000
         TabIndex        =   29
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   24
         Left            =   2280
         TabIndex        =   28
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   1560
         TabIndex        =   27
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   840
         TabIndex        =   26
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   3000
         TabIndex        =   24
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   2280
         TabIndex        =   23
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   1560
         TabIndex        =   22
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   840
         TabIndex        =   21
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   2280
         TabIndex        =   18
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   1560
         TabIndex        =   17
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   840
         TabIndex        =   16
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   3000
         TabIndex        =   14
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   2280
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   1560
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3000
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblBingoNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraWinningNum 
      BackColor       =   &H00FFFF00&
      Caption         =   "Winning numbers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   11055
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   75
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   74
         Left            =   9480
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   73
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   72
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   71
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   70
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   69
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   68
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   67
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   66
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   65
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   64
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   63
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   62
         Left            =   840
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   61
         Left            =   120
         Stretch         =   -1  'True
         Top             =   3240
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   60
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   59
         Left            =   9480
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   58
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   57
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   56
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   55
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   54
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   53
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   52
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   51
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   50
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   49
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   48
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   47
         Left            =   840
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   46
         Left            =   120
         Stretch         =   -1  'True
         Top             =   2520
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   45
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   44
         Left            =   9480
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   43
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   42
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   41
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   40
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   39
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   38
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   37
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   36
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   35
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   34
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   33
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   32
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   31
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1800
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   30
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   29
         Left            =   9480
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   28
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   27
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   26
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   25
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   24
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   23
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   22
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   21
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   20
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   19
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   18
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   17
         Left            =   840
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   16
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   15
         Left            =   10200
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   14
         Left            =   9480
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   13
         Left            =   8760
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   12
         Left            =   8040
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   11
         Left            =   7320
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   10
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   9
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   8
         Left            =   5160
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   7
         Left            =   4440
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   6
         Left            =   3720
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   5
         Left            =   3000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   4
         Left            =   2280
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   3
         Left            =   1560
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   2
         Left            =   840
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image imgCover2 
         Height          =   735
         Index           =   1
         Left            =   120
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   75
         Left            =   10200
         TabIndex        =   104
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   74
         Left            =   9480
         TabIndex        =   103
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   73
         Left            =   8760
         TabIndex        =   102
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   72
         Left            =   8040
         TabIndex        =   101
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   71
         Left            =   7320
         TabIndex        =   100
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   70
         Left            =   6600
         TabIndex        =   99
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   69
         Left            =   5880
         TabIndex        =   98
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   68
         Left            =   5160
         TabIndex        =   97
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   67
         Left            =   4440
         TabIndex        =   96
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   66
         Left            =   3720
         TabIndex        =   95
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   65
         Left            =   3000
         TabIndex        =   94
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   64
         Left            =   2280
         TabIndex        =   93
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   63
         Left            =   1560
         TabIndex        =   92
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   62
         Left            =   840
         TabIndex        =   91
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   61
         Left            =   120
         TabIndex        =   90
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   60
         Left            =   10200
         TabIndex        =   89
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   59
         Left            =   9480
         TabIndex        =   88
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   58
         Left            =   8760
         TabIndex        =   87
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   57
         Left            =   8040
         TabIndex        =   86
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   56
         Left            =   7320
         TabIndex        =   85
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   55
         Left            =   6600
         TabIndex        =   84
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   54
         Left            =   5880
         TabIndex        =   83
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   53
         Left            =   5160
         TabIndex        =   82
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   52
         Left            =   4440
         TabIndex        =   81
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   51
         Left            =   3720
         TabIndex        =   80
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   50
         Left            =   3000
         TabIndex        =   79
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   49
         Left            =   2280
         TabIndex        =   78
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   48
         Left            =   1560
         TabIndex        =   77
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   47
         Left            =   840
         TabIndex        =   76
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   46
         Left            =   120
         TabIndex        =   75
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   45
         Left            =   10200
         TabIndex        =   74
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   44
         Left            =   9480
         TabIndex        =   73
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   43
         Left            =   8760
         TabIndex        =   72
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   42
         Left            =   8040
         TabIndex        =   71
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   41
         Left            =   7320
         TabIndex        =   70
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   40
         Left            =   6600
         TabIndex        =   69
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   39
         Left            =   5880
         TabIndex        =   68
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   38
         Left            =   5160
         TabIndex        =   67
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   37
         Left            =   4440
         TabIndex        =   66
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   36
         Left            =   3720
         TabIndex        =   65
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   35
         Left            =   3000
         TabIndex        =   64
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   34
         Left            =   2280
         TabIndex        =   63
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   33
         Left            =   1560
         TabIndex        =   62
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   32
         Left            =   840
         TabIndex        =   61
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   31
         Left            =   120
         TabIndex        =   60
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   30
         Left            =   10200
         TabIndex        =   59
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   29
         Left            =   9480
         TabIndex        =   58
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   28
         Left            =   8760
         TabIndex        =   57
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   27
         Left            =   8040
         TabIndex        =   56
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   26
         Left            =   7320
         TabIndex        =   55
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   25
         Left            =   6600
         TabIndex        =   54
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   24
         Left            =   5880
         TabIndex        =   53
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   23
         Left            =   5160
         TabIndex        =   52
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   22
         Left            =   4440
         TabIndex        =   51
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   21
         Left            =   3720
         TabIndex        =   50
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   20
         Left            =   3000
         TabIndex        =   49
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   19
         Left            =   2280
         TabIndex        =   48
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   18
         Left            =   1560
         TabIndex        =   47
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   17
         Left            =   840
         TabIndex        =   46
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   16
         Left            =   120
         TabIndex        =   45
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   15
         Left            =   10200
         TabIndex        =   44
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   14
         Left            =   9480
         TabIndex        =   43
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   13
         Left            =   8760
         TabIndex        =   42
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   8040
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   7320
         TabIndex        =   40
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   6600
         TabIndex        =   39
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   5880
         TabIndex        =   38
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   5160
         TabIndex        =   37
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   4440
         TabIndex        =   36
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   3720
         TabIndex        =   35
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3000
         TabIndex        =   34
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   2280
         TabIndex        =   33
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   1560
         TabIndex        =   32
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   840
         TabIndex        =   31
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblWinNum 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image imgPHold2 
      Height          =   735
      Left            =   5880
      Picture         =   "Bingo - Quach, A.frx":030A
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgFree 
      Height          =   735
      Left            =   5040
      Picture         =   "Bingo - Quach, A.frx":0614
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgPHold1 
      Height          =   735
      Left            =   4200
      Picture         =   "Bingo - Quach, A.frx":091E
      Stretch         =   -1  'True
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblDisplay2 
      Caption         =   "Click auto draw if you would like to auto draw the bingo numbers!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4200
      TabIndex        =   4
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblDisplay1 
      Caption         =   "Click reveal game cards to start the game!"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      Height          =   4095
      Left            =   4080
      Top             =   240
      Width           =   3135
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuReveal 
         Caption         =   "&Reveal game cards"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAutoDraw 
         Caption         =   "A&uto draw"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "&High Scores"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Programmer: Alex Quach
'Date: June 7/2016
'Purpose: To create the game of bingo
Option Explicit
Const WMax = 75
Const BMax = 25
Dim BingoNum(1 To BMax) As Integer
Dim BingoNum2(1 To BMax) As Integer
Dim WinningNum(1 To WMax) As Integer
Dim WinCount(1 To BMax) As Integer
Dim WinCount2(1 To BMax) As Integer
Dim WinTotal As Integer
Dim Clicks As Integer
Private Sub Form_Load()
    
    Dim K As Integer
    Dim X As Integer
    
    'Intializes the form
    Randomize
    
    mnuNewGame.Enabled = False
    mnuAutoDraw.Enabled = False
    tmrTimer.Enabled = False
    Clicks = 0
    WinTotal = 0
    GameStatus = False
    
    'Creates the file if no high score file exists
    If Dir$(FNAME) = "" Then
    
        Createfile
    
    End If
    
    'Reads the contents of the high score file
    ReadFile
    
    'Intializes the variables
    For K = 1 To BMax
    
        BingoNum(K) = 0
        WinCount(K) = 0
        WinCount2(K) = 0
        BingoNum2(K) = 0
        
    Next K
    
    For X = 1 To WMax
    
        WinningNum(X) = X
        
    Next X
    
    'Starts a new game
    NewGame BingoNum(), WinningNum(), WMax, BMax, WinCount(), WinCount2(), BingoNum2()
    
End Sub
Private Sub imgCover2_Click(Index As Integer)
    
    Dim DrawType As String
    Dim TempName As String
    Dim TempScore As Integer
    
    'Checks to see which gamemode the user is playing
    If tmrTimer.Enabled = True Then
        
        DrawType = " automatic"
        
    ElseIf tmrTimer.Enabled = False Then
    
        DrawType = " manual"
    
    End If
    
    'If the user clicks on a game card on the winning board,
    'disable the automatic draw
    mnuAutoDraw.Enabled = False
    
    Clicks = Clicks + 1
    imgCover2(Index).Visible = False
    imgCover2(Index).Enabled = False
    'Checks to see if the number clicked matches a number on
    'the bingo board
    
    MatchCheck WinCount(), WinCount2(), BMax, BingoNum(), BingoNum2(), WinningNum(), Index
    
    'Makes sure theres a minimum of 4 clicks
    'before checking if theres a win
    If Clicks >= 4 Then
    
        'Checks if the user has won
        WinCheck WinCount(), WinTotal, WinCount2()
        If WinTotal = 5 Then
            
            'Displays the appropriate messages
            mnuNewGame.Enabled = True
            MsgBox "Congratulations, you won in" & Str$(Clicks) & DrawType & " draws!", vbInformation, "Winner!"
            tmrTimer.Enabled = False
            GameStatus = False
            WinTotal = 0
            TempScore = Clicks
            'Checks to see if there is a new highscore
            If TempScore < HighestScore(HIGHSCOREMAX) Then
            
                'Allows the user to input their name if they
                'achieved a high score
                TempName = Trim$(InputBox$("You have gotten a high score!", "High score!"))
                If TempName = "" Then
                
                    TempName = "Anonymous"
                    
                End If
                
                'Displays the high score screen to show the user their new high score
                ChangeHighScore TempScore, TempName
                frmHighScores.Show vbModal
                
            End If
            fraWinningNum.Enabled = False
            
        End If
        
    End If
    
End Sub
Private Sub mnuAbout_Click()
    
    'Directs the user to the about form
    frmAbout.Show vbModal
    
End Sub
Private Sub mnuAutoDraw_Click()
    
    'Enables the time and sets the game status to true
    tmrTimer.Enabled = True
    GameStatus = True

End Sub
Private Sub mnuExit_Click()
    
    Dim Reply As Integer
    Dim RType As Integer
    
    RType = vbYesNo + vbInformation
    
    'Asks the user if they would like to exit the program
    Reply = MsgBox("Are you sure you want to exit?", RType, "Exit")
    
    'Ends the program if the user replies with yes
    If Reply = vbYes Then
    
        End
    
    End If
    
End Sub
Private Sub mnuHighScore_Click()

    'Displays the high scores to theuser
    frmHighScores.Show vbModal

End Sub
Private Sub mnuNewGame_Click()
    
    'Starts a new game
    NewGame BingoNum(), WinningNum(), WMax, BMax, WinCount(), WinCount2(), BingoNum2()
    'Intializes the win checking variables and re-enables the menu items
    ReEnable WinTotal, Clicks
  
    
End Sub
Private Sub mnuReveal_Click()
    
    'Reveals the game boards and enables the winning numbers board
    Dim K As Integer
    
    mnuReveal.Enabled = False
    mnuNewGame.Enabled = False
    mnuAutoDraw.Enabled = True
    For K = 1 To BMax
    
        'Reveals the bingo boards
        If K <> 13 Then
            'Loads the free image if it is the middle of the board
            imgCover(K).Picture = LoadPicture("")
            imgCover3(K).Picture = LoadPicture("")
        Else
            imgCover(K).Picture = imgFree.Picture
            imgCover3(K).Picture = imgFree.Picture
        End If
    
    Next K
    fraWinningNum.Enabled = True
    
End Sub
Private Sub tmrTimer_Timer()

    'Draws the winning numbers automatically
    Call imgCover2_Click(Clicks)
    fraWinningNum.Enabled = False

End Sub
