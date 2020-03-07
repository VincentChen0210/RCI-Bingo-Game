VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About This Program"
   ClientHeight    =   3525
   ClientLeft      =   7575
   ClientTop       =   3000
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   7545
   Begin VB.Frame Frame1 
      Caption         =   "Tips"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3960
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
      Begin VB.PictureBox picOutput 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   3255
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Created By: Vincent Chen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "Created For: ICS3U Culminating"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3000
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "Game: RCI BINGO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   $"About.frx":0000
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
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

Private Sub picOutput_Paint()
    picOutput.Print "Ctrl + N = New Game"
    picOutput.Print
    picOutput.Print "You can only start a new game if your current"
    picOutput.Print " game has ended."
    picOutput.Print ""
End Sub
