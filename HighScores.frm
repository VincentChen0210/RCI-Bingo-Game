VERSION 5.00
Begin VB.Form frmHighScores 
   BackColor       =   &H00C0C0FF&
   Caption         =   "RCI-GO: High Scores"
   ClientHeight    =   2895
   ClientLeft      =   7905
   ClientTop       =   6075
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   5160
   Begin VB.PictureBox picData 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2355
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   " Name                                   Score"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
