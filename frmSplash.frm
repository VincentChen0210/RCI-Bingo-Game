VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3750
   ClientLeft      =   6645
   ClientTop       =   5505
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrStart 
      Interval        =   1000
      Left            =   5400
      Top             =   3240
   End
   Begin VB.Image Image1 
      Height          =   3735
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim X As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblProductName.Caption = App.Title
    
    X = 2
End Sub

Private Sub tmrStart_Timer()
    X = X - 1
    
    If X = 0 Then
        frmMain.Show
        frmSplash.Hide
    End If
End Sub
