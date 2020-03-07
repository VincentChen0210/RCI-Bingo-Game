Attribute VB_Name = "Module1"
Option Explicit

Public Sub PrintScore(ByVal Name As String, ByVal Highscores As Single)
    frmHighScores.Show 0
    If Highscores <> 0 Then
        frmHighScores.picData.Print Name; Tab(40); Str$(Highscores)
    End If
End Sub
