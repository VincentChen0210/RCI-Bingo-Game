VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RCI Bingo"
   ClientHeight    =   9330
   ClientLeft      =   5685
   ClientTop       =   2385
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   11655
   Begin VB.Timer tmrScore 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5880
      Top             =   7920
   End
   Begin VB.CheckBox chkAutoRun 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Demo Mode"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Frame fraCard2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Card 2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   5880
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      Begin VB.Label lblCard2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblTitle2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame fraCard1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Card 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5655
      Begin VB.Label lblTitle1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCard1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.CheckBox chkSixBySix 
      BackColor       =   &H00C0C0FF&
      Caption         =   "6v6 Version"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   7800
      Width           =   1455
   End
   Begin VB.Timer tmrDemo 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6360
      Top             =   7920
   End
   Begin VB.CommandButton cmdCheckWin 
      Caption         =   "RCI GO!!!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   1
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton cmdCall 
      Caption         =   "&Call Number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      TabIndex        =   0
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label lblWin 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   6840
      TabIndex        =   13
      Top             =   7560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblLastCalled 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   9240
      TabIndex        =   12
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Number Called:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8160
      TabIndex        =   11
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label lblCalledNum 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   3
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label lblCalledTitle 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   6360
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuHighScore 
         Caption         =   "High Score"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MAX = 90
Const INTERVAL = 100
Const COMPUTERNAME = "COMPUTER"

Dim Num(1 To 36) As Integer
Dim Mirror(1 To MAX) As Integer
Dim CallMirror(1 To MAX) As Integer
Dim WinLine1(1 To 36) As Integer
Dim WinLine2(1 To 36) As Integer
Dim CurrentCalled As Integer
Dim CallValid As Boolean
Dim VLine As Integer
Dim NumWins As Integer
Dim AutoRun As Boolean
Dim Start As Single
Dim Score As Single
Dim ScoreStart As Single
Dim AlreadyWin As Boolean
Dim ScoreTimeDiff As Single
Dim NumCalledCards As Integer
Dim HouseChance As Integer
Dim NameArray(1 To 10) As String
Dim HighScoreArray(1 To 10) As Single
Dim HouseTrue As Boolean
Dim MyFile As String
Dim LineColour1(1 To 36) As Boolean
Dim LineColour2(1 To 36) As Boolean

Public Sub CheckHighScore()
    Dim X As Integer
    Dim Y As Integer
    Dim K As Integer
    Dim Leader As Integer
    Dim Hs As Boolean
    Dim HSName As String
    Dim SwapScore As Single
    Dim SwapName As String
    
    SwapScore = 0
    SwapName = ""
    Y = 0
    
    Do While Y < 10
        Y = Y + 1
        If Score > HighScoreArray(Y) Then
            Hs = True
            If HouseTrue = False Then
                Leader = MsgBox("You got a Highscore, Do you want to put it on the leaderboards?", vbYesNo)
            Else
                Leader = vbYes
            End If
            
            X = Y
            Y = 11
        End If
    Loop
    
    If Leader = vbYes Then
        If HouseTrue = False Then
            HSName = InputBox$("Please enter your name, You got a Highscore")
        Else
            HSName = COMPUTERNAME
        End If
        
        For K = X To 10
            SwapScore = HighScoreArray(K)
            HighScoreArray(K) = Score
            Score = SwapScore
            
            SwapName = NameArray(K)
            NameArray(K) = HSName
            HSName = SwapName
        Next K
    End If
End Sub

Public Sub ShowScore()
    Dim X As Integer
    Dim Name As String
    Dim Scores As Single
    
    For X = 1 To 10
        Name = NameArray(X)
        Scores = HighScoreArray(X)
        PrintScore Name, Scores
    Next X
End Sub

Public Sub WriteScore()
    Dim X As Integer
    
    Open MyFile For Output As #1
    
    For X = 1 To 10
        If HighScoreArray(X) <> 0 Then
            Write #1, NameArray(X), HighScoreArray(X)
        End If
    Next X
    
    Close #1
End Sub

Public Sub ReadFile()
    Dim X As Integer
    Dim NumRecs As Integer
    
    X = 0
    
    Open MyFile For Input As #1
    Do Until EOF(1) Or X = 10
        X = X + 1
        Input #1, NameArray(X)
        Input #1, HighScoreArray(X)
    Loop
    Close #1
    
    NumRecs = X
    
    If NumRecs < 10 Then
        For X = (NumRecs + 1) To 10
            NameArray(X) = ""
            HighScoreArray(X) = 0
        Next X
    End If
End Sub

Public Sub DetermineScore(Score As Single)
    Const LINE = 5
    Const TIMESTART = 60
    Const HOUSESTART = 100
    Dim TimeDiff As Integer
    Dim LineScore As Integer
    Dim TimeScore As Single
    Dim CardScore As Integer
    Dim HouseScore As Single
    Dim HouseDiff As Integer
    Dim CardStart As Integer
    
    If VLine = 5 Then
        CardStart = 50
    Else
        CardStart = 72
    End If
    
    LineScore = NumWins * LINE

    TimeDiff = TIMESTART - ScoreTimeDiff
    
    If TimeDiff > 0 Then
        TimeScore = TimeDiff
    Else
        TimeScore = 0
    End If
    
    HouseDiff = HOUSESTART - ScoreTimeDiff
    
    If HouseDiff > 0 Then
        HouseScore = HouseDiff / 10
    Else
        HouseScore = 1
    End If
    
    CardScore = CardStart - 2 * NumCalledCards
    
    Score = LineScore * HouseScore + TimeScore + CardScore
    
    'MsgBox LineScore & vbCrLf & HouseScore & vbCrLf & TimeScore & vbCrLf & CardScore
    
    MsgBox "Your Score:" & Str$(Score), vbOKOnly + vbInformation
    
    tmrDemo.Enabled = False
    tmrScore.Enabled = False
End Sub

Private Sub cmdCall_Click()
    Const High = 100
    Const LOW = 1
    Dim Chance As Integer
    
    CallClick
    
    If AlreadyWin Then
        HouseChance = HouseChance + 1
        Chance = Int(Rnd * (High - LOW + 1) + LOW)
        If Chance <= HouseChance Then
            DetermineScore Score
            MsgBox "HOUSE! The CPU has detected that you are trying to rack up points! " & vbCrLf & vbCrLf & "The score you would've had:" & Str$(Score), vbCritical, "House!"
            HouseTrue = True
            CheckHighScore
            ShowScore
            WriteScore
            mnuNewGame.Enabled = True
            chkSixBySix.Enabled = True
            chkAutoRun.Enabled = True
            mnuExit.Enabled = True
        End If
'        lblWin.Caption = Chance
    End If
End Sub

Public Sub CallClick()
    Dim RowLetter As String
    Dim BoxNumber As Integer
    Dim X As Integer
    Dim CaptionValue As Integer
    
    If tmrDemo.Enabled = False Then
        tmrDemo.Enabled = True
        Start = Timer
    End If
    
    If CallValid = True Then
        CallNumber RowLetter, BoxNumber
        CurrentCalled = BoxNumber
        NumCalledCards = NumCalledCards + 1
        For X = 1 To VLine ^ 2
            If VLine = 5 Then
                If X <> 13 Then
                    CaptionValue = lblCard1(X).Caption
                    If CaptionValue = BoxNumber Then
                        WinLine1(X) = 1
                        If AutoRun = True Then
                            ClickCalledBox1 X
                        End If
                    End If
                Else
                    WinLine1(X) = 2
                End If
            Else
                If X <> 15 Then
                    CaptionValue = lblCard1(X).Caption
                    If CaptionValue = BoxNumber Then
                        WinLine1(X) = 1
                        If AutoRun = True Then
                            ClickCalledBox1 X
                        End If
                    End If
                Else
                    WinLine1(X) = 2
                End If
            End If
        Next X
        For X = 1 To VLine ^ 2
            If VLine = 5 Then
                If X <> 13 Then
                    CaptionValue = lblCard2(X).Caption
                    If CaptionValue = BoxNumber Then
                        WinLine2(X) = 1
                        If AutoRun = True Then
                            ClickCalledBox2 X
                        End If
                    End If
                Else
                    WinLine2(X) = 2
                End If
            Else
                If X <> 15 Then
                    CaptionValue = lblCard2(X).Caption
                    If CaptionValue = BoxNumber Then
                        WinLine2(X) = 1
                        If AutoRun = True Then
                            ClickCalledBox2 X
                        End If
                    End If
                Else
                    WinLine2(X) = 2
                End If
            End If
        Next X
    Else
        MsgBox "You may not call any numbers!", vbCritical, "Error"
        tmrDemo.Enabled = False
        cmdCall.Enabled = False
    End If
End Sub

Public Sub CallNumber(Row As String, Box As Integer)
    Dim Check As Boolean
    Dim K As Integer
    Dim ArraySpace As Boolean
    Dim TopRange As Integer
    
    If VLine = 6 Then
        TopRange = 90
    Else
        TopRange = 75
    End If
    
    If CallValid = True Then
        Check = False
        Do While Check = False
            Box = Int(Rnd * (TopRange) + 1)
            If CallMirror(Box) <> 1 Then
                CallMirror(Box) = 1
                Check = True
            Else
                Check = False
            End If
        Loop
        
        Select Case Box
            Case 1 To 15
                Row = "R"
            Case 16 To 30
                Row = "C"
            Case 31 To 45
                Row = "I"
            Case 46 To 60
                If VLine = 6 Then
                    Row = "N"
                Else
                    Row = "G"
                End If
            Case Else
                If VLine = 6 Then
                    If Box > 75 Then
                        Row = "O"
                    Else
                        Row = "G"
                    End If
                Else
                    Row = "O"
                End If
        End Select
        
        lblCalledNum(Box).BackColor = &HFFFF00
        lblLastCalled.Caption = Row & Box
        
        ArraySpace = False
        K = 0
        Do Until ArraySpace = True Or K = VLine * 15
            K = K + 1
            If CallMirror(K) = 0 Then
                ArraySpace = True
            Else
                ArraySpace = False
            End If
        Loop
        
        If ArraySpace = False Then
            MsgBox "You have reached the maximum amount of numbers that you may call!", vbCritical, "Error"
            tmrDemo.Enabled = False
            CallValid = False
        End If
    End If
End Sub

Public Sub LoadCalledTitle()
    Dim X As Integer
    Dim Row As String
    Dim Top As Integer
    Dim Left As Integer
    
    Top = lblCalledTitle(0).Top
    Left = lblCalledTitle(0).Left
    Row = ""
    
    For X = 1 To VLine
        Load lblCalledTitle(X)
        lblCalledTitle(X).Top = Top
        lblCalledTitle(X).Left = Left
        
        If VLine = 6 Then
            Top = Top + 480
        Else
            Top = Top + 595
        End If
        
        lblCalledTitle(X).Visible = True
        Select Case X
            Case 1
                Row = "R"
            Case 2
                Row = "C"
            Case 3
                Row = "I"
            Case 4
                If VLine = 6 Then
                    Row = "N"
                Else
                    Row = "G"
                End If
            Case Else
                If VLine = 6 Then
                    If X = 5 Then
                        Row = "G"
                    Else
                        Row = "O"
                    End If
                Else
                    Row = "O"
                End If
        End Select
        lblCalledTitle(X).Caption = Row
    Next X
End Sub

Public Sub LoadTitle1()
    Dim X As Integer
    Dim Row As String
    Dim Top As Integer
    Dim Left As Integer
    
    Top = lblTitle1(0).Top
    Left = lblTitle1(0).Left
    Row = ""
    
    For X = 1 To VLine
        Load lblTitle1(X)
        lblTitle1(X).Top = Top
        lblTitle1(X).Left = Left
        
        If VLine = 6 Then
            Left = Left + 855
        Else
            Left = Left + 1075
        End If
        
        lblTitle1(X).Visible = True
        
        Select Case X
            Case 1
                Row = "R"
            Case 2
                Row = "C"
            Case 3
                Row = "I"
            Case 4
                If VLine = 6 Then
                    Row = "N"
                Else
                    Row = "G"
                End If
            Case Else
                If VLine = 6 Then
                    If X = 5 Then
                        Row = "G"
                    Else
                        Row = "O"
                    End If
                Else
                    Row = "O"
                End If
        End Select
        lblTitle1(X).Caption = Row
    Next X
End Sub

Public Sub LoadTitle2()
    Dim X As Integer
    Dim Row As String
    Dim Top As Integer
    Dim Left As Integer
    
    Top = lblTitle2(0).Top
    Left = lblTitle2(0).Left
    Row = ""
    
    For X = 1 To VLine
        Load lblTitle2(X)
        lblTitle2(X).Top = Top
        lblTitle2(X).Left = Left
        
        If VLine = 6 Then
            Left = Left + 855
        Else
            Left = Left + 1075
        End If
        
        lblTitle2(X).Visible = True
        
        Select Case X
            Case 1
                Row = "R"
            Case 2
                Row = "C"
            Case 3
                Row = "I"
            Case 4
                If VLine = 6 Then
                    Row = "N"
                Else
                    Row = "G"
                End If
            Case Else
                If VLine = 6 Then
                    If X = 5 Then
                        Row = "G"
                    Else
                        Row = "O"
                    End If
                Else
                    Row = "O"
                End If
        End Select
        lblTitle2(X).Caption = Row
    Next X
End Sub

Public Sub LoadCard1(ByVal Limit As Integer)
    Dim Left As Integer
    Dim Top As Integer
    Dim X As Integer
    
    Top = lblCard1(0).Top
    Left = lblCard1(0).Left
    
    For X = 1 To VLine ^ 2
        Load lblCard1(X)
        lblCard1(X).Top = Top
        lblCard1(X).Left = Left
        
        If VLine = 6 Then
            Top = Top + 735
        Else
            Top = Top + 925
        End If
        
        If Top > Limit Then
            Top = 1440
            If VLine = 6 Then
                Left = Left + 855
            Else
                Left = Left + 1075
            End If
        End If
        
        lblCard1(X).Visible = True
    Next X
End Sub

Public Sub LoadCard2(ByVal Limit As Integer)
    Dim Left As Integer
    Dim Top As Integer
    Dim X As Integer
    
    Top = lblCard2(0).Top
    Left = lblCard2(0).Left
    
    For X = 1 To VLine ^ 2
        Load lblCard2(X)
        lblCard2(X).Top = Top
        lblCard2(X).Left = Left
        
        If VLine = 6 Then
            Top = Top + 735
        Else
            Top = Top + 925
        End If
        
        If Top > Limit Then
            Top = 1440
            If VLine = 6 Then
                Left = Left + 855
            Else
                Left = Left + 1075
            End If
        End If
        
        lblCard2(X).Visible = True
    Next X
End Sub

Public Sub LoadWinTable()
    Dim Left As Integer
    Dim Top As Integer
    Dim X As Integer
    
    Top = lblCalledNum(0).Top
    Left = lblCalledNum(0).Left
    
    For X = 1 To VLine * 15
        Load lblCalledNum(X)
        lblCalledNum(X).Top = Top
        lblCalledNum(X).Left = Left
        Left = Left + 480
        If Left > 7440 Then
            If VLine = 6 Then
                Top = Top + 480
            Else
                Top = Top + 595
            End If
            Left = lblCalledNum(0).Left
        End If
        lblCalledNum(X).Visible = True
        lblCalledNum(X).Caption = X
    Next X
End Sub

Private Sub cmdCheckWin_Click()
    CheckWinClick
    If AlreadyWin Then
        DetermineScore Score
'        lblWin.Caption = Score
        CheckHighScore
        ShowScore
        WriteScore
    Else
        tmrScore.Enabled = False
        tmrDemo.Enabled = False
        MsgBox "You have called RCIGO preemptively! This game is forfeited!" & vbCrLf & vbCrLf & "Please click New Game", vbExclamation + vbOKOnly, "You Lose!"
        cmdCall.Enabled = False
        cmdCheckWin.Enabled = False
    End If
    
    mnuNewGame.Enabled = True
    chkSixBySix.Enabled = True
    chkAutoRun.Enabled = True
    mnuExit.Enabled = True
End Sub

Public Sub CheckWinClick()
    Dim X As Integer
    Dim Start As Integer
    Dim Finish As Integer
    Dim Step As Integer
    Dim LINE As Integer
    Dim Free As Boolean
    
    NumWins = 0
    
    Start = 1
    Finish = VLine
    For X = 1 To VLine
        Step = 1
        LINE = VLine
        CheckWin1 Start, Finish, Step
        Start = Start + LINE
        Finish = Finish + LINE
    Next X
    
    Start = 1
    Finish = VLine * (VLine - 1) + 1
    For X = 1 To VLine
        Step = VLine
        LINE = 1
        CheckWin1 Start, Finish, Step
        Start = Start + LINE
        Finish = Finish + LINE
    Next X
    
    Start = 1
    Finish = VLine ^ 2
    Step = VLine + 1
    CheckWin1 Start, Finish, Step
    
    Start = VLine
    Finish = VLine * (VLine - 1) + 1
    Step = VLine - 1
    CheckWin1 Start, Finish, Step
    
    'Divider between Card 1 and Card 2
    
    Start = 1
    Finish = VLine
    For X = 1 To VLine
        Step = 1
        LINE = VLine
        CheckWin2 Start, Finish, Step
        Start = Start + LINE
        Finish = Finish + LINE
    Next X
    
    Start = 1
    Finish = VLine * (VLine - 1) + 1
    For X = 1 To VLine
        Step = VLine
        LINE = 1
        CheckWin2 Start, Finish, Step
        Start = Start + LINE
        Finish = Finish + LINE
    Next X
    
    Start = 1
    Finish = VLine ^ 2
    Step = VLine + 1
    CheckWin2 Start, Finish, Step
    
    Start = VLine
    Finish = VLine * (VLine - 1) + 1
    Step = VLine - 1

    CheckWin2 Start, Finish, Step
    
    If AutoRun = False And NumWins > 0 And AlreadyWin = False Then
        tmrScore.Enabled = True
        ScoreStart = Timer
        AlreadyWin = True
    End If
End Sub

Public Sub CheckWin1(ByVal Start As Integer, ByVal Finish As Integer, ByVal Step As Integer)
    Dim Win As Boolean
    Dim PlusOne As Integer
    Dim TempStart As Integer
    Dim X As Integer
    
    Win = True
    TempStart = Start
    
    Do While Win = True And Start < Finish + 1
        If WinLine1(Start) = 3 Then
            Win = True
        ElseIf WinLine1(Start) = 2 Then
            Win = True
        Else
            Win = False
        End If
        
        Start = Start + Step
    Loop
    
    If Win = True Then
        NumWins = NumWins + 1
        For X = TempStart To Finish Step Step
            LineColour1(X) = True
        Next X
    End If
End Sub

Public Sub CheckWin2(ByVal Start As Integer, ByVal Finish As Integer, ByVal Step As Integer)
    Dim Win As Boolean
    Dim PlusOne As Integer
    Dim TempStart As Integer
    Dim X As Integer
    
    Win = True
    TempStart = Start
    
    Do While Win = True And Start < Finish + 1
        If WinLine2(Start) = 3 Then
            Win = True
        ElseIf WinLine2(Start) = 2 Then
            Win = True
        Else
            Win = False
        End If
        
        Start = Start + Step
    Loop
    
    If Win = True Then
        NumWins = NumWins + 1
        For X = TempStart To Finish Step Step
            LineColour2(X) = True
        Next X
    End If
End Sub

Private Sub Form_Load()
    lblCard1(0).Visible = False
    lblCard2(0).Visible = False
    lblTitle1(0).Visible = False
    lblTitle2(0).Visible = False
    lblCalledNum(0).Visible = False
    lblCalledTitle(0).Visible = False
    cmdCall.Enabled = False
    cmdCheckWin.Enabled = False
'    cmdCall.BackColor = vbBlue
'    cmdCheckWin.BackColor = vbBlue
    MyFile = App.Path & "\HIGHSCORES.txt"
    
    Randomize
    
    ReadFile
End Sub

Private Sub lblCard1_Click(Index As Integer)
    ClickCalledBox1 Index
End Sub

Private Sub lblCard2_Click(Index As Integer)
    ClickCalledBox2 Index
End Sub

Public Sub ClickCalledBox1(ByVal Index As Integer)
    Dim Value As Integer
    
    Value = Val(lblCard1(Index).Caption)
    
    If WinLine1(Index) = 1 And Value = CurrentCalled Then
        lblCard1(Index).ForeColor = vbRed
        WinLine1(Index) = 3
    End If
End Sub

Public Sub ClickCalledBox2(ByVal Index As Integer)
    Dim Value As Integer
    
    Value = Val(lblCard2(Index).Caption)
    
    If WinLine2(Index) = 1 And Value = CurrentCalled Then
        lblCard2(Index).ForeColor = vbRed
        WinLine2(Index) = 3
    End If
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Dim ExitResponse As Integer
    
    ExitResponse = MsgBox("Are you sure you want to exit?", vbYesNo, "Exit")
    
    Select Case ExitResponse
    Case vbYes
        End
    End Select
End Sub

Private Sub mnuHighScore_Click()
    ShowScore
End Sub

Private Sub mnuNewGame_Click()
    Dim X As Integer
    Dim NumLow As Integer
    Dim NumHigh As Integer
    Dim Temp As Integer
    Dim Limit As Integer
    
    cmdCall.Enabled = True
    cmdCheckWin.Enabled = False
    
    CallValid = True
    lblLastCalled.Caption = ""
    
'    lblWin.Caption = ""
    
    AlreadyWin = False
    
    Score = 0
    
    NumCalledCards = 0
    
    HouseChance = 0
    
    For X = 1 To MAX
        Mirror(X) = 0
        CallMirror(X) = 0
    Next X
    
    For X = 1 To 36
        WinLine1(X) = 0
        WinLine2(X) = 0
    Next X
    
    Limit = 0
    
    If VLine > 1 Then
        For X = 1 To VLine ^ 2
            Unload lblCard1(X)
            Unload lblCard2(X)
        Next X
        
        For X = 1 To VLine
            Unload lblTitle1(X)
            Unload lblTitle2(X)
            Unload lblCalledTitle(X)
        Next X
        
        For X = 1 To VLine * 15
            Unload lblCalledNum(X)
        Next X
    End If
    
    If chkSixBySix.Value = False Then
        VLine = 5
    Else
        VLine = 6
    End If
    
    If chkAutoRun.Value = False Then
        AutoRun = False
        cmdCall.Enabled = True
        cmdCheckWin.Enabled = True
        tmrDemo.Enabled = False
        'tmrDemo.INTERVAL = INTERVAL
        'tmrScore.INTERVAL = INTERVAL
    Else
        AutoRun = True
        cmdCall.Enabled = False
        cmdCheckWin.Enabled = False
        tmrDemo.Enabled = True
        Start = Timer
        'tmrDemo.INTERVAL = INTERVAL
        'tmrScore.INTERVAL = INTERVAL
    End If
    
    For X = 1 To VLine ^ 2
        Num(X) = 0
    Next X
    
    Limit = 5715
    
    LoadTitle1
    LoadTitle2
    LoadCard1 Limit
    LoadCard2 Limit
    LoadCalledTitle
    LoadWinTable
    
    NumLow = 1
    NumHigh = 15
    
    For X = 1 To VLine ^ 2
        Temp = 0
        GenerateRND NumLow, NumHigh, Temp
        Num(X) = Temp
        
        If VLine = 5 Then
            If X = 13 Then
                lblCard1(X).Caption = "F"
                lblCard1(X).ForeColor = vbRed
            Else
                lblCard1(X).Caption = Num(X)
            End If
        Else
            If X = 15 Then
                lblCard1(X).Caption = "F"
                lblCard1(X).ForeColor = vbRed
            Else
                lblCard1(X).Caption = Num(X)
            End If
        End If
        
        If X Mod VLine = 0 Then
            NumLow = NumLow + 15
            NumHigh = NumHigh + 15
        End If
    Next X
    
    For X = 1 To MAX
        Mirror(X) = 0
    Next X
    
    NumLow = 1
    NumHigh = 15
    
    For X = 1 To VLine ^ 2
        Temp = 0
        GenerateRND NumLow, NumHigh, Temp
        Num(X) = Temp
        
        If VLine = 5 Then
            If X = 13 Then
                lblCard2(X).Caption = "F"
                lblCard2(X).ForeColor = vbRed
            Else
                lblCard2(X).Caption = Num(X)
            End If
        Else
            If X = 15 Then
                lblCard2(X).Caption = "F"
                lblCard2(X).ForeColor = vbRed
            Else
                lblCard2(X).Caption = Num(X)
            End If
        End If
        
        If X Mod VLine = 0 Then
            NumLow = NumLow + 15
            NumHigh = NumHigh + 15
        End If
    Next X
    
    mnuNewGame.Enabled = False
    chkSixBySix.Enabled = False
    chkAutoRun.Enabled = False
    mnuExit.Enabled = False
End Sub

Public Sub GenerateRND(ByVal LOW As Integer, ByVal High As Integer, Num As Integer)
    Dim Check As Boolean
    Dim K As Integer
    
    Check = False
    Do While Check = False
        Num = Int(Rnd * (High - LOW + 1) + LOW)
        If Mirror(Num) <> 1 Then
            Mirror(Num) = 1
            Check = True
        Else
            Check = False
        End If
    Loop
End Sub

Public Function TODOLIST()
'additional forms
End Function

Private Sub tmrDemo_Timer()
    Dim Current As Single
    Dim Diff As Single
    Dim X As Integer
    
    Current = Timer
    
    Diff = Current - Start
    
    If AutoRun = True Then
        CallClick
    End If
    
    CheckWinClick
    
    If AutoRun = True Then
        If NumWins > 0 Then
'            lblWin.Caption = NumWins & " lines"
            For X = 1 To 36
                If LineColour1(X) = True Then
                    lblCard1(X).BackColor = &HFFFF00
                End If
                If LineColour2(X) = True Then
                    lblCard2(X).BackColor = &HFFFF00
                End If
            Next X
            
            tmrDemo.Enabled = False
            mnuNewGame.Enabled = True
            chkSixBySix.Enabled = True
            chkAutoRun.Enabled = True
            mnuExit.Enabled = True
        Else
            For X = 1 To 36
                LineColour1(X) = False
                LineColour2(X) = False
            Next X
'            lblWin.Caption = "No lines"
        End If
        'lblTimer.Caption = Format$(Diff, "0")
    End If
End Sub

Private Sub tmrScore_Timer()
    Dim Current As Single
    
    Current = Timer
    
    ScoreTimeDiff = Current - ScoreStart
End Sub
