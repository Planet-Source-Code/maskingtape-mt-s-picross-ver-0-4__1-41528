VERSION 5.00
Begin VB.Form frm5x5Select 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Choose A Stage"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7530
   Icon            =   "frm5x5Select.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   7530
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Preview Picture"
      Height          =   1680
      Left            =   5760
      TabIndex        =   36
      ToolTipText     =   "Shows a preview pic if you already beat the stage"
      Top             =   45
      Width           =   1680
      Begin VB.PictureBox PrePic 
         Height          =   1275
         Left            =   180
         ScaleHeight     =   80
         ScaleMode       =   0  'User
         ScaleWidth      =   80
         TabIndex        =   37
         ToolTipText     =   "Shows a preview pic if you already beat the stage"
         Top             =   270
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose/Create Profile"
      ForeColor       =   &H00000000&
      Height          =   1410
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   2400
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1395
         TabIndex        =   38
         Top             =   1035
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create"
         Height          =   285
         Left            =   90
         TabIndex        =   4
         Top             =   1035
         Width           =   870
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Text            =   "Profile Name"
         Top             =   720
         Width           =   2220
      End
      Begin VB.FileListBox File1 
         Appearance      =   0  'Flat
         Height          =   420
         Left            =   90
         Pattern         =   "*.ezp"
         TabIndex        =   2
         Top             =   270
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Please Select Profile"
      Height          =   1680
      Left            =   2520
      TabIndex        =   0
      Top             =   45
      Width           =   3165
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   10
         Left            =   1710
         TabIndex        =   29
         Top             =   1350
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   1710
         TabIndex        =   28
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   8
         Left            =   1710
         TabIndex        =   27
         Top             =   810
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   7
         Left            =   1710
         TabIndex        =   12
         Top             =   540
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   6
         Left            =   1710
         TabIndex        =   11
         Top             =   270
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   10
         Top             =   1350
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   9
         Top             =   1080
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   8
         Top             =   810
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   7
         Top             =   540
         Width           =   195
      End
      Begin VB.CheckBox ckStage 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   195
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   10
         Left            =   2700
         TabIndex        =   35
         ToolTipText     =   "Time Stage Completed"
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   9
         Left            =   2700
         TabIndex        =   34
         ToolTipText     =   "Time Stage Completed"
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   8
         Left            =   2700
         TabIndex        =   33
         ToolTipText     =   "Time Stage Completed"
         Top             =   810
         Width           =   420
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   10
         Left            =   1980
         TabIndex        =   32
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   9
         Left            =   1980
         TabIndex        =   31
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   8
         Left            =   1980
         TabIndex        =   30
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   7
         Left            =   2700
         TabIndex        =   26
         ToolTipText     =   "Time Stage Completed"
         Top             =   540
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   6
         Left            =   2700
         TabIndex        =   25
         ToolTipText     =   "Time Stage Completed"
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   5
         Left            =   1080
         TabIndex        =   24
         ToolTipText     =   "Time Stage Completed"
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   4
         Left            =   1080
         TabIndex        =   23
         ToolTipText     =   "Time Stage Completed"
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   3
         Left            =   1080
         TabIndex        =   22
         ToolTipText     =   "Time Stage Completed"
         Top             =   810
         Width           =   420
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   2
         Left            =   1080
         TabIndex        =   21
         ToolTipText     =   "Time Stage Completed"
         Top             =   540
         Width           =   420
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   7
         Left            =   1980
         TabIndex        =   20
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   1980
         TabIndex        =   19
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   18
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   16
         Top             =   810
         Width           =   645
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   15
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "N/A"
         Height          =   195
         Index           =   1
         Left            =   1080
         TabIndex        =   14
         ToolTipText     =   "Time Stage Completed"
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblStage 
         Caption         =   "Stage 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   270
         Width           =   645
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   1485
      Width           =   1050
   End
End
Attribute VB_Name = "frm5x5Select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
File = App.Path & "\profiles\" & Text1.Text & ".ezp"

If StrComp(Dir(File), "", vbTextCompare) <> 0 Then
    rc = MsgBox("Profile Exists.. over write?", vbYesNo)
    If rc = vbYes Then
        Open File For Output As #1
        Write #1, "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FASLE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE", "0", "#FALSE#", "0"
        Close #1
        File1.Refresh
        Label1.Caption = File1.ListCount & " Profiles."
    Else
    End If
Else
    Open File For Output As #1
    Write #1, "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FASLE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE#", "0", "#FALSE", "0", "#FALSE#", "0"
    Close #1
    File1.Refresh
    Label1.Caption = File1.ListCount & " Profiles."
End If

Dim tmp As Integer
tmp = 1

Frame2.Caption = Text1.Text & ".ezp"

Open File For Input As #1
Input #1, EasyStagePass(1), EasyStageTime(1), EasyStagePass(2), EasyStageTime(2), EasyStagePass(3), EasyStageTime(3), EasyStagePass(4), EasyStageTime(4), EasyStagePass(5), EasyStageTime(5), EasyStagePass(6), EasyStageTime(6), EasyStagePass(7), EasyStageTime(7), EasyStagePass(8), EasyStageTime(8), EasyStagePass(9), EasyStageTime(9), EasyStagePass(10), EasyStageTime(10)
Close #1

Do While tmp < 11
    If EasyStagePass(tmp) = True Then
        ckStage(tmp).Value = Checked
        lblTime(tmp).Caption = EasyStageTime(tmp)
        lblStage(tmp).Font.Strikethrough = True
    Else
        ckStage(tmp).Value = Unchecked
        lblTime(tmp).Caption = "N/A"
        lblStage(tmp).Font.Strikethrough = False
    End If
    tmp = tmp + 1
Loop

End Sub

Private Sub Command2_Click()
Dim tmp As Integer
tmp = 1

If File1.FileName = "" Then
    MsgBox "Please select a profile to delete!"
Else
    rc = MsgBox("Profile " & File1.FileName & " will be deleted. Continue?", vbYesNo)
    If rc = vbYes Then
        Kill App.Path & "\profiles\" & File1.FileName
        File1.Refresh
        Label1.Caption = File1.ListCount & " Profiles."
        Frame2.Caption = "Please Select Profile"
        PrePic.Picture = LoadPicture(App.Path & "\stages\unknown.jpg")
        Do While tmp < 11
            ckStage(tmp).Value = Unchecked
            lblTime(tmp).Caption = "N/A"
            lblStage(tmp).Font.Strikethrough = False
            tmp = tmp + 1
        Loop
    Else
    End If
End If

End Sub

Public Sub File1_Click()
File = App.Path & "\profiles\" & File1.FileName
Dim tmp As Integer
tmp = 1

Frame2.Caption = File1.FileName
Text1.Text = Left$(File1.FileName, (Len(File1.FileName) - 4))
Call EasyRefresh
Command2.Enabled = True

End Sub

Private Sub Form_Load()

File1.Path = App.Path & "\Profiles"
Label1.Caption = File1.ListCount & " Profiles."
PrePic.Picture = LoadPicture(App.Path & "\stages\unknown.jpg")
File = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)

Done = True
frmMain.Show

End Sub

Private Sub lblStage_Click(Index As Integer)

If Frame2.Caption = "Please Select Profile" Then
    MsgBox "Please Choose a Profile."
Else
    Stage = App.Path & "\stages\" & "5x5stage" & Index & ".pcs"
    StageNum = Index
    frm5x5.Show
    Me.Hide
End If

End Sub

Private Sub lblStage_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If ckStage(Index).Value = Checked Then
    PrePic.Picture = LoadPicture(App.Path & "\stages\5x5stage" & Index & ".pcsp")
Else
    PrePic.Picture = LoadPicture(App.Path & "\stages\unknown.jpg")
End If
  
End Sub

Private Sub Text1_Change()
If InStr(1, Text1.Text, "?") Then
    MsgBox "You can't use '?' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, "\") Then
    MsgBox "You can't use '\' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, "<") Then
    MsgBox "You can't use '<' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, ">") Then
    MsgBox "You can't use '>' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, ":") Then
    MsgBox "You can't use ':' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, "|") Then
    MsgBox "You can't use '|' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, "*") Then
    MsgBox "You can't use '*' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
ElseIf InStr(1, Text1.Text, """") Then
    MsgBox "You can't use '""' in the profile name."
    Text1.Text = Left(Text1.Text, (Len(Text1.Text) - 1))
    Text1.SelStart = Len(Text1.Text)
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call Command1_Click
End Sub
