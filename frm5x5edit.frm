VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm5x5edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "5x5 Editor"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   2625
   Icon            =   "frm5x5edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   2625
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRow5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   38
      Text            =   "0"
      Top             =   2385
      Width           =   600
   End
   Begin VB.TextBox txtRow4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   37
      Text            =   "0"
      Top             =   2070
      Width           =   600
   End
   Begin VB.TextBox txtRow3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   36
      Text            =   "0"
      Top             =   1755
      Width           =   600
   End
   Begin VB.TextBox txtRow2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   35
      Text            =   "0"
      Top             =   1440
      Width           =   600
   End
   Begin VB.TextBox txtRow1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   34
      Text            =   "0"
      Top             =   1125
      Width           =   600
   End
   Begin VB.TextBox txtCol1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   855
      MultiLine       =   -1  'True
      TabIndex        =   33
      Text            =   "frm5x5edit.frx":1272
      Top             =   180
      Width           =   195
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   32
      Top             =   3570
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stage Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   45
      TabIndex        =   30
      Top             =   2835
      Width           =   2490
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   90
         MaxLength       =   15
         TabIndex        =   31
         Text            =   "Stage"
         Top             =   315
         Width           =   2265
      End
   End
   Begin VB.TextBox txtCol5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   2115
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "frm5x5edit.frx":1276
      Top             =   180
      Width           =   195
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   675
      TabIndex        =   3
      Top             =   900
      Width           =   1860
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   28
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   27
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   26
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   225
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   23
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   22
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   21
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   20
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   19
         Top             =   540
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   18
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   17
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   16
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   15
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   855
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   11
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   10
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   1170
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   8
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   1485
         Width           =   300
      End
   End
   Begin VB.TextBox txtCol4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "frm5x5edit.frx":127A
      Top             =   180
      Width           =   195
   End
   Begin VB.TextBox txtCol3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1485
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frm5x5edit.frx":127E
      Top             =   180
      Width           =   195
   End
   Begin VB.TextBox txtCol2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   1170
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frm5x5edit.frx":1282
      Top             =   180
      Width           =   195
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.pcs|*.pcs"
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu open 
         Caption         =   "&Open.."
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "&Save As.."
         Shortcut        =   ^S
      End
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "E&dit"
      Begin VB.Menu clear 
         Caption         =   "Clear All"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu line 
         Caption         =   "-"
      End
      Begin VB.Menu protect 
         Caption         =   "Protect"
      End
   End
End
Attribute VB_Name = "frm5x5edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub clear_Click()
Dim PicNum As Integer
PicNum = 1

Do While PicNum < 26
    Square(PicNum).Picture = LoadPicture()
    PicEDIT(PicNum) = False
    PicNum = PicNum + 1
Loop

txtDescription = "Stage"

txtRow1.Text = "0"
txtRow2.Text = "0"
txtRow3.Text = "0"
txtRow4.Text = "0"
txtRow5.Text = "0"
txtCol1.Text = "0"
txtCol2.Text = "0"
txtCol3.Text = "0"
txtCol4.Text = "0"
txtCol5.Text = "0"

End Sub

Private Sub exit_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()
Protected = False
Dim Tmp As Integer
Tmp = 1

Do While Tmp < 26
    PicEDIT(Tmp) = False
Tmp = Tmp + 1
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub open_Click()

On Error GoTo errorhandler:

Dim Tmp As Integer
Dim Tmp2 As Boolean
Tmp = 1

CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Col(1), Col(2), Col(3), Col(4), Col(5), Description, Tmp2
Close #1

If StageSize = "5x5" And Tmp2 = False Then
    
    Do While Tmp < 26
        If Pic(Tmp) = True Then
            Square(Tmp).Picture = LoadPicture(App.Path & "\true.gif")
            PicEDIT(Tmp) = True
        ElseIf Pic(Tmp) = False Then
            Square(Tmp).Picture = LoadPicture()
            PicEDIT(Tmp) = False
        End If
    Tmp = Tmp + 1
    Loop
    
    txtRow1.Text = Row(1)
    txtRow2.Text = Row(2)
    txtRow3.Text = Row(3)
    txtRow4.Text = Row(4)
    txtRow5.Text = Row(5)
    
    txtCol1.Text = Col(1)
    txtCol2.Text = Col(2)
    txtCol3.Text = Col(3)
    txtCol4.Text = Col(4)
    txtCol5.Text = Col(5)
    
    txtDescription.Text = Description
    StatusBar1.SimpleText = "Load OK! -- " & CommonDialog1.FileTitle
ElseIf Tmp2 = True Then
    StatusBar1.SimpleText = "ERROR! This file is protected!"
Else
    StatusBar1.SimpleText = "ERROR! Grid Size not correct!"
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub protect_Click()

If protect.Checked = False Then
    protect.Checked = True
    MsgBox "Warning: You and/or anyone else will not be able to open this file for editing once it is saved!"
    Protected = True
ElseIf protect.Checked = True Then
    protect.Checked = False
    Protected = False
End If

End Sub

Private Sub save_Click()
On Error GoTo errorhandler
CommonDialog1.ShowSave

Open CommonDialog1.FileName For Output As #1
Write #1, "5x5", PicEDIT(1), PicEDIT(2), PicEDIT(3), PicEDIT(4), PicEDIT(5), PicEDIT(6), PicEDIT(7), PicEDIT(8), PicEDIT(9), PicEDIT(10), PicEDIT(11), PicEDIT(12), PicEDIT(13), PicEDIT(14), PicEDIT(15), PicEDIT(16), PicEDIT(17), PicEDIT(18), PicEDIT(19), PicEDIT(20), PicEDIT(21), PicEDIT(22), PicEDIT(23), PicEDIT(24), PicEDIT(25)
Write #1, txtRow1.Text, txtRow2.Text, txtRow3.Text, txtRow4.Text, txtRow5.Text, txtCol1.Text, txtCol2.Text, txtCol3.Text, txtCol4.Text, txtCol5.Text, txtDescription.Text, Protected
Close #1

StatusBar1.SimpleText = "Stage Saved Successfully!"

errorhandler:
    Select Case Err
    Case Is = 75
    End Select
End Sub

Private Sub Square_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
    PicEDIT(Index) = True
    Square(Index).Picture = LoadPicture(App.Path & "\true.gif")
ElseIf Button = 2 Then
    PicEDIT(Index) = False
    Square(Index).Picture = LoadPicture()
End If

Call CalcRow1
Call CalcRow2
Call CalcRow3
Call CalcRow4
Call CalcRow5
Call CalcCol1
Call CalcCol2
Call CalcCol3
Call CalcCol4
Call CalcCol5
End Sub

Private Sub CalcRow1()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 1
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 6
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & RowCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow1.Text = "0" Else txtRow1.Text = RowEdit
End Sub

Private Sub CalcRow2()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 6
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 11
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & RowCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow2.Text = "0" Else txtRow2.Text = RowEdit
End Sub

Private Sub CalcRow3()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 11
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 16
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & RowCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow3.Text = "0" Else txtRow3.Text = RowEdit
End Sub

Private Sub CalcRow4()

Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 16
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 21
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & RowCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow4.Text = "0" Else txtRow4.Text = RowEdit
End Sub

Private Sub CalcRow5()
Dim RowCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 21
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 26
    If PicEDIT(PicCount) = True And LastEdit = True Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        RowCount(SkipNum) = RowCount(SkipNum) + 1
        PicCount = PicCount + 1
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 1
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            RowEdit = RowEdit & RowCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow5.Text = "0" Else txtRow5.Text = RowEdit
End Sub

Private Sub CalcCol1()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 1
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 22
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & ColCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol1.Text = 0 Else txtCol1.Text = ColEdit
End Sub

Private Sub CalcCol2()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 2
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 23
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & ColCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol2.Text = 0 Else txtCol2.Text = ColEdit
End Sub

Private Sub CalcCol3()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 3
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 24
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & ColCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol3.Text = 0 Else txtCol3.Text = ColEdit
End Sub

Private Sub CalcCol4()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 4
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 25
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & ColCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol4.Text = 0 Else txtCol4.Text = ColEdit

End Sub

Private Sub CalcCol5()

Dim ColCount(1 To 5) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim Tmp As Integer
Dim First As Boolean

PicCount = 5
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 26
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 5
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 5
    End If
Loop

Tmp = 1
Do Until Tmp = 6
    If First = False Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(Tmp)
            Tmp = Tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(Tmp) = 0 Then
            Tmp = Tmp + 1
        Else
            ColEdit = ColEdit & ColCount(Tmp)
            Tmp = Tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol5.Text = 0 Else txtCol5.Text = ColEdit

End Sub
