VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MT's Picross"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":1272
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   322
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1995
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "*.pcs|*.pcs"
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   161
      X2              =   161
      Y1              =   147
      Y2              =   168
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Load a new stage off          the Internet!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2520
      TabIndex        =   8
      Top             =   1995
      Width           =   2115
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   84
      X2              =   161
      Y1              =   147
      Y2              =   147
   End
   Begin VB.Image Image7 
      Height          =   375
      Left            =   105
      Picture         =   "frmMain.frx":5038
      Top             =   1995
      Width           =   1410
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   4410
      TabIndex        =   7
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3990
      TabIndex        =   6
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3570
      TabIndex        =   5
      Top             =   0
      Width           =   435
   End
   Begin VB.Image Image6 
      Height          =   375
      Left            =   2730
      Picture         =   "frmMain.frx":516E
      Top             =   2625
      Width           =   1815
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   315
      Picture         =   "frmMain.frx":54BA
      Top             =   2625
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   135
      Picture         =   "frmMain.frx":57E9
      Top             =   1590
      Width           =   1410
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   180
      Picture         =   "frmMain.frx":5A90
      Top             =   1185
      Width           =   1410
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   180
      Picture         =   "frmMain.frx":5D1C
      Top             =   780
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   135
      Picture         =   "frmMain.frx":5FAA
      Top             =   375
      Width           =   1410
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   108
      X2              =   162
      Y1              =   124
      Y2              =   124
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Load a custom stage!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2610
      TabIndex        =   4
      Top             =   1680
      Width           =   2085
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   114
      X2              =   129
      Y1              =   97
      Y2              =   97
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stages have a 10x10 grid."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2070
      TabIndex        =   3
      Top             =   1275
      Width           =   2625
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   111
      X2              =   141
      Y1              =   70
      Y2              =   70
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Stages have a 5x5 grid."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2295
      TabIndex        =   2
      Top             =   870
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Learn to play the game!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2340
      TabIndex        =   1
      Top             =   420
      Width           =   2310
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   111
      X2              =   147
      Y1              =   43
      Y2              =   43
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   147
      X2              =   147
      Y1              =   43
      Y2              =   34
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   141
      X2              =   141
      Y1              =   70
      Y2              =   61
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   129
      X2              =   129
      Y1              =   97
      Y2              =   91
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   162
      X2              =   162
      Y1              =   124
      Y2              =   115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MT's Picross! --- Version 0.4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   105
      TabIndex        =   0
      Top             =   0
      Width           =   2265
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Visible         =   0   'False
      Begin VB.Menu starteasy 
         Caption         =   "Start Easy Level"
      End
      Begin VB.Menu starthard 
         Caption         =   "Start Hard Level"
      End
      Begin VB.Menu loadcustom 
         Caption         =   "Load Custom Level"
      End
      Begin VB.Menu netlevel 
         Caption         =   "Load Net Level"
      End
      Begin VB.Menu break1 
         Caption         =   "-"
      End
      Begin VB.Menu associate 
         Caption         =   "Associate .pcs Files"
      End
      Begin VB.Menu break2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu easyedit 
         Caption         =   "5 x 5 Editor"
      End
      Begin VB.Menu hardedit 
         Caption         =   "10 x 10 Editor"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Visible         =   0   'False
      Begin VB.Menu loadtutorial 
         Caption         =   "Load Tutorial"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub associate_Click()
rc = MsgBox("This will associate .pcs files with MT's Picross. This will allow you to load a stage by double clicking on the stage file. Continue?", vbYesNo, "Associate?")
If rc = vbYes Then
    Call SaveString(HKEY_CLASSES_ROOT, ".pcs", "", "pcsfile")
    Call SaveString(HKEY_CLASSES_ROOT, ".pcs", "Content Type", "text/plain")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile", "", "MT's Picross Stage")
    Call SaveDWord(HKEY_CLASSES_ROOT, "pcsfile", "EditFlags", "0000")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\DefaultIcon", "", App.Path & "\" & App.EXEName & ".exe,0")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell", "", "")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell\Open", "", "")
    Call SaveString(HKEY_CLASSES_ROOT, "pcsfile\Shell\Open\Command", "", App.Path & "\" & App.EXEName & ".exe %1")
    MsgBox "Association Complete!"
Else
End If
End Sub

Private Sub easyedit_Click()

frm5x5edit.Show
Me.Hide

End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()

If Command$ <> "%1" And Command$ <> "" And Done = False Then
    Open Command$ For Input As #1
    Input #1, StageSize
    Close #1

    If StageSize = "5x5" Then
        Stage = Command$
        frm5x5.Show
        Me.Hide
        Custom = True
    ElseIf StageSize = "10x10" Then
        Stage = Command$
        frm10x10.Show
        Me.Hide
        Custom = True
    End If
End If
       
End Sub

Private Sub hardedit_Click()

frm10x10edit.Show
Me.Hide

End Sub

Private Sub Image1_Click()

frmTutorial.Show
Me.Hide

End Sub

Private Sub Image2_Click()

frm5x5Select.Show
Me.Hide

End Sub

Private Sub Image3_Click()

frm10x10select.Show
Me.Hide

End Sub

Private Sub Image4_Click()

On Error GoTo errorhandler

CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Input As #1
Input #1, StageSize
Close #1

If StageSize = "5x5" Then
    Stage = CommonDialog1.FileName
    frm5x5.Show
    Me.Hide
    Custom = True
ElseIf StageSize = "10x10" Then
    Stage = CommonDialog1.FileName
    frm10x10.Show
    Me.Hide
    Custom = True
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub Image5_Click()

frm5x5edit.Show
Me.Hide

End Sub

Private Sub Image6_Click()

frm10x10edit.Show
Me.Hide
End Sub

Private Sub Image7_Click()
frmNetLevels.Show
Me.Hide
End Sub

Private Sub Label7_Click()
Me.PopupMenu file
End Sub

Private Sub Label8_Click()
Me.PopupMenu edit
End Sub

Private Sub Label9_Click()
Me.PopupMenu help
End Sub

Private Sub loadcustom_Click()
On Error GoTo errorhandler

CommonDialog1.ShowOpen

Open CommonDialog1.FileName For Input As #1
Input #1, StageSize
Close #1

If StageSize = "5x5" Then
    Stage = CommonDialog1.FileName
    frm5x5.Show
    Me.Hide
    Custom = True
ElseIf StageSize = "10x10" Then
    Stage = CommonDialog1.FileName
    frm10x10.Show
    Me.Hide
    Custom = True
End If

errorhandler:
    Select Case Err
    Case Is = 75
    End Select

End Sub

Private Sub loadtutorial_Click()

frmTutorial.Show
Me.Hide

End Sub

Private Sub netlevel_Click()
frmNetLevels.Show
Me.Hide
End Sub

Private Sub starteasy_Click()

frm5x5Select.Show
Me.Hide

End Sub

Private Sub starthard_Click()

frm10x10select.Show
Me.Hide

End Sub
