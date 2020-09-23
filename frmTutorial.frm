VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTutorial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: Tutorial"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5565
   ControlBox      =   0   'False
   Icon            =   "frmTutorial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   45
      Top             =   225
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4545
      MaskColor       =   &H8000000F&
      TabIndex        =   39
      Top             =   2880
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Caption         =   "Instructions"
      Height          =   2625
      Left            =   2880
      TabIndex        =   36
      Top             =   225
      Width           =   2490
      Begin VB.Label Label1 
         Caption         =   "Continue - ->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   405
         TabIndex        =   40
         Top             =   2250
         Width           =   1680
      End
      Begin VB.Label lblInstructions 
         Caption         =   "Welcome to MT's Picross! Let's Get started!"
         Height          =   2265
         Left            =   135
         TabIndex        =   37
         Top             =   270
         Width           =   2265
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   675
      TabIndex        =   0
      Top             =   945
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   1485
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   1485
         Width           =   300
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   38
      Top             =   2895
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4145
            MinWidth        =   4145
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "XXX"
            TextSave        =   "XXX"
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
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   2655
      X2              =   2655
      Y1              =   2340
      Y2              =   2835
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      Height          =   1815
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   1035
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   870
      Left            =   675
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label lblRow2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   34
      Top             =   1485
      Width           =   555
   End
   Begin VB.Label lblRow3 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   33
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label lblRow4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   32
      Top             =   2115
      Width           =   555
   End
   Begin VB.Label lblCol2 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1125
      TabIndex        =   30
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1440
      TabIndex        =   29
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol4 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   1755
      TabIndex        =   28
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblRow5 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   27
      Top             =   2430
      Width           =   555
   End
   Begin VB.Line Line1 
      X1              =   855
      X2              =   855
      Y1              =   90
      Y2              =   990
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   2340
      Y1              =   90
      Y2              =   990
   End
   Begin VB.Line Line3 
      X1              =   585
      X2              =   0
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line4 
      X1              =   585
      X2              =   0
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label lblRow1 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   45
      TabIndex        =   35
      Top             =   1170
      Width           =   555
   End
   Begin VB.Label lblCol5 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2070
      TabIndex        =   26
      Top             =   0
      Width           =   330
   End
   Begin VB.Label lblCol1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   810
      TabIndex        =   31
      Top             =   0
      Width           =   330
   End
End
Attribute VB_Name = "frmTutorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()

Tutorial = 1
Time = 300
StatusBar1.Panels(2).Text = Time

Open App.Path & "\stages\tutorial.pcs" For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Col(1), Col(2), Col(3), Col(4), Col(5), Description
Close #1

lblRow1.Caption = Row(1)
lblRow2.Caption = Row(2)
lblRow3.Caption = Row(3)
lblRow4.Caption = Row(4)
lblRow5.Caption = Row(5)

lblCol1.Caption = Col(1)
lblCol2.Caption = Col(2)
lblCol3.Caption = Col(3)
lblCol4.Caption = Col(4)
lblCol5.Caption = Col(5)

End Sub

Private Sub Label1_Click()
Tutorial = Tutorial + 1
If Tutorial = 1 Then
    Shape1.Visible = True
    lblInstructions.Caption = "The numbers above the grid tells you how many boxes you have to draw in the downward direction. For example:"
End If

If Tutorial = 2 Then
    Shape1.Width = 330
    Shape1.Left = 810
    lblInstructions.Caption = "The first column has a 3 over it. That means there are 3 blocks that you have to click on in that column."
    Square(11).Picture = LoadPicture(App.Path & "\true.gif")
    Square(16).Picture = LoadPicture(App.Path & "\true.gif")
    Square(21).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 3 Then
    Shape1.Visible = False
    Shape2.Visible = True
    lblInstructions.Caption = "The numbers on the left side of the grid tells you how many boxes you have to draw in a left to right direction. For Example:"
End If

If Tutorial = 4 Then
    Shape2.Top = 2070
    Shape2.Height = 375
    lblInstructions.Caption = "The fourth row from the top has a 5 next to it. That means there are 5 blocks you have to click on in that row."
    Square(16).Picture = LoadPicture(App.Path & "\true.gif")
    Square(17).Picture = LoadPicture(App.Path & "\true.gif")
    Square(18).Picture = LoadPicture(App.Path & "\true.gif")
    Square(19).Picture = LoadPicture(App.Path & "\true.gif")
    Square(20).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 5 Then
    Shape2.Top = 1485
    Shape2.Height = 375
    lblInstructions.Caption = "This row has a 1 and a 1 next to it, so that means you click 1 block and then another block seperated by at least 1 space."
    Square(7).Picture = LoadPicture(App.Path & "\true.gif")
    Square(9).Picture = LoadPicture(App.Path & "\true.gif")
End If

If Tutorial = 6 Then
    Shape2.Visible = False
    Line5.Visible = True
    lblInstructions.Caption = "This is the time limit. For each stage you only have a certain amount of time to compete it."
    Timer1.Enabled = True
End If

If Tutorial = 7 Then
    lblInstructions.Caption = "If you click on a bad square your time will be reduced by a certain amount. If your time goes down to zero, the game is over. But don't panic! Take your time!"
    Square(1).Picture = LoadPicture(App.Path & "\false.gif")
    StatusBar1.Panels(1).Text = "Miss! -40 seconds!"
    Time = Time - 40
End If

If Tutorial = 8 Then
    Line5.Visible = False
    StatusBar1.Panels(1).Text = ""
    Timer1.Enabled = False
    lblInstructions.Caption = "That's it! You now know everything you need to start playing. Good Luck!"
    Label1.Enabled = False
End If

End Sub

Private Sub Timer1_Timer()

Time = Time - 1
StatusBar1.Panels(2).Text = Time

End Sub
