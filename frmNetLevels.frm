VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmNetLevels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Levels"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4965
   Icon            =   "frmNetLevels.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNetLevels.frx":1272
   ScaleHeight     =   3120
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "frmNetLevels.frx":41D2
      Left            =   105
      List            =   "frmNetLevels.frx":41D4
      TabIndex        =   10
      Top             =   315
      Width           =   2220
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   225
      Left            =   1470
      TabIndex        =   9
      Top             =   2295
      Width           =   855
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3675
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2835
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3255
      Top             =   1680
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   210
      Top             =   2205
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   2895
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   397
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Download List"
      Height          =   330
      Left            =   2940
      TabIndex        =   0
      Top             =   2205
      Width           =   1170
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   2415
      X2              =   2415
      Y1              =   0
      Y2              =   2625
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   2415
      X2              =   0
      Y1              =   2625
      Y2              =   2625
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Stage List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   840
      TabIndex        =   11
      Top             =   0
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   2835
      X2              =   4935
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   2835
      X2              =   2835
      Y1              =   0
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Size:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   8
      Top             =   315
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   7
      Top             =   735
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Created By:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2940
      TabIndex        =   6
      Top             =   1050
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   3885
      TabIndex        =   5
      Top             =   315
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   3885
      TabIndex        =   4
      Top             =   735
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   2835
      TabIndex        =   3
      Top             =   1365
      Width           =   2115
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Stage Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3360
      TabIndex        =   2
      Top             =   0
      Width           =   1380
   End
End
Attribute VB_Name = "frmNetLevels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
FileOK = True
Timer1.Enabled = True
URL = "http://mtproductions.tripod.com/netpicross/stagelist.dat"
Call DownloadFile
Call CheckFile(App.Path & "\stagelist.dat")
If FileOK = False Then
    StatusBar1.SimpleText = "File Error. Please try again."
    Timer1.Enabled = False
    Exit Sub
End If

End Sub

Private Sub Command2_Click()
FileOK = True
Timer3.Enabled = True
URL = "http://mtproductions.tripod.com/netpicross/" & List1.Text & ".pcs"
Call DownloadFile
Call CheckFile(App.Path & "\" & List1.Text & ".pcs")
If FileOK = False Then
    StatusBar1.SimpleText = "File Error. Please try again."
    Timer3.Enabled = False
    Exit Sub
End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub List1_Click()
FileOK = True
If List1.Text = "" Or List1.Text = "Coming Soon" Then
    Command2.Enabled = False
Else
    Timer2.Enabled = True
    URL = "http://mtproductions.tripod.com/netpicross/" & List1.Text & ".dat"
    Call DownloadFile
    Call CheckFile(App.Path & "\" & List1.Text & ".dat")
    If FileOK = False Then
        StatusBar1.SimpleText = "File Error. Please try again."
        Label4.Caption = "ERROR"
        Label5.Caption = "ERROR"
        Label6.Caption = "ERROR"
        Timer2.Enabled = False
        Exit Sub
    End If
    Command2.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
Dim MyString As String

If Inet1.StillExecuting = True Then
    StatusBar1.SimpleText = "Downloading List, Please Wait."
Else
    List1.clear
    Open App.Path & "\stagelist.dat" For Input As #1
    While Not EOF(1)
    Input #1, MyString$
    DoEvents
    List1.AddItem MyString$
    List1.Refresh
    Wend
    Close #1
    Kill App.Path & "\stagelist.dat"
    StatusBar1.SimpleText = "Ready...."
    Timer1.Enabled = False
End If
End Sub

Private Sub DownloadFile()
On Error GoTo errorhandler
Dim myData() As Byte
If Inet1.StillExecuting = True Then Exit Sub
myData() = Inet1.OpenURL(URL, icByteArray)
For X = Len(URL) To 1 Step -1
If Left$(Right$(URL, X), 1) = "/" Then RealFile$ = Right$(URL, X - 1)
Next X
myFile$ = App.Path + "\" + RealFile$
Open myFile$ For Binary Access Write As #1
Put #1, , myData()
Close #1

Exit Sub

errorhandler:
    Timer1.Enabled = False
    StatusBar1.SimpleText = "File error! Try again later."

End Sub

Private Sub Timer2_Timer()
Dim Grid As String
Dim Description As String
Dim Created As String

If Inet1.StillExecuting = True Then
    StatusBar1.SimpleText = "Downloading Info, Please Wait."
Else
    Open App.Path & "\" & List1.Text & ".dat" For Input As #1
    Input #1, Grid, Description, Created
    Close #1
    Kill App.Path & "\" & List1.Text & ".dat"
    Label4.Caption = Grid
    Label5.Caption = Description
    Label6.Caption = Created
    StatusBar1.SimpleText = "Ready...."
    Timer2.Enabled = False
End If

End Sub

Private Sub Timer3_Timer()
If Inet1.StillExecuting = True Then
    StatusBar1.SimpleText = "Loading Stage, Please Wait."
Else
    Stage = App.Path & "\" & List1.Text & ".pcs"
    Open App.Path & "\" & List1.Text & ".pcs" For Input As #1
    Input #1, StageSize
    Close #1

    If StageSize = "5x5" Then
        frm5x5.Show
        Me.Hide
        NetGame = True
    ElseIf StageSize = "10x10" Then
        frm10x10.Show
        Me.Hide
        NetGame = True
    End If
    Timer3.Enabled = False
End If

End Sub

Public Sub CheckFile(file As String)
On Error GoTo errorhandler

Dim tmp As String

Open file For Input As #1
Input #1, tmp
Close #1
Tmp2 = Len(tmp)

If InStr(1, tmp, "<!") Then FileOK = False Else FileOK = True
Exit Sub

errorhandler:
    FileOK = False
    Close #1
End Sub

