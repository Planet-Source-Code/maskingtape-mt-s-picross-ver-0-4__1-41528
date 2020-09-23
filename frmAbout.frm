VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MT's Picross"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   735
      TabIndex        =   3
      Top             =   0
      Width           =   3270
   End
   Begin VB.Label Label3 
      Caption         =   "Copyright (c) 2002 - 2002 MT Productions - All Rights Reserved."
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   1470
      Width           =   4530
   End
   Begin VB.Label Label2 
      Caption         =   "Ver: 0.4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3465
      TabIndex        =   1
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "http://mtproductions.tripod.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   2310
      MouseIcon       =   "frmAbout.frx":1272
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1890
      Width           =   2325
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
ShellExecute hWnd, "open", "http://mtproductions.tripod.com", vbNullString, vbNullString, conSwNormal
End Sub
