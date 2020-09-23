VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm10x10edit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: 10x10 Editor"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   4515
   Icon            =   "frm10x10edit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   900
      TabIndex        =   23
      Top             =   1125
      Width           =   3435
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   123
         Top             =   810
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
         TabIndex        =   122
         Top             =   810
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
         TabIndex        =   121
         Top             =   810
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
         TabIndex        =   120
         Top             =   810
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
         TabIndex        =   119
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   118
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   117
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   116
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   115
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   114
         Top             =   495
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
         TabIndex        =   113
         Top             =   495
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
         TabIndex        =   112
         Top             =   495
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
         TabIndex        =   111
         Top             =   495
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
         TabIndex        =   110
         Top             =   495
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
         TabIndex        =   109
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   108
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   107
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   106
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   105
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   104
         Top             =   180
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
         TabIndex        =   103
         Top             =   180
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
         TabIndex        =   102
         Top             =   180
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
         TabIndex        =   101
         Top             =   180
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
         TabIndex        =   100
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   99
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   26
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   98
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   27
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   97
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   28
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   96
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   29
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   95
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   30
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   94
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   31
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   93
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   32
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   92
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   33
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   91
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   34
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   90
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   35
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   89
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   36
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   88
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   37
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   87
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   38
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   86
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   39
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   85
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   40
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   84
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   41
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   83
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   42
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   82
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   43
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   81
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   44
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   80
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   45
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   79
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   46
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   78
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   47
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   77
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   48
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   76
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   49
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   75
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   50
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   74
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   51
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   73
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   52
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   72
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   53
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   71
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   54
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   70
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   55
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   69
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   56
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   68
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   57
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   67
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   58
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   66
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   59
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   65
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   60
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   64
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   61
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   63
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   62
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   62
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   63
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   61
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   64
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   60
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   65
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   59
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   66
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   58
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   67
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   57
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   68
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   56
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   69
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   55
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   70
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   54
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   71
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   53
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   72
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   52
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   73
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   51
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   74
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   50
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   75
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   49
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   76
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   48
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   77
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   47
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   78
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   46
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   79
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   45
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   80
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   44
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   81
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   43
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   82
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   42
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   83
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   41
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   84
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   40
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   85
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   39
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   86
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   38
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   87
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   37
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   88
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   36
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   89
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   35
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   90
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   34
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   91
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   33
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   92
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   93
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   31
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   94
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   30
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   95
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   29
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   96
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   28
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   97
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   27
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   98
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   26
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   99
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   100
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   3015
         Width           =   300
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   22
      Top             =   5385
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   423
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
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
      Left            =   1080
      TabIndex        =   20
      Top             =   4590
      Width           =   2490
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   90
         MaxLength       =   15
         TabIndex        =   21
         Text            =   "Stage"
         Top             =   315
         Width           =   2265
      End
   End
   Begin VB.TextBox txtCol10 
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
      Height          =   1050
      Left            =   3870
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frm10x10edit.frx":1272
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtCol9 
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
      Height          =   1050
      Left            =   3555
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frm10x10edit.frx":1276
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtCol8 
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
      Height          =   1050
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   17
      Text            =   "frm10x10edit.frx":127A
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtCol7 
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
      Height          =   1050
      Left            =   2925
      MultiLine       =   -1  'True
      TabIndex        =   16
      Text            =   "frm10x10edit.frx":127E
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtCol6 
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
      Height          =   1050
      Left            =   2610
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "frm10x10edit.frx":1282
      Top             =   45
      Width           =   285
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
      Height          =   1050
      Left            =   2295
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frm10x10edit.frx":1286
      Top             =   45
      Width           =   285
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
      Height          =   1050
      Left            =   1980
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frm10x10edit.frx":128A
      Top             =   45
      Width           =   285
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
      Height          =   1050
      Left            =   1665
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "frm10x10edit.frx":128E
      Top             =   45
      Width           =   285
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
      Height          =   1050
      Left            =   1350
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frm10x10edit.frx":1292
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtCol1 
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
      Height          =   1050
      Left            =   1035
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "frm10x10edit.frx":1296
      Top             =   45
      Width           =   285
   End
   Begin VB.TextBox txtRow10 
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
      TabIndex        =   9
      Text            =   "0"
      Top             =   4140
      Width           =   825
   End
   Begin VB.TextBox txtRow9 
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
      TabIndex        =   8
      Text            =   "0"
      Top             =   3825
      Width           =   825
   End
   Begin VB.TextBox txtRow8 
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
      TabIndex        =   7
      Text            =   "0"
      Top             =   3510
      Width           =   825
   End
   Begin VB.TextBox txtRow7 
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
      TabIndex        =   6
      Text            =   "0"
      Top             =   3195
      Width           =   825
   End
   Begin VB.TextBox txtRow6 
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
      TabIndex        =   5
      Text            =   "0"
      Top             =   2880
      Width           =   825
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
      TabIndex        =   4
      Text            =   "0"
      Top             =   1305
      Width           =   825
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
      TabIndex        =   3
      Text            =   "0"
      Top             =   1620
      Width           =   825
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
      TabIndex        =   2
      Text            =   "0"
      Top             =   1935
      Width           =   825
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
      TabIndex        =   1
      Text            =   "0"
      Top             =   2250
      Width           =   825
   End
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
      TabIndex        =   0
      Text            =   "0"
      Top             =   2565
      Width           =   825
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   180
      Top             =   180
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
      Caption         =   "&Edit"
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
Attribute VB_Name = "frm10x10edit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub clear_Click()

Dim PicNum As Integer
PicNum = 1

Do While PicNum < 101
    Square(PicNum).Picture = LoadPicture()
    PicEDIT(PicNum) = False
    PicNum = PicNum + 1
Loop

txtRow1.Text = "0"
txtRow2.Text = "0"
txtRow3.Text = "0"
txtRow4.Text = "0"
txtRow5.Text = "0"
txtRow6.Text = "0"
txtRow7.Text = "0"
txtRow8.Text = "0"
txtRow9.Text = "0"
txtRow10.Text = "0"

txtCol1.Text = "0"
txtCol2.Text = "0"
txtCol3.Text = "0"
txtCol4.Text = "0"
txtCol5.Text = "0"
txtCol6.Text = "0"
txtCol7.Text = "0"
txtCol8.Text = "0"
txtCol9.Text = "0"
txtCol10.Text = "0"

txtDescription = "Stage"

End Sub

Private Sub exit_Click()

frmMain.Show
Unload Me

End Sub

Private Sub Form_Load()

Protected = False
Dim tmp As Integer
tmp = 1

Do While tmp < 101
    PicEDIT(tmp) = False
tmp = tmp + 1
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub open_Click()

On Error GoTo errorhandler:

Dim tmp As Integer
Dim Tmp2
tmp = 1

CommonDialog1.ShowOpen
Open CommonDialog1.FileName For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25), Pic(26), Pic(27), Pic(28), Pic(29), Pic(30), Pic(31), Pic(32), Pic(33), Pic(34), Pic(35), Pic(36), Pic(37), Pic(38), Pic(39), Pic(40), Pic(41), Pic(42), Pic(43), Pic(44), Pic(45), Pic(46), Pic(47), Pic(48), Pic(49), Pic(50)
Input #1, Pic(51), Pic(52), Pic(53), Pic(54), Pic(55), Pic(56), Pic(57), Pic(58), Pic(59), Pic(60), Pic(61), Pic(62), Pic(63), Pic(64), Pic(65), Pic(66), Pic(67), Pic(68), Pic(69), Pic(70), Pic(71), Pic(72), Pic(73), Pic(74), Pic(75), Pic(76), Pic(77), Pic(78), Pic(79), Pic(80), Pic(81), Pic(82), Pic(83), Pic(84), Pic(85), Pic(86), Pic(87), Pic(88), Pic(89), Pic(90), Pic(91), Pic(92), Pic(93), Pic(94), Pic(95), Pic(96), Pic(97), Pic(98), Pic(99), Pic(100)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Row(6), Row(7), Row(8), Row(9), Row(10), Col(1), Col(2), Col(3), Col(4), Col(5), Col(6), Col(7), Col(8), Col(9), Col(10), Description, Tmp2
Close #1

If StageSize = "10x10" And Tmp2 = False Then
    
    Do While tmp < 101
        If Pic(tmp) = True Then
            Square(tmp).Picture = LoadPicture(App.Path & "\true.gif")
            PicEDIT(tmp) = True
        ElseIf Pic(tmp) = False Then
            Square(tmp).Picture = LoadPicture()
            PicEDIT(tmp) = False
        End If
    tmp = tmp + 1
    Loop
    
    txtRow1.Text = Row(1)
    txtRow2.Text = Row(2)
    txtRow3.Text = Row(3)
    txtRow4.Text = Row(4)
    txtRow5.Text = Row(5)
    txtRow6.Text = Row(6)
    txtRow7.Text = Row(7)
    txtRow8.Text = Row(8)
    txtRow9.Text = Row(9)
    txtRow10.Text = Row(10)
        
    txtCol1.Text = Col(1)
    txtCol2.Text = Col(2)
    txtCol3.Text = Col(3)
    txtCol4.Text = Col(4)
    txtCol5.Text = Col(5)
    txtCol6.Text = Col(6)
    txtCol7.Text = Col(7)
    txtCol8.Text = Col(8)
    txtCol9.Text = Col(9)
    txtCol10.Text = Col(10)

    
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
Write #1, "10x10", PicEDIT(1), PicEDIT(2), PicEDIT(3), PicEDIT(4), PicEDIT(5), PicEDIT(6), PicEDIT(7), PicEDIT(8), PicEDIT(9), PicEDIT(10), PicEDIT(11), PicEDIT(12), PicEDIT(13), PicEDIT(14), PicEDIT(15), PicEDIT(16), PicEDIT(17), PicEDIT(18), PicEDIT(19), PicEDIT(20), PicEDIT(21), PicEDIT(22), PicEDIT(23), PicEDIT(24), PicEDIT(25), PicEDIT(26), PicEDIT(27), PicEDIT(28), PicEDIT(29), PicEDIT(30), PicEDIT(31), PicEDIT(32), PicEDIT(33), PicEDIT(34), PicEDIT(35), PicEDIT(36), PicEDIT(37), PicEDIT(38), PicEDIT(39), PicEDIT(40), PicEDIT(41), PicEDIT(42), PicEDIT(43), PicEDIT(44), PicEDIT(45), PicEDIT(46), PicEDIT(47), PicEDIT(48), PicEDIT(49), PicEDIT(50)
Write #1, PicEDIT(51), PicEDIT(52), PicEDIT(53), PicEDIT(54), PicEDIT(55), PicEDIT(56), PicEDIT(57), PicEDIT(58), PicEDIT(59), PicEDIT(60), PicEDIT(61), PicEDIT(62), PicEDIT(63), PicEDIT(64), PicEDIT(65), PicEDIT(66), PicEDIT(67), PicEDIT(68), PicEDIT(69), PicEDIT(70), PicEDIT(71), PicEDIT(72), PicEDIT(73), PicEDIT(74), PicEDIT(75), PicEDIT(76), PicEDIT(77), PicEDIT(78), PicEDIT(79), PicEDIT(80), PicEDIT(81), PicEDIT(82), PicEDIT(83), PicEDIT(84), PicEDIT(85), PicEDIT(86), PicEDIT(87), PicEDIT(88), PicEDIT(89), PicEDIT(90), PicEDIT(91), PicEDIT(92), PicEDIT(93), PicEDIT(94), PicEDIT(95), PicEDIT(96), PicEDIT(97), PicEDIT(98), PicEDIT(99), PicEDIT(100)
Write #1, txtRow1.Text, txtRow2.Text, txtRow3.Text, txtRow4.Text, txtRow5.Text, txtRow6.Text, txtRow7.Text, txtRow8.Text, txtRow9.Text, txtRow10.Text, txtCol1.Text, txtCol2.Text, txtCol3.Text, txtCol4.Text, txtCol5.Text, txtCol6.Text, txtCol7.Text, txtCol8.Text, txtCol9.Text, txtCol10.Text, txtDescription.Text, Protected
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
Call CalcRow6
Call CalcRow7
Call CalcRow8
Call CalcRow9
Call CalcRow10
Call CalcCol1
Call CalcCol2
Call CalcCol3
Call CalcCol4
Call CalcCol5
Call CalcCol6
Call CalcCol7
Call CalcCol8
Call CalcCol9
Call CalcCol10
End Sub

Private Sub CalcRow1()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 1
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow1.Text = "0" Else txtRow1.Text = RowEdit

End Sub

Private Sub CalcRow2()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 11
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow2.Text = "0" Else txtRow2.Text = RowEdit

End Sub

Private Sub CalcRow3()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 21
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 31
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow3.Text = "0" Else txtRow3.Text = RowEdit

End Sub

Private Sub CalcRow4()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 31
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 41
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow4.Text = "0" Else txtRow4.Text = RowEdit

End Sub

Private Sub CalcRow5()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 41
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 51
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow5.Text = "0" Else txtRow5.Text = RowEdit

End Sub

Private Sub CalcRow6()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 51
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 61
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow6.Text = "0" Else txtRow6.Text = RowEdit

End Sub

Private Sub CalcRow7()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 61
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 71
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow7.Text = "0" Else txtRow7.Text = RowEdit

End Sub

Private Sub CalcRow8()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 71
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 81
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow8.Text = "0" Else txtRow8.Text = RowEdit

End Sub

Private Sub CalcRow9()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 81
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 91
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow9.Text = "0" Else txtRow9.Text = RowEdit

End Sub

Private Sub CalcRow10()

Dim RowCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim RowEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 91
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 101
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

tmp = 1
Do Until tmp = 11
    If First = False Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & "-" & RowCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If RowCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            RowEdit = RowEdit & RowCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If RowEdit = "" Then txtRow10.Text = "0" Else txtRow10.Text = RowEdit

End Sub

Private Sub CalcCol1()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 1
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 92
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol1.Text = 0 Else txtCol1.Text = ColEdit

End Sub

Private Sub CalcCol2()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 2
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 93
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol2.Text = 0 Else txtCol2.Text = ColEdit

End Sub

Private Sub CalcCol3()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 3
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 94
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol3.Text = 0 Else txtCol3.Text = ColEdit

End Sub

Private Sub CalcCol4()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 4
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 95
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol4.Text = 0 Else txtCol4.Text = ColEdit

End Sub

Private Sub CalcCol5()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 5
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 96
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol5.Text = 0 Else txtCol5.Text = ColEdit

End Sub

Private Sub CalcCol6()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 6
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 97
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol6.Text = 0 Else txtCol6.Text = ColEdit
End Sub

Private Sub CalcCol7()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 7
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 98
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol7.Text = 0 Else txtCol7.Text = ColEdit

End Sub

Private Sub CalcCol8()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 8
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 99
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol8.Text = 0 Else txtCol8.Text = ColEdit

End Sub

Private Sub CalcCol9()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 9
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 100
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol9.Text = 0 Else txtCol9.Text = ColEdit

End Sub

Private Sub CalcCol10()

Dim ColCount(1 To 10) As Integer
Dim PicCount As Integer
Dim SkipNum As Integer
Dim LastEdit As Boolean
Dim ColEdit As String
Dim tmp As Integer
Dim First As Boolean

PicCount = 10
LastEdit = False
SkipNum = 1
First = True

Do While PicCount < 101
    If PicEDIT(PicCount) = True And LastEdit = True Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        LastEdit = True
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = True And LastEdit = False Then
        ColCount(SkipNum) = ColCount(SkipNum) + 1
        PicCount = PicCount + 10
        LastEdit = True
    ElseIf PicEDIT(PicCount) = False And LastEdit = True Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    ElseIf PicEDIT(PicCount) = False And LastEdit = False Then
        SkipNum = SkipNum + 1
        LastEdit = False
        PicCount = PicCount + 10
    End If
Loop

tmp = 1
Do Until tmp = 11
    If First = False Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & vbCrLf & ColCount(tmp)
            tmp = tmp + 1
        End If
    ElseIf First = True Then
        If ColCount(tmp) = 0 Then
            tmp = tmp + 1
        Else
            ColEdit = ColEdit & ColCount(tmp)
            tmp = tmp + 1
            First = False
        End If
    End If
Loop

If ColEdit = "" Then txtCol10.Text = 0 Else txtCol10.Text = ColEdit

End Sub
