VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frm10x10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MT's Picross: 10x10 PlayField"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   660
   ClientWidth     =   5085
   ControlBox      =   0   'False
   Icon            =   "frm10x10.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm10x10.frx":1272
   ScaleHeight     =   348
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1125
      Top             =   270
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00955800&
      ForeColor       =   &H00000000&
      Height          =   3435
      Left            =   1530
      TabIndex        =   0
      Top             =   1485
      Width           =   3435
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   100
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   100
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   99
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   99
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   98
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   98
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   97
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   97
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   96
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   96
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   95
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   95
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   94
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   94
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   93
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   93
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   92
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   92
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   91
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   91
         Top             =   3015
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   90
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   90
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   89
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   89
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   88
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   88
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   87
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   87
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   86
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   86
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   85
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   85
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   84
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   84
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   83
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   83
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   82
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   82
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   81
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   81
         Top             =   2700
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   80
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   80
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   79
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   79
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   78
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   78
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   77
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   77
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   76
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   76
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   75
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   75
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   74
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   74
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   73
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   73
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   72
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   72
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   71
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   71
         Top             =   2385
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   70
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   70
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   69
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   69
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   68
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   68
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   67
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   67
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   66
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   66
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   65
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   65
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   64
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   64
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   63
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   63
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   61
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   61
         Top             =   2070
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   60
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   60
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   59
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   59
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   58
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   58
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   57
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   57
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   56
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   56
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   55
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   55
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   54
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   54
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   53
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   53
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   52
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   52
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   51
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   51
         Top             =   1755
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   50
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   50
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   49
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   49
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   48
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   48
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   47
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   47
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   46
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   46
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   45
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   45
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   44
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   44
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   43
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   43
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   42
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   42
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   41
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   41
         Top             =   1440
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   40
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   40
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   39
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   39
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   38
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   38
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   37
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   37
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   36
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   36
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   35
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   35
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   34
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   34
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   33
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   33
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   32
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   32
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   31
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   31
         Top             =   1125
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   30
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   30
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   29
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   29
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   28
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   28
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   27
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   27
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   26
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   26
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   25
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   24
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   23
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   22
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   21
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   20
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   7
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   19
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   8
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   18
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   17
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   16
         Top             =   180
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   11
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   15
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   14
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   13
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   14
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   12
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   15
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   11
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   1710
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   10
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   2025
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   9
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   2340
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   8
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   2655
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   7
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   2970
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   6
         Top             =   495
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   21
         Left            =   135
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   22
         Left            =   450
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   765
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   3
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   1080
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   2
         Top             =   810
         Width           =   300
      End
      Begin VB.PictureBox Square 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1395
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   1
         Top             =   810
         Width           =   300
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   111
      Top             =   4935
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7939
            MinWidth        =   7939
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
   Begin VB.Label lblExit 
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   585
      TabIndex        =   122
      Top             =   630
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Line Line4 
      X1              =   318
      X2              =   318
      Y1              =   6
      Y2              =   99
   End
   Begin VB.Line Line3 
      X1              =   114
      X2              =   114
      Y1              =   6
      Y2              =   105
   End
   Begin VB.Label lblCol10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   4500
      TabIndex        =   121
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   4185
      TabIndex        =   120
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   3870
      TabIndex        =   119
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   3555
      TabIndex        =   118
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   3240
      TabIndex        =   117
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   2925
      TabIndex        =   116
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   2610
      TabIndex        =   115
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   2295
      TabIndex        =   114
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCOl2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   1980
      TabIndex        =   113
      Top             =   45
      Width           =   330
   End
   Begin VB.Label lblCol1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1500
      Left            =   1665
      TabIndex        =   112
      Top             =   45
      Width           =   330
   End
   Begin VB.Line Line2 
      X1              =   99
      X2              =   3
      Y1              =   321
      Y2              =   321
   End
   Begin VB.Line Line1 
      X1              =   99
      X2              =   3
      Y1              =   111
      Y2              =   111
   End
   Begin VB.Label lblRow10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   110
      Top             =   4500
      Width           =   1410
   End
   Begin VB.Label lblRow9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   109
      Top             =   4185
      Width           =   1410
   End
   Begin VB.Label lblRow8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   108
      Top             =   3870
      Width           =   1410
   End
   Begin VB.Label lblRow7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   107
      Top             =   3555
      Width           =   1410
   End
   Begin VB.Label lblRow6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   106
      Top             =   3240
      Width           =   1410
   End
   Begin VB.Label lblRow1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   105
      Top             =   1665
      Width           =   1410
   End
   Begin VB.Label lblRow2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   104
      Top             =   1980
      Width           =   1410
   End
   Begin VB.Label lblRow3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   103
      Top             =   2295
      Width           =   1410
   End
   Begin VB.Label lblRow4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   102
      Top             =   2610
      Width           =   1410
   End
   Begin VB.Label lblRow5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   45
      TabIndex        =   101
      Top             =   2925
      Width           =   1410
   End
   Begin VB.Menu game 
      Caption         =   "&Game"
      Begin VB.Menu exit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "frm10x10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()

If Custom = True Then
    Custom = False
    frmMain.Show
    Unload Me
ElseIf NetGame = True Then
    NetGame = False
    frmNetLevels.Show
    Kill Stage
    Unload Me
Else
    Call HardRefresh
    frm10x10select.Show
    Unload Me
End If

End Sub

Private Sub Form_Load()

IsTrue = 0
Time = 600
StatusBar1.Panels(2).Text = Time

Open Stage For Input As #1
Input #1, StageSize, Pic(1), Pic(2), Pic(3), Pic(4), Pic(5), Pic(6), Pic(7), Pic(8), Pic(9), Pic(10), Pic(11), Pic(12), Pic(13), Pic(14), Pic(15), Pic(16), Pic(17), Pic(18), Pic(19), Pic(20), Pic(21), Pic(22), Pic(23), Pic(24), Pic(25), Pic(26), Pic(27), Pic(28), Pic(29), Pic(30), Pic(31), Pic(32), Pic(33), Pic(34), Pic(35), Pic(36), Pic(37), Pic(38), Pic(39), Pic(40), Pic(41), Pic(42), Pic(43), Pic(44), Pic(45), Pic(46), Pic(47), Pic(48), Pic(49), Pic(50)
Input #1, Pic(51), Pic(52), Pic(53), Pic(54), Pic(55), Pic(56), Pic(57), Pic(58), Pic(59), Pic(60), Pic(61), Pic(62), Pic(63), Pic(64), Pic(65), Pic(66), Pic(67), Pic(68), Pic(69), Pic(70), Pic(71), Pic(72), Pic(73), Pic(74), Pic(75), Pic(76), Pic(77), Pic(78), Pic(79), Pic(80), Pic(81), Pic(82), Pic(83), Pic(84), Pic(85), Pic(86), Pic(87), Pic(88), Pic(89), Pic(90), Pic(91), Pic(92), Pic(93), Pic(94), Pic(95), Pic(96), Pic(97), Pic(98), Pic(99), Pic(100)
Input #1, Row(1), Row(2), Row(3), Row(4), Row(5), Row(6), Row(7), Row(8), Row(9), Row(10), Col(1), Col(2), Col(3), Col(4), Col(5), Col(6), Col(7), Col(8), Col(9), Col(10), Description
Close #1

lblRow1.Caption = Row(1)
lblRow2.Caption = Row(2)
lblRow3.Caption = Row(3)
lblRow4.Caption = Row(4)
lblRow5.Caption = Row(5)
lblRow6.Caption = Row(6)
lblRow7.Caption = Row(7)
lblRow8.Caption = Row(8)
lblRow9.Caption = Row(9)
lblRow10.Caption = Row(10)

lblCol1.Caption = Col(1)
lblCol2.Caption = Col(2)
lblCol3.Caption = Col(3)
lblCol4.Caption = Col(4)
lblCol5.Caption = Col(5)
lblCol6.Caption = Col(6)
lblCol7.Caption = Col(7)
lblCol8.Caption = Col(8)
lblCol9.Caption = Col(9)
lblCol10.Caption = Col(10)

Call CalcTrue

rc = MsgBox("Would you like a starting hint?", vbYesNo, "Hint?")
If rc = vbYes Then
    Call StartHint
Else
End If

End Sub

Private Sub Win()

Dim PicNum As Integer

PicNum = 1

Do While PicNum < 101
    If Pic(PicNum) = False Then
        Square(PicNum).BackColor = &H955800
        Square(PicNum).Picture = LoadPicture()
        Square(PicNum).BorderStyle = 0
    End If
    Sleep 50
    PicNum = PicNum + 1
    DoEvents
Loop

If Custom = False And NetGame = False Then
    HardStagePass(StageNum) = True
    If HardStageTime(StageNum) < Time Then
        HardStageTime(StageNum) = Time
    ElseIf HardStageTime(StageNum) > Time Then
    End If

    Open file For Output As #1
    Write #1, HardStagePass(1), HardStageTime(1), HardStagePass(2), HardStageTime(2), HardStagePass(3), HardStageTime(3), HardStagePass(4), HardStageTime(4), HardStagePass(5), HardStageTime(5), HardStagePass(6), HardStageTime(6), HardStagePass(7), HardStageTime(7), HardStagePass(8), HardStageTime(8), HardStagePass(9), HardStageTime(9), HardStagePass(10), HardStageTime(10)
    Close #1
End If

lblExit.Visible = True

End Sub

Public Sub CalcTrue()
Dim PicNum As Integer
NumTrue = 0
PicNum = 1

Do While PicNum < 101
    If Pic(PicNum) = True Then NumTrue = NumTrue + 1
    PicNum = PicNum + 1
Loop

End Sub

Private Sub lblExit_Click()

If Custom = True Then
    Custom = False
    frmMain.Show
    Unload Me
ElseIf NetGame = True Then
    NetGame = False
    frmNetLevels.Show
    Kill Stage
    Unload Me
Else
    Call HardRefresh
    frm10x10select.Show
    Unload Me
End If

End Sub

Private Sub Square_Click(Index As Integer)

If Pic(Index) = True Then
    Square(Index).Picture = LoadPicture(App.Path & "\true.gif")
    StatusBar1.Panels(1).Text = ""
    IsTrue = IsTrue + 1
    If IsTrue = NumTrue Then
        Timer1.Enabled = False
        StatusBar1.Panels(1).Text = "You Win! It's " & Description & "!"
        Frame1.Enabled = False
        Call Win
    End If
ElseIf Pic(Index) = False Then
    Square(Index).Picture = LoadPicture(App.Path & "\false.gif")
    Time = Time - 120
    StatusBar1.Panels(1).Text = "Miss! -120 seconds!"
    StatusBar1.Panels(2).Text = Time
    If Time < 0 Then
        Frame1.Enabled = False
        Timer1.Enabled = False
        StatusBar1.Panels(1).Text = "Time's Up! Game Over!"
        StatusBar1.Panels(2).Text = "0"
        lblExit.Visible = True
    End If
End If

Square(Index).Enabled = False

End Sub

Private Sub Timer1_Timer()

If Time = 0 Then
    Frame1.Enabled = False
    Timer1.Enabled = False
    StatusBar1.Panels(1).Text = "Time's Up! Game Over!"
    lblExit.Visible = True
Else
    Time = Time - 1
    StatusBar1.Panels(2).Text = Time
End If

End Sub

Private Sub StartHint()

Dim Random As Integer
Dim Random2 As Integer
Dim tmp As Integer
Dim Tmp2 As Integer
Dim Tmp3 As Integer
Dim Tmp4 As Integer
Randomize

Random = Int((10 * Rnd) + 1)
Random2 = Int((10 * Rnd) + 1)

If Random = 1 Then tmp = 1: Tmp3 = 10
If Random = 2 Then tmp = 11: Tmp3 = 20
If Random = 3 Then tmp = 21: Tmp3 = 30
If Random = 4 Then tmp = 31: Tmp3 = 40
If Random = 5 Then tmp = 41: Tmp3 = 50
If Random = 6 Then tmp = 51: Tmp3 = 60
If Random = 7 Then tmp = 61: Tmp3 = 70
If Random = 8 Then tmp = 71: Tmp3 = 80
If Random = 9 Then tmp = 81: Tmp3 = 90
If Random = 10 Then tmp = 91: Tmp3 = 100

If Random2 = 1 Then Tmp2 = 1: Tmp4 = 91
If Random2 = 2 Then Tmp2 = 2: Tmp4 = 92
If Random2 = 3 Then Tmp2 = 3: Tmp4 = 93
If Random2 = 4 Then Tmp2 = 4: Tmp4 = 94
If Random2 = 5 Then Tmp2 = 5: Tmp4 = 95
If Random2 = 6 Then Tmp2 = 6: Tmp4 = 96
If Random2 = 7 Then Tmp2 = 7: Tmp4 = 97
If Random2 = 8 Then Tmp2 = 8: Tmp4 = 98
If Random2 = 9 Then Tmp2 = 9: Tmp4 = 99
If Random2 = 10 Then Tmp2 = 10: Tmp4 = 100


Do While tmp <= Tmp3
    If Pic(tmp) = True And Square(tmp).Enabled = True Then
        Square(tmp).Picture = LoadPicture(App.Path & "\true.gif")
        Square(tmp).Enabled = False
        IsTrue = IsTrue + 1
    ElseIf Pic(tmp) = False Then
        Square(tmp).Picture = LoadPicture(App.Path & "\false.gif")
        Square(tmp).Enabled = False
    End If
tmp = tmp + 1
Loop

Do While Tmp2 <= Tmp4
    If Pic(Tmp2) = True And Square(Tmp2).Enabled = True Then
        Square(Tmp2).Picture = LoadPicture(App.Path & "\true.gif")
        Square(Tmp2).Enabled = False
        IsTrue = IsTrue + 1
    ElseIf Pic(Tmp2) = False Then
        Square(Tmp2).Picture = LoadPicture(App.Path & "\false.gif")
        Square(Tmp2).Enabled = False
    End If
Tmp2 = Tmp2 + 10
Loop

End Sub

