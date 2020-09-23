VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFF80&
   Caption         =   "Sudoku Solver Designed by M.M.Kahlil"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   Begin VB.CommandButton mailme 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mail Me"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   101
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ListBox List1 
      Height          =   1770
      Left            =   0
      TabIndex        =   100
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   7560
      TabIndex        =   96
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton savepuzzle 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Save"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ListBox samplelist 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00800000&
      Height          =   2310
      ItemData        =   "Form1.frx":000C
      Left            =   9720
      List            =   "Form1.frx":001C
      TabIndex        =   94
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton loadpuzzle 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Load Samples"
      Height          =   735
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton clearnum 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Clear"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton slovenow 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Solve now"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton newslove 
      BackColor       =   &H00FFC0C0&
      Caption         =   "New"
      Height          =   615
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame numfrm 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7365
      Left            =   2040
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
      Begin VB.Frame Frame7 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7215
         Left            =   0
         TabIndex        =   88
         Top             =   360
         Width           =   255
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   87
         Top             =   2520
         Width           =   7095
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   6840
         TabIndex        =   86
         Top             =   360
         Width           =   255
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   240
         TabIndex        =   85
         Top             =   7080
         Width           =   7095
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   4800
         Width           =   7095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   4560
         TabIndex        =   83
         Top             =   480
         Width           =   255
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   2280
         TabIndex        =   82
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   80
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   79
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   78
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   77
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   76
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   75
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   74
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   73
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   72
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   6480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   71
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   70
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   69
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   68
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   67
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   66
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   65
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   64
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   63
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   5760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   62
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   61
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   60
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   59
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   58
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   57
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   56
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   55
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   54
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   5040
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   53
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   52
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   51
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   50
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   49
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   48
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   47
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   46
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   45
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   44
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   43
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   42
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   41
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   40
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   39
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   38
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   37
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   36
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   35
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   34
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   33
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   32
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   31
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   30
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   29
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   28
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   27
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2760
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   26
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   25
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   24
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   23
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   22
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   21
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   20
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   19
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   18
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1920
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   17
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   16
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   15
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   14
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   13
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   12
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   11
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   10
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   9
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1200
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   8
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   7
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   6
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   5
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   4
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   3
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   2
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   1
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   650
      End
      Begin VB.CommandButton cell 
         BackColor       =   &H00C0FFFF&
         Height          =   650
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   650
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "This application degined by M.M.Khalil.    Email: khlilism2000@yahoo.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   97
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Label waitlbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   735
      Left            =   120
      TabIndex        =   99
      Top             =   6000
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label notes 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form1.frx":002C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3735
      Left            =   9720
      TabIndex        =   98
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Menu nums 
      Caption         =   "select  number"
      Visible         =   0   'False
      Begin VB.Menu selectnum 
         Caption         =   "1"
         Index           =   0
      End
      Begin VB.Menu selectnum 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu selectnum 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu selectnum 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu selectnum 
         Caption         =   "5"
         Index           =   4
      End
      Begin VB.Menu selectnum 
         Caption         =   "6"
         Index           =   5
      End
      Begin VB.Menu selectnum 
         Caption         =   "7"
         Index           =   6
      End
      Begin VB.Menu selectnum 
         Caption         =   "8"
         Index           =   7
      End
      Begin VB.Menu selectnum 
         Caption         =   "9"
         Index           =   8
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim unsolved(1 To 9, 1 To 9), solved(1 To 9, 1 To 9) As Byte
Dim dataarray(80) As String
Dim srtagn As Boolean
Dim inx, value1, s As Integer
Private Sub cell_Click(Index As Integer)
srtagn = True
inx = Index
ypos = (Index Mod 9) + 1: xpos = (Index \ 9) + 1
cell(Index).Caption = (Val(cell(Index).Caption) Mod 9) + 1
unsolved(xpos, ypos) = cell(Index).Caption
End Sub
Private Sub cell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ypos = (Index Mod 9) + 1: xpos = (Index \ 9) + 1
xstart = 3 * ((xpos - 1) \ 3) + 1: ystart = 3 * ((ypos - 1) \ 3) + 1
Label1 = "x=" & xpos & " y= " & ypos & "     xbox= " & xstart & " ybox= " & ystart
Label2 = dataarray(Index)
Label3 = unsolved(xpos, ypos)
End Sub
Private Sub cell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
inx = Index
PopupMenu nums, , X + cell(Index).Left + numfrm.Left, Y + cell(Index).Top + numfrm.Top
End If
End Sub

Private Sub mailme_Click()
On Error Resume Next
Shell "start mailto:khlilism2000@yahoo.com", vbNormalFocus

End Sub

Private Sub selectnum_Click(Index As Integer)
srtagn = True
cell(inx).Caption = selectnum(Index).Caption
ypos = (inx Mod 9) + 1: xpos = (inx \ 9) + 1
unsolved(xpos, ypos) = Val(cell(inx).Caption)
getmissingnum
End Sub
Private Sub clearnum_Click()
If inx >= 0 Then
cell(inx).Caption = ""
ypos = (inx Mod 9) + 1: xpos = (inx \ 9) + 1
unsolved(xpos, ypos) = 0
End If
End Sub
Private Sub loadpuzzle_Click()
samplelist.Clear
File1.Pattern = "*.txt"
File1.Path = App.Path & "\puzzles"
File1.Refresh
For n = 0 To File1.ListCount - 1
File1.ListIndex = n
samplelist.AddItem " puzzle : " & Left(File1.FileName, Len(File1.FileName) - 4)
Next
samplelist.Visible = Not samplelist.Visible
If samplelist.Visible = False Then
notes.Top = samplelist.Top
Else
notes.Top = samplelist.Top + samplelist.Height + 100

End If
End Sub

Private Sub Form_Load()
notes.Top = samplelist.Top
samplelist.Clear
File1.Pattern = "*.txt"
File1.Path = App.Path
For n = 0 To File1.ListCount - 1
File1.ListIndex = n
samplelist.AddItem Left(File1.FileName, Len(File1.FileName) - 4)
Next
End Sub
Private Sub samplelist_Click()
Dim ipt As String
slovenow.Visible = True

getmissingnum
items = Right(samplelist.Text, Len(samplelist.Text) - 9)
File1.ListIndex = samplelist.ListIndex
newslove_Click
srtagn = True
Open File1.Path & "\" & File1.FileName For Input As #1
Input #1, ipt
For m = 1 To 81
poss = m - 1
ypos = (poss Mod 9) + 1: xpos = (poss \ 9) + 1
unsolved(xpos, ypos) = Val(Mid(ipt, m, 1))
solved(xpos, ypos) = Val(Mid(ipt, m, 1))
If Val(Mid(ipt, m, 1)) <> 0 Then cell(poss).Caption = Val(Mid(ipt, m, 1))
Next
Close #1
End Sub

Private Sub savepuzzle_Click()
Dim lines, words As String
hh:
namepuzzle = InputBox("Please enter puzzle Name. the file will saved in the same path of Sudoku solver.exe")
If namepuzzle = "" Then
res = MsgBox("Do want enter the name again", vbYesNo, "Sudoku Solver")
If res = vbYes Then GoTo hh:
GoTo jj:
End If
Open App.Path & "\puzzles\" & namepuzzle & ".txt" For Output As #1
For m = 1 To 9
lines = ""
For n = 1 To 9
words = unsolved(m, n)
If unsolved(m, n) = 0 Then words = "0"

lines = lines & words
Next
Print #1, lines;
Next
Close #1
jj:

End Sub

Private Sub slovenow_Click()
slovenow.Visible = False
waitlbl.Visible = True
DoEvents

If srtagn = True Then
test1
test3
test4
test5
If testarray = False Then test2
End If
waitlbl.Visible = False
slovenow.Visible = True
End Sub
Private Sub Form_Resize()
numfrm.Left = (Me.Width - numfrm.Width) / 2
numfrm.Top = (Me.Height - numfrm.Height) / 2
End Sub
Private Sub newslove_Click()
s = 1: srtagn = False
For m = 0 To 80
cell(m).Caption = ""
dataarray(m) = ""
cell(m).BackColor = &HC0FFFF
Next
For m = 1 To 9
For n = 1 To 9
unsolved(m, n) = 0
solved(m, n) = 0
Next
Next
End Sub

Private Function getmissingnum()
Dim txt As String
For j = 0 To 80
dataarray(j) = ""
Next j
For i = 0 To 80
txt = ""
ypos = (i Mod 9) + 1: xpos = (i \ 9) + 1
xstart = 3 * ((xpos - 1) \ 3) + 1: ystart = 3 * ((ypos - 1) \ 3) + 1
If unsolved(xpos, ypos) <> 0 Then dataarray(i) = "": GoTo kh:
datas = "123456789"
'remove h,v
For xy = 1 To 9
If unsolved(xy, ypos) <> 0 Then
txt = txt & unsolved(xy, ypos)
End If
If unsolved(xpos, xy) <> 0 Then
txt = txt & unsolved(xpos, xy)
End If
Next xy
'remove box
For xx = xstart To xstart + 2
For yy = ystart To ystart + 2
If unsolved(xx, yy) <> 0 Then
txt = txt & unsolved(xx, yy)
End If
Next yy
Next xx
For c = 1 To 9
If InStr(txt, c) = 0 Then dataarray(i) = dataarray(i) & c
Next c
kh:
Next i
End Function
Private Function testarray() As Boolean
testarray = True
str1 = ""
For xx1 = 1 To 9
str1 = ""
For yy1 = 1 To 9
str1 = str1 & unsolved(xx1, yy1)
Next
For m = 1 To 9
If InStr(str1, m) = 0 Then testarray = False: GoTo kk:
Next
Next

For xx1 = 1 To 9
str1 = ""
For yy1 = 1 To 9
str1 = str1 & unsolved(yy1, xx1)
Next
For m = 1 To 9
If InStr(str1, m) = 0 Then testarray = False: GoTo kk:
Next
Next

For xx1 = 1 To 9 Step 3
For yy1 = 1 To 9 Step 3
For i = xx1 To xx1 + 2
For j = yy1 To yy1 + 2
str1 = str1 & unsolved(i, j)
Next
Next
For m = 1 To 9
If InStr(str1, m) = 0 Then testarray = False: GoTo kk:
Next
Next
Next
testarray = True
For i1 = 1 To 9
For j1 = 1 To 9
poss = 9 * (i1 - 1) + j1 - 1
cell(poss).Caption = unsolved(i1, j1)
Next j1
Next i1
kk:
End Function

Private Function test1()
kk:
getmissingnum
For m = 0 To 80
If Len(dataarray(m)) = 1 Then
ypos = (m Mod 9) + 1: xpos = (m \ 9) + 1
cell(m).BackColor = &HFFFF80
unsolved(xpos, ypos) = Val(dataarray(m))

GoTo kk:
End If
Next

End Function
Private Function test3()
Dim txt As String
Dim a, b As Integer
jj:
For m = 1 To 9 Step 3
For n = 1 To 9 Step 3
txt = ""
For i = m To m + 2
For j = n To n + 2
poss = 9 * (i - 1) + j - 1
If dataarray(poss) <> "" Then txt = txt & dataarray(poss)
Next
Next
a = 1000: b = 1000
For xx = 1 To 9
  a = InStr(txt, xx): b = InStrRev(txt, xx)
  If a = b And a <> 1000 Then
    For i1 = m To m + 2
    For j1 = n To n + 2
    poss = 9 * (i1 - 1) + j1 - 1
      If InStr(dataarray(poss), xx) <> 0 Then
        unsolved(i1, j1) = xx
        cell(poss).BackColor = &HFFFF80
        test1
        GoTo jj:
      End If
    Next
    Next
  End If
Next


Next
Next


End Function
Private Function test4()
Dim txt As String
Dim a, b As Integer
jj:
getmissingnum
For m = 1 To 9
txt = ""
For n = 1 To 9
poss = 9 * (m - 1) + n - 1
If dataarray(poss) <> "" Then txt = txt & dataarray(poss)
Next n
a = 1000: b = 1000
For xx = 1 To 9
a = InStr(txt, xx): b = InStrRev(txt, xx)
If a = b And a <> 1000 Then
For n1 = 1 To 9
poss = 9 * (m - 1) + n1 - 1
If InStr(dataarray(poss), xx) <> 0 Then
unsolved(m, n1) = xx
cell(poss).BackColor = &HFFFF80
test3
GoTo jj:
End If
Next n1
End If
Next xx
Next m
End Function
Private Function test5()
Dim txt As String
Dim a, b As Integer
jj:
getmissingnum
For m = 1 To 9
txt = ""
For n = 1 To 9
poss = 9 * (n - 1) + m - 1
If dataarray(poss) <> "" Then txt = txt & dataarray(poss)
Next n
a = 1000: b = 1000
For xx = 1 To 9
a = InStr(txt, xx): b = InStrRev(txt, xx)
If a = b And a <> 1000 Then
For n1 = 1 To 9
poss = 9 * (n1 - 1) + m - 1
If InStr(dataarray(poss), xx) <> 0 Then
unsolved(n1, m) = xx
cell(poss).BackColor = &HFFFF80
cell(poss).Caption = xx
test3
GoTo jj:
End If
Next n1
End If
Next xx
Next m

End Function

Private Function test2()
Dim tray(1 To 2) As String
Dim xxtray(1 To 2) As Byte
Dim yytray(1 To 2) As Byte
Dim lenghts(1 To 2) As Byte
Dim num1, num2 As Byte
Dim txt, txt1, txt2 As String
Dim splt() As String

For i3 = 1 To 9
For j3 = 1 To 9
solved(i3, j3) = unsolved(i3, j3)
Next
Next

gg:

List1.Clear
getmissingnum
For m = 0 To 80

If Len(dataarray(m)) <> 0 Then
ypos = (m Mod 9) + 1: xpos = (m \ 9) + 1
List1.AddItem Trim(xpos & ypos & "=" & dataarray(m)) & "=" & Len(dataarray(m))
End If

Next m

If List1.ListCount > 1 Then
k = 1

Do Until k = 3
xx = Int(List1.ListCount * Rnd)
List1.ListIndex = xx
txt = List1.Text
splt = Split(txt, "=")
tray(k) = splt(1)
xxtray(k) = Val(Mid(splt(0), 1, 1))
yytray(k) = Val(Mid(splt(0), 2, 1))
lenghts(k) = Val(splt(2))
List1.RemoveItem xx
List1.Refresh
k = k + 1
Loop

End If

txt1 = tray(1)
txt2 = tray(2)

s1 = lenghts(1)
s2 = lenghts(2)


For m = 1 To s1
num1 = Val(Mid(txt1, m, 1))

For n = 1 To s2
num2 = Val(Mid(txt2, n, 1))

unsolved(xxtray(1), yytray(1)) = num1
If testnum(num2, xxtray(2), yytray(2)) = False Then GoTo nextnum:

unsolved(xxtray(2), yytray(2)) = num2
test1
test3
test4
test5

test1
test3
test4
test5

If testarray = False Then
For i3 = 1 To 9
For j3 = 1 To 9
unsolved(i3, j3) = solved(i3, j3)
Next
Next
Else
GoTo hh:
End If
nextnum:
Next
Next
If testarray = False Then
For i3 = 1 To 9
For j3 = 1 To 9
unsolved(i3, j3) = solved(i3, j3)
Next
Next
GoTo gg:
End If
hh:
End Function
Private Function testnum(numval As Byte, xval As Byte, yval As Byte) As Boolean
Dim txt1, txt2, txt3 As String
txt1 = "": txt2 = "": txt3 = ""
testnum = True
For m = 1 To 9
If unsolved(xval, m) <> 0 Then txt1 = txt1 & unsolved(xval, m)
If unsolved(m, yval) <> 0 Then txt2 = txt2 & unsolved(m, yval)
Next

a = InStr(txt1, numval): b = InStrRev(txt1, numval)
If a <> b Then
If a <> 0 And b <> 0 Then testnum = False: GoTo kk:
End If

a = InStr(txt2, numval): b = InStrRev(txt2, numval)
If a <> b Then
If a <> 0 And b <> 0 Then testnum = False: GoTo kk:
End If

xstart = 3 * ((xval - 1) \ 3) + 1: ystart = 3 * ((yval - 1) \ 3) + 1
For m = xstart To xstart + 2
For n = ystart To ystart + 2
If unsolved(m, n) <> 0 Then txt3 = txt3 & unsolved(m, n)
Next
Next
a = InStr(txt3, numval): b = InStrRev(txt3, numval)
If a <> b Then
If a <> 0 And b <> 0 Then testnum = False: GoTo kk:
End If
kk:
End Function


