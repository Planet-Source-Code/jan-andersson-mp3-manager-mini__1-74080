VERSION 5.00
Begin VB.Form frmMX3 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12630
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MX3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   628
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   842
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1035
      Index           =   7
      Left            =   2940
      ScaleHeight     =   1035
      ScaleWidth      =   5055
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   3720
      Width           =   5055
      Visible         =   0   'False
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   17
         Left            =   960
         TabIndex        =   114
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   21
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   22
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   116
         Top             =   480
         Width           =   555
      End
      Begin VB.Image img 
         Height          =   720
         Index           =   10
         Left            =   60
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   29
         Left            =   960
         TabIndex        =   117
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3495
      Index           =   4
      Left            =   -5280
      ScaleHeight     =   3495
      ScaleWidth      =   8595
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   1140
      Width           =   8595
      Visible         =   0   'False
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   7560
         Picture         =   "MX3.frx":0ECA
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   2460
         Width           =   555
      End
      Begin VB.ListBox lst 
         Height          =   2400
         Index           =   0
         IntegralHeight  =   0   'False
         ItemData        =   "MX3.frx":131D
         Left            =   1320
         List            =   "MX3.frx":131F
         TabIndex        =   20
         Top             =   4320
         Width           =   7515
         Visible         =   0   'False
      End
      Begin VB.CheckBox chk 
         Height          =   195
         Index           =   0
         Left            =   6420
         Picture         =   "MX3.frx":1321
         TabIndex        =   21
         Top             =   3120
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   15
         Left            =   4140
         MaxLength       =   4
         TabIndex        =   83
         Top             =   3060
         Width           =   555
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   14
         Left            =   960
         MaxLength       =   30
         TabIndex        =   81
         Top             =   2340
         Width           =   2535
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   13
         Left            =   960
         MaxLength       =   30
         TabIndex        =   79
         Top             =   3060
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Height          =   195
         Index           =   1
         Left            =   4800
         Picture         =   "MX3.frx":1F63
         TabIndex        =   78
         Top             =   3120
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   4
         Left            =   4800
         TabIndex        =   76
         Top             =   2700
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   12
         Left            =   960
         MaxLength       =   30
         TabIndex        =   74
         Top             =   2700
         Width           =   2535
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   3
         Left            =   4800
         Sorted          =   -1  'True
         TabIndex        =   72
         Top             =   2340
         Width           =   2715
      End
      Begin VB.ListBox lst 
         Height          =   1185
         Index           =   1
         ItemData        =   "MX3.frx":2BA5
         Left            =   0
         List            =   "MX3.frx":2BA7
         Style           =   1  'Checkbox
         TabIndex        =   23
         Top             =   1020
         Width           =   3555
         Visible         =   0   'False
      End
      Begin VB.ListBox lst 
         Height          =   1185
         Index           =   5
         ItemData        =   "MX3.frx":2BA9
         Left            =   3240
         List            =   "MX3.frx":2BAB
         Style           =   1  'Checkbox
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1020
         Width           =   3795
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   7560
         Picture         =   "MX3.frx":2BAD
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   2940
         Width           =   555
      End
      Begin VB.TextBox txt 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Text            =   "MX3.frx":2FFA
         Top             =   60
         Width           =   6975
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   7560
         Picture         =   "MX3.frx":3000
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   1980
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   20
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   1500
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   19
         Left            =   7560
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   1020
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   18
         Left            =   7080
         Picture         =   "MX3.frx":3428
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   1620
         Width           =   375
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   17
         Left            =   7080
         Picture         =   "MX3.frx":380F
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1020
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   38
         Left            =   3600
         TabIndex        =   84
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   36
         Left            =   0
         TabIndex        =   82
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   34
         Left            =   0
         TabIndex        =   80
         Top             =   3120
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   32
         Left            =   3600
         TabIndex        =   77
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   28
         Left            =   0
         TabIndex        =   75
         Top             =   2760
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   26
         Left            =   3600
         TabIndex        =   73
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   24
         Left            =   3600
         TabIndex        =   70
         Top             =   780
         Width           =   105
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   22
         Left            =   60
         TabIndex        =   69
         Top             =   780
         Width           =   105
      End
      Begin VB.Image img 
         Height          =   720
         Index           =   0
         Left            =   7260
         Picture         =   "MX3.frx":3BF2
         Top             =   120
         Width           =   720
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   3195
      Index           =   1
      Left            =   -900
      ScaleHeight     =   3195
      ScaleWidth      =   8595
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "-1"
      Top             =   5400
      Width           =   8595
      Visible         =   0   'False
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   7620
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   6300
         Picture         =   "MX3.frx":42E5
         Style           =   1  'Graphical
         TabIndex        =   103
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   4740
         Picture         =   "MX3.frx":46CA
         Style           =   1  'Graphical
         TabIndex        =   102
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   0
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   0
         Width           =   555
      End
      Begin VB.TextBox txt 
         Height          =   1155
         Index           =   11
         Left            =   4020
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Top             =   420
         Width           =   4155
      End
      Begin VB.ListBox lst 
         Height          =   1230
         Index           =   4
         ItemData        =   "MX3.frx":4AB0
         Left            =   0
         List            =   "MX3.frx":4AB2
         TabIndex        =   11
         Top             =   1680
         Width           =   8175
      End
      Begin VB.ListBox lst 
         Height          =   1185
         Index           =   3
         ItemData        =   "MX3.frx":4AB4
         Left            =   0
         List            =   "MX3.frx":4AB6
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   420
         Width           =   1815
      End
      Begin VB.ListBox lst 
         Height          =   1185
         Index           =   2
         ItemData        =   "MX3.frx":4AB8
         Left            =   1980
         List            =   "MX3.frx":4ABA
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   420
         Width           =   1875
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   4
         Left            =   660
         TabIndex        =   6
         Top             =   0
         Width           =   3195
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "x"
         ForeColor       =   &H8000000D&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   12
         Top             =   2940
         Width           =   8235
      End
      Begin VB.Label lbl 
         Caption         =   "Words:"
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   10
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "0/0"
         Height          =   195
         Index           =   3
         Left            =   5280
         TabIndex        =   7
         Top             =   105
         Width           =   1035
      End
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   4140
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Text            =   "MX3.frx":4ABC
      Top             =   3180
      Width           =   1695
      Visible         =   0   'False
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   2475
      Index           =   6
      Left            =   3540
      ScaleHeight     =   2475
      ScaleWidth      =   5595
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   420
      Width           =   5595
      Visible         =   0   'False
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   5
         Left            =   3000
         TabIndex        =   43
         Top             =   60
         Width           =   1095
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   30
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   4140
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   30
         Width           =   555
      End
      Begin VB.Image img 
         Height          =   720
         Index           =   2
         Left            =   4740
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl 
         Caption         =   "((Length Milliseconds / 1000) *  Bitrate * 1024) / 8 =  Size"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   720
         TabIndex        =   57
         Top             =   1200
         Width           =   4845
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   50
         Left            =   720
         TabIndex        =   66
         Top             =   0
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   52
         Left            =   720
         TabIndex        =   65
         Top             =   900
         Width           =   45
      End
      Begin VB.Label lbl 
         Caption         =   "Length:"
         Height          =   195
         Index           =   21
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Size:"
         Height          =   195
         Index           =   20
         Left            =   0
         TabIndex        =   63
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Formula:"
         Height          =   195
         Index           =   19
         Left            =   0
         TabIndex        =   62
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "Bitrate:"
         Height          =   195
         Index           =   18
         Left            =   2400
         TabIndex        =   61
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl 
         Caption         =   "File:"
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   60
         Top             =   600
         Width           =   435
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   53
         Left            =   720
         TabIndex        =   59
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Index           =   51
         Left            =   720
         TabIndex        =   58
         Top             =   300
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   40
         Left            =   840
         TabIndex        =   56
         Top             =   1860
         Width           =   45
      End
      Begin VB.Label lbl 
         Caption         =   "Bitrate:"
         Height          =   195
         Index           =   23
         Left            =   0
         TabIndex        =   55
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   42
         Left            =   4680
         TabIndex        =   54
         Top             =   1860
         Width           =   45
      End
      Begin VB.Label lbl 
         Caption         =   "Bits:"
         Height          =   195
         Index           =   25
         Left            =   3840
         TabIndex        =   53
         Top             =   1860
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   41
         Left            =   2940
         TabIndex        =   52
         Top             =   1860
         Width           =   45
      End
      Begin VB.Label lbl 
         Caption         =   "Frequency:"
         Height          =   195
         Index           =   27
         Left            =   1740
         TabIndex        =   51
         Top             =   1860
         Width           =   1095
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   45
         Left            =   840
         TabIndex        =   50
         Top             =   2140
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Channels:"
         Height          =   195
         Index           =   31
         Left            =   0
         TabIndex        =   49
         Top             =   2140
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   48
         Left            =   4680
         TabIndex        =   48
         Top             =   2140
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "File size:"
         Height          =   195
         Index           =   37
         Left            =   3840
         TabIndex        =   47
         Top             =   2140
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   47
         Left            =   2940
         TabIndex        =   46
         Top             =   2140
         Width           =   45
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Play time:"
         Height          =   195
         Index           =   39
         Left            =   1740
         TabIndex        =   45
         Top             =   2140
         Width           =   705
      End
      Begin VB.Label lbl 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   30
         Left            =   0
         TabIndex        =   44
         Top             =   1560
         Width           =   4275
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   4095
      Index           =   5
      Left            =   8100
      ScaleHeight     =   4095
      ScaleWidth      =   4215
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4215
      Visible         =   0   'False
      Begin VB.ListBox lst 
         Height          =   2025
         Index           =   6
         IntegralHeight  =   0   'False
         ItemData        =   "MX3.frx":4AC6
         Left            =   120
         List            =   "MX3.frx":4AC8
         Style           =   1  'Checkbox
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   3900
         Width           =   3675
         Visible         =   0   'False
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   16
         Left            =   1020
         TabIndex        =   112
         Top             =   2100
         Width           =   2655
         Visible         =   0   'False
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   4
         Left            =   0
         Picture         =   "MX3.frx":4ACA
         TabIndex        =   110
         Top             =   2520
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   3
         Left            =   3480
         Picture         =   "MX3.frx":570C
         TabIndex        =   108
         Top             =   2520
         Value           =   1  'Checked
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Height          =   195
         Index           =   2
         Left            =   3480
         Picture         =   "MX3.frx":634E
         TabIndex        =   107
         Top             =   2520
         Value           =   1  'Checked
         Width           =   195
         Visible         =   0   'False
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   10
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   38
         Top             =   2100
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   9
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   37
         Top             =   1260
         Width           =   555
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   8
         Left            =   1020
         MaxLength       =   4
         TabIndex        =   36
         Top             =   1260
         Width           =   555
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   7
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   29
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   6
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   28
         Top             =   420
         Width           =   2655
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   5
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   26
         Top             =   0
         Width           =   2655
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   2
         Left            =   1020
         Sorted          =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   1720
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   2420
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   3240
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   13
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   54
         Left            =   300
         TabIndex        =   111
         Top             =   2520
         Width           =   90
         Visible         =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "x"
         Height          =   195
         Index           =   49
         Left            =   3300
         TabIndex        =   109
         Top             =   2520
         Width           =   90
         Visible         =   0   'False
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   9
         Left            =   2820
         Picture         =   "MX3.frx":6F90
         Top             =   1260
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   8
         Left            =   2520
         Picture         =   "MX3.frx":7401
         Top             =   1260
         Width           =   315
         Visible         =   0   'False
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   7
         Left            =   2160
         Picture         =   "MX3.frx":78C1
         Top             =   1260
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "x"
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   41
         Top             =   2520
         Width           =   3675
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   14
         Left            =   2400
         TabIndex        =   40
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "x"
         ForeColor       =   &H8000000D&
         Height          =   195
         Index           =   17
         Left            =   60
         TabIndex        =   39
         Top             =   2820
         Width           =   3675
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   5
         Left            =   1860
         Picture         =   "MX3.frx":7CC5
         Top             =   1260
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   4
         Left            =   1620
         Picture         =   "MX3.frx":8093
         Top             =   1260
         Width           =   300
         Visible         =   0   'False
      End
      Begin VB.Image img 
         Height          =   480
         Index           =   6
         Left            =   0
         Top             =   3000
         Width           =   480
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   15
         Left            =   0
         TabIndex        =   35
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lbl 
         Caption         =   "Track:"
         Height          =   14
         Index           =   1299
         Left            =   2460
         TabIndex        =   34
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   33
         Top             =   1740
         Width           =   855
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   32
         Top             =   60
         Width           =   855
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   13
         Left            =   0
         TabIndex        =   30
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   12
         Left            =   0
         TabIndex        =   27
         Top             =   900
         Width           =   855
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   1875
      Index           =   3
      Left            =   9720
      ScaleHeight     =   1875
      ScaleWidth      =   2715
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   960
      Width           =   2715
      Visible         =   0   'False
      Begin VB.TextBox txt 
         Alignment       =   1  'Right Justify
         Height          =   315
         Index           =   1
         Left            =   60
         MaxLength       =   4
         TabIndex        =   18
         Top             =   1140
         Width           =   555
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   9
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   1080
         Width           =   555
      End
      Begin VB.Image img 
         Height          =   300
         Index           =   3
         Left            =   960
         Picture         =   "MX3.frx":8471
         Top             =   1140
         Width           =   300
      End
      Begin VB.Label lbl 
         Height          =   975
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   795
      Index           =   2
      Left            =   9180
      ScaleHeight     =   795
      ScaleWidth      =   2955
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   60
      Width           =   2955
      Visible         =   0   'False
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         Left            =   780
         TabIndex        =   15
         Top             =   60
         Width           =   615
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1740
         Picture         =   "MX3.frx":8858
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   30
         Width           =   555
      End
      Begin VB.Image img 
         Height          =   240
         Index           =   1
         Left            =   1440
         Picture         =   "MX3.frx":8C3F
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lbl 
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   14
         Top             =   120
         Width           =   675
      End
   End
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   555
      Index           =   0
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5295
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4800
      Width           =   5295
      Visible         =   0   'False
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   2820
         MaxLength       =   8
         TabIndex        =   3
         Top             =   30
         Width           =   915
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton cmd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3960
         Picture         =   "MX3.frx":904D
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   0
         Width           =   555
      End
      Begin VB.Label lbl 
         Caption         =   "CDID:"
         Height          =   195
         Index           =   2
         Left            =   2280
         TabIndex        =   4
         Top             =   90
         Width           =   495
      End
      Begin VB.Label lbl 
         Caption         =   "Category:"
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmMX3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*************************************************************************************************
'*  Copyright© Pappsegull Sweden, http://freetranslator.webs.com <pappsegull@yahoo.se>
'*
'*
'* FEATURES
'* --------
'* - MP3 encoding
'* - Create playlists using filter
'* - Batch edit MP3 tags using filter.
'* - CD audio ripping with auto tagging.
'* - Auto add MP3 tags, integrated with CDDB.
'* - + some more useful stuff;-)
'* - For more features Download my FREE MX3.dll at http://www.mediafire.com/?2rqfqid592a7c
'*
'* This software is provided "as-is," without any express or implied warranty.
'* In no event shall the author be held liable for any damages arising from the use of this software.
'* If you do not agree with these terms, do not use it!
'* Use of the program implicitly means you have agreed to these terms.
'*
'* Permission is granted to anyone to use this software for any purpose,
'* including commercial use, and to alter and redistribute it, provided that
'* the following conditions are met:
'*
'* CONDITIONS
'* ----------
'*   1. All redistribution of source code files must retain all copyright
'*      notices that are currently in place, and this list of conditions without
'*      any modification.
'*   2. All redistribution in binary form must retain all occurrences of the
'*      above copyright notice and web site addresses that are currently in
'*      place (for example, in the About boxes).
'*   3. Modified versions in source or binary form must be plainly marked as
'*      such, and must not be misrepresented as being the original software.
'*
'*************************************************************************************************

Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type TrackQ
    Time As String * 5
    Title As String
End Type
Private Type AlbumsQ
    DiscID As String * 8
    Category As String
    Year As String * 4
    TimeTot As String * 5
    Artist As String
    Album As String
    Track() As TrackQ
    nTracks As Long
End Type
Private Type PagesQ
    nAlbums As Long
    Read As Boolean
    Albums() As AlbumsQ
End Type
Private Type PageQ
    ToFile As String
    nPages As Long
    nResults As Long
    CurPage As Long
    Page() As PagesQ
End Type
Public FormType As mx3FormTypes, TimeToUnload As Boolean
Private Const c_W = 8400, c_H = 4125, c_Green = &H8000&, c_Red = &H80&, c_SL = "/", c_T1 = ":5px;"">", c_T2 = ".</td><td style=""width:100%"">", c_BR = "<br>", c_TD = "</td>"
Private sFile$, sInfo$, sFolder$, SzFile&, sEncFileInfo$, bNewFile As Boolean, bNotNow As Boolean, _
  CDDBQ As PageQ, Albums() As AlbumsQ, sValue$, lLeft&, lTop&, sAlias$, sDirtyTag$, sDirtyTag2$, _
  sFiles$(), bBatch As Boolean, oUserform As Object, FocusCtrl As Control, bFocus As Boolean

'// Click on buttons
Private Sub cmd_Click(Index As Integer)
Dim n&, l&, c&, s$, m$, v$(), b As Boolean
    
    Select Case Index
        Case 0, 22: sValue$ = "": Unload Me    'Canceled
        Case 1                              'Click on Tag button
            If FormType = frmTagRename Then
                m_Settings.DefComment = txt(13): m_Settings.MP3FileFormat = cmb(4).ListIndex
                m_Settings.CreatePlaylist = chk(1): sValue$ = "": bNotNow = True
                For n& = 0 To lst(1).ListCount - 1
                    InfoCDDB.tr(n& + 1).PathMP3 = "": lst(5).ListIndex = n&: s$ = lst(5).Text
                    InfoCDDB.tr(n& + 1).ListTime = Left$(s$, 9)
                    InfoCDDB.tr(n& + 1).ListName = Right$(s$, Len(s$) - 9)
                    If lst(1).Selected(n&) And lst(5).ListCount - 1 >= n& Then
                        lst(1).ListIndex = n&
                        If lst(5).Text <> "  " & IL(47) Then 'Missing
                            l& = InfoCDDB.tr(n& + 1).LengthMs - lst(5).ItemData(n&)
                            If l& < 0 Then l& = l& * -1
                            If l& > m_Settings.MP3MaxDiffMs Then
                                m$ = m$ & Left$(lst(1).Text, 12) & " <--> " & _
                                  " {" & l& / 1000 & " Sec}" & vbTab & s$ & vbLf
                                c& = c& + 1: If l& / 1000 > 2 Then b = True
                            End If
                            s$ = Right$(s$, Len(s$) - 9)
                            InfoCDDB.tr(n& + 1).TrackNoTmp = n& + 1
                            InfoCDDB.tr(n& + 1).PathMP3 = sFolder$ & s$
                        Else: lst(1).Selected(n&) = False: End If
                    End If
                Next
                bNotNow = False: If LenB(m$) Then If M3.MsgBoxW(m$ & vbLf & _
                  IL(109) & " " & LCase(IIf(c& > 1, IL(21), IL(54))) & "." & _
                  vbLf & IL(84), IIf(b, vbDefaultButton2 + vbYesNo + vbExclamation, _
                  vbDefaultButton2 + vbYesNo + vbQuestion)) = vbNo Then _
                  sValue$ = "": Exit Sub 'Diff in time between tracks
                If LenB(sAlias$) Then cmd_Click 20 'Stop play
                If M3.TagRenameApply Then 'Update listbox if file names have change
                    bNotNow = True
                    For n& = 0 To lst(1).ListCount - 1
                        s$ = InfoCDDB.tr(n& + 1).ListTime & InfoCDDB.tr(n& + 1).ListName
                        lst(5).ListIndex = n&: l& = lst(5).ItemData(n&)
                        lst(5).RemoveItem n&: lst(5).AddItem s$, n&: lst(5).ItemData(n&) = l&
                    Next
                    bNotNow = False
                End If
            ElseIf FormType = frmShowCDDBMatches Then 'Inexcact matches from CDDB
                v$() = Split(lst(0).Text, vbTab): sTmpStr$ = Trim$(v(1)) & ";" & LCase(Trim$(v(2)))
                sTmpFolder$ = sFolder$: M3.TagRename [By CDID and category]
            End If
        Case 2: sValue$ = cmb(0).Text: Unload Me             'CD-Drive selected
        Case 3, 4, 16                                        'Show details of album
            s$ = IL(112) & "... " & IL(114) 'Wait... requesting details from CDDB
            If Index = 3 Then 'By clicking the "Details" button
                If Not lst(0).Visible Then
                    M3.ShowTextForm Me, sInfo$, Caption: Exit Sub
                End If
                v$() = Split(lst(0).Text, vbTab): sValue$ = txt(3): txt(3) = s$: txt(3).Refresh
            ElseIf Index = 16 Then 'By clicking the "Details" button in free search
                v$() = Split(lst(4).Text, vbTab): sValue$ = lbl(4): lbl(4) = s$
                v$(2) = v$(1): v$(1) = Right$(v$(0), 8): v$(0) = "": txt(3).Refresh
            Else 'Index = 4: Search CDDB by category and CDID, by clicking the "Search" button
                sValue$ = txt(2) & ";" & LCase(ILOrg(cmb(1).ListIndex + 90))
                Caption = s$: DoEvents: Unload Me: Exit Sub
            End If
            MousePointer = 11
            If M3.CDDBQuery(, v$(0), , , Trim$(v$(1)), Trim$(LCase(v$(2))), , , s$) = 200 Then
               s$ = Replace(v$(0), "\n", vbNewLine)
               MousePointer = 0: txt(3) = sValue$: M3.ShowTextForm Me, s$, Caption
            Else
                MousePointer = 0: M3.ShowTextForm Me, s$, Caption
            End If
            If Index = 16 Then lbl(4) = sValue$: sValue$ = ""
        Case 5 'New search
            GoToPage
        Case 6 'Previous page
            CDDBQ.CurPage = CDDBQ.CurPage - 1: GoToPage CDDBQ.CurPage
        Case 7 'Next page
            CDDBQ.CurPage = CDDBQ.CurPage + 1: GoToPage CDDBQ.CurPage
        Case 8 'Tag and rename files from Free Text search
            'Call CDDBQuery to calculate time on tracks
            l& = lst(4).ListIndex + 1: sTmpStr$ = Albums(l&).DiscID & ";" & LCase(Albums(l&).Category)
            M3.TagRename [By CDID and category]: sTmpStr$ = ""
        Case 9 'Adjust max diff Ms
            sValue$ = txt(1): Unload Me
        Case 10 'Search from Tagger
            If FormType = frmTaggerForm Then
                M3.ShowFreeSearch Me, Chr(34) & txt(5) & Chr(34) & _
                  " " & Chr(34) & txt(6) & Chr(34) & _
                  " " & Chr(34) & txt(7) & Chr(34)
            Else 'Search MP3 files, in Create playlist or Tag batch edit
                If cmd(10).Picture = cmd(4).Picture Then
                    
                    lbl(17) = IL(112) & "... " & IL(115): ReDim sFiles$(0)
                    SearchTag.CheckTag = True: lst(6).Clear: cmd(10).Enabled = False
                    For l& = 0 To 5
                        SearchTag.Check(l&) = IIf(txt(l& + 5) = "", "?", txt(l& + 5))
                    Next
                    If cmb(2).ListIndex < 0 Then cmb(2).ListIndex = 0
                    SearchTag.Check(6) = cmb(2).ItemData(cmb(2).ListIndex)
                    sFiles$() = M3.DirW(sFolder$, , CBool(chk(2)), True)
                    If sFiles$(0) <> vbNullString Then 'MP3-Files found in folder(s)
                        n& = UBound(sFiles$()): cmd(11).Enabled = True: lst(6).Clear
                        For l& = 0 To n&
                            lst(6).AddItem M3.PathInfo(sFiles$(l&), encNameExt)
                        Next
                        lst(6).ListIndex = 0: cmd(10).Enabled = True: lst_Click 6: chk_Click 3
                    End If
                End If: FixForm 10: SearchTag.CheckTag = False
            End If
        Case 11, 14, 20 'Play/Stop play in File size calc. and in File tagger
            If FormType = frmPlayListCreator Or FormType = frmTagBatchEdit Then
                If cmd(11).Picture = img(4) Or cmd(11).Picture = img(5) Then
                Else
                    If FixForm(11) Then Exit Sub
                End If
            End If
            If LenB(sAlias$) Then   'Stop play
                If M3.MP3Stop(sAlias$) Then
                    cmd(Index).Picture = img(4): sAlias$ = vbNullString
                End If
            Else                    'Start play
                sAlias$ = "MX3" & M3.GetTicCount
                If M3.MP3Play(sFile$, sAlias$) Then
                    cmd(Index).Picture = img(5)
                End If
            End If
        Case 12 'Open MP3-File in Tagger
            If FormType = frmTaggerForm Then
                s$ = sFile$: sFile$ = "": ShowTagger , , True: If sFile$ = "" Then sFile$ = s$
            Else 'Select folder to search for MP3 files
                s$ = sFolder$
                sFolder$ = M3.FileFolder(encShowFolder, , IL(116))
                If LenB(sFolder$) Then
                    If Not M3.FolderExists(sFolder$) Then
                        M3.MsgBoxW IL(82) & sFolder$, 48
                        If LenB(s$) Then sFolder$ = s$ Else sFolder$ = ""
                        Exit Sub
                    End If
                Else
                    If LenB(s$) Then sFolder$ = s$: Exit Sub
                End If
                If LenB(sFolder$) Then
                    lst(6).Clear: If Right$(sFolder$, 1) <> "\" Then sFolder$ = sFolder$ & "\"
                End If: FixForm 12
            End If
        Case 13
            If FormType = frmTaggerForm Then 'Save tag to MP3 file in Tagger
                If LenB(sAlias$) Then cmd_Click 11 'Stop play
                If M3.MP3TagInfoLet(sFile$, txt(5), txt(6), txt(7), txt(8), _
                  cmb(2).Text, txt(9), txt(10), sAlias$) Then _
                    TagDirty sDirtyTag$: cmd(13).Enabled = False
            ElseIf FormType = frmTagBatchEdit Then 'Go to save tag batch
                If cmd(13).Picture = cmd(7).Picture Then 'Show Edit tag batch on last page
                Else: BatchEditDo: End If: FixForm 13
            ElseIf FormType = frmPlayListCreator Then: PlaylistDo: End If
        Case 15 'Select file, "File size calculator"
            For l& = 40 To 42: lbl(l&) = "": Next: lbl(45) = "": For l& = 47 To 48: lbl(l&) = "": Next
            If LenB(Tag) = 0 Then
                s$ = "WAVE " & IL(126) & " (*.wav)|*.wav|All Files (*.*)|*.*"
                s$ = M3.ShowOpen(s$, IL(125)): Refresh: DoEvents
                If LenB(s$) = 0 Then Exit Sub
            Else: s$ = Tag: Tag = "": End If
            If Not M3.FileExists(s$) Then Exit Sub
            sFile$ = s$: SzFile& = FileLen(s$): bNewFile = True
            sEncFileInfo$ = M3.PathInfo(s$, encNameExt)
            v$() = Split(" Kbps; kHz", ";"): s$ = ""
            With M3
                If .WAVHeaderRead(sFile$) Then 'Check if find RIFF in header
                    lbl(40) = .WAVHeaderInfo(wavKbps) & v$(0)
                    lbl(41) = .WAVHeaderInfo(wavFrequency) / 1000 & v$(1)
                    lbl(42) = .WAVHeaderInfo(wavBits): s$ = "WAVE-"
                    lbl(45) = .WAVHeaderInfo(wavChannels)
                    lbl(47) = .WAVHeaderInfo(wavPlaytime)
                    lbl(48) = .FormatKMG(.WAVHeaderInfo(wavFilesize), True, 2)
                End If
            End With
            Erase v$(): lbl(30) = s$ & IL(117) & ":": Calc
        Case 17 'Move item UP
            If lst(5).ListIndex = 0 Then Exit Sub
            s$ = lst(5).List(lst(5).ListIndex): l& = lst(5).ListIndex: n& = lst(5).ItemData(l&)
            lst(5).RemoveItem lst(5).ListIndex: lst(5).AddItem s$, l& - 1
            lst(5).Selected(l& - 1) = True: lst(5).ItemData(l& - 1) = n&
            If lst(1).ListCount - 1 >= l& - 1 Then
                lst(1).ListIndex = l& - 1: lst(1).Selected(l& - 1) = True
            End If
            lst_Click 5
        Case 18 'Move item DOWN
            If lst(5).ListCount - 1 = lst(5).ListIndex Then Exit Sub
            s$ = lst(5).List(lst(5).ListIndex): l& = lst(5).ListIndex: n& = lst(5).ItemData(l&)
            lst(5).RemoveItem lst(5).ListIndex: lst(5).AddItem s$, l& + 1
            lst(5).Selected(l& + 1) = True: lst(5).ItemData(l& + 1) = n&
            If lst(1).ListCount - 1 >= l& + 1 Then
                lst(1).ListIndex = l& + 1: lst(1).Selected(l& + 1) = True
            End If
            lst_Click 5
        Case 19 'Show Settings
            M3.ShowAdjustDiffMs Me
        Case 21 'Input form, Genre from number, Genre number from name
            s$ = txt(17)
            Select Case FormType
                Case frmInput
                    sValue$ = s$: Unload Me
                Case frmGenreFromNumber     'Get genre name from number
                    s$ = Val(s$): bNotNow = True: txt(17) = s$: bNotNow = False
                    If Val(s$) > 255 Or Val(s$) < 0 Then _
                      M3.MsgBoxW IL(169) & ": " & s$, 48: txt(17) = "": Exit Sub
                    l& = Val(s$): M3.MP3TagInfoGet encTagGenre, m$: If LenB(s$) = 0 Then Exit Sub
                    M3.MsgBoxW m$ & " " & s$ & ": " & M3.MP3GetTagGenreName((s$))
                Case frmGenreNumberFromName 'Get genre number from name
                    If s$ = vbNullString Then Exit Sub
                    For l& = 0 To 255
                        If InStr(1, LCase(M3.MP3GetTagGenreName((l&))), LCase(s$)) Then
                            m$ = m$ & Format(l&, "000") & " = " & _
                              M3.MP3GetTagGenreName((l&)) & vbLf: n& = n& + 1
                        End If
                    Next
                    If n& > 0 Then
                        m$ = IL(170) & " " & n& & " " & IL(171) & ":" & vbLf & vbLf & m$
                    Else: m$ = vbLf & IL(172): End If
                    M3.MsgBoxW m$
            End Select
            If txt(17).Visible Then txt(17).SetFocus: txt_GotFocus 17
    End Select
End Sub

'// Show the adjust value form
Friend Function ShowAdjustDiffMs$(OwnerForm As Form, ByVal DefValue$)
Dim n&: Caption = IL(118): cmd(9).Default = True
    Width = 2520: Height = 2190: lbl(0) = M3.GetText([Select Max Diff Ms]): txt(1) = DefValue$
    pic(3).Move 8, 8: pic(3).Visible = True: sValue$ = "": cmd(9).Picture = cmd(2).Picture
    bFocus = True: Set FocusCtrl = txt(1): Set oUserform = OwnerForm
    ShowMe 1, OwnerForm: ShowAdjustDiffMs$ = sValue$
End Function

'// Show Input form, Genre from number, Genre number from name
Friend Function ShowInputForm$(OwnerForm As Object, Promt$, Title$, Default$, OnlyNumbers As Boolean, MaxLength&, Alignment As txtAlignment)
On Local Error Resume Next: pic(7).Move 8, 8: pic(7).Visible = True
    cmd(21).Picture = cmd(2).Picture: cmd(22).Picture = cmd(0).Picture
    txt(17) = Default$: Caption = Title$: lbl(29) = Promt$: img(10) = img(0)
    Width = 5000: Height = 1600: WindowState = 0: Set oUserform = OwnerForm
    bFocus = True: Set FocusCtrl = txt(17): cmd(21).Default = True
    txt(17).Alignment = Alignment: If FormType = frmInput Then lbl(29).FontBold = False
    txt(17).MaxLength = MaxLength&: txt(17).Tag = IIf(OnlyNumbers, "#", "")
    ShowMe IIf(FormType = frmInput, 1, 0): ShowInputForm$ = sValue$
End Function


'// Show a pure text form
Friend Sub ShowTextForm(OwnerForm As Object, sText$, sCaption$)
On Local Error Resume Next
     ': SizableBorder hWnd: Width = c_W: Height = c_H: Caption = sCaption$: txt(0) = sText$
    'SizableBorder frm(frmTextForm).hWnd: WindowState = 2
    txt(0) = sText$: Caption = sCaption$: Width = c_W: Height = c_H
    WindowState = 2: Set oUserform = OwnerForm ': txt(0).Visible = True: txt(0).Move 0, 0, ScaleWidth, ScaleHeight: Show
    txt(0).Visible = True: txt(0).Move 0, 0: ShowMe ', ScaleWidth, ScaleHeight: Show
End Sub

'// Show MP3 tag batch edit form
Public Sub ShowTagBatchEdit(Optional OwnerForm As Object)
    ShowPlaylistCreator OwnerForm: Caption = "MP3 tag batch editor"
    sDirtyTag2$ = "0;?;?;?;?;?;?": Set oUserform = OwnerForm
    cmd(13).Picture = cmd(7).Picture: lst(6).Height = 2460
    lbl(7).ForeColor = lbl(17).ForeColor: sFolder$ = ""
     cmd(10).Picture = cmd(6).Picture: FixForm 10: lbl(17) = IL(76)
End Sub

'// Show playlist creator form
Public Sub ShowPlaylistCreator(Optional OwnerForm As Object)
Dim l&: For l& = 5 To 10: txt(l&) = "": Next
    Caption = "Playlist creator": ShowTagger OwnerForm: cmd(13).Picture = img(7)
    lbl(49) = IL(77): lbl(54) = IL(78): lbl(17) = IL(76): lbl(7) = "": chk(2).Visible = True
    cmd(10).Enabled = False: cmd(11).Enabled = False: cmd(13).Enabled = False
    lbl(49).Visible = True: cmd(11).Picture = cmd(7).Picture: lst(6).Move 0, 0
    cmb(2).AddItem "<" & IL(81) & ">", 0: cmb(2).ItemData(cmb(2).NewIndex) = -1
    chk(4).Visible = True: lbl(54).Visible = True: txt(16) = IL(75): cmb(2).ListIndex = 0
End Sub

'// Show file tagger
Friend Sub ShowTagger(Optional OwnerForm As Form, Optional MP3File$, Optional FormIsOpen As Boolean)
Dim l&, n&, i&, s$, t$, Sps#, Bps#: Static IsLoaded As Boolean: M3.MP3TagReset
    
    If FormType = frmTaggerForm Then
        If SaveQ = vbCancel Then Exit Sub
SelectFile:
        If LenB(MP3File$) = 0 Then
            If LenB(sFile$) = 0 Then
                MP3File$ = M3.FileFolder(encShowOpen, "MP3 " & IL(38) & " (*.mp3)|*.mp3", _
                  IL(119), , encExplorer + encHideReadOnly)
            Else: MP3File$ = sFile$: End If
        End If
        If LenB(MP3File$) Then
            If Not M3.FileExists(MP3File$) Then sFile$ = "": MP3File$ = "": GoTo SelectFile
            M3.MP3ReadTagV1 MP3File$
        Else
            If Not FormIsOpen Then: Unload Me: Exit Sub
        End If
        lbl(17) = M3.PathInfo(MP3File$, encNameExt): sFile$ = MP3File$
        Caption = "MP3-File tagger": lbl(7) = M3.FormatMs(M3.MP3LenMs(sFile$), "nn:ss") & ", " & _
          M3.FormatKMG(FileLen(sFile$), True, 2)
    End If
    'Get tag info and names to lables.
    For l& = 0 To 5: txt(l& + 5) = M3.MP3TagInfoGet(l&, t$): lbl(l& + 10) = t$ & ":": Next
    cmd(10).Picture = cmd(4).Picture: cmd(11).Picture = img(4): img(6) = img(0)
    cmd(13).Picture = img(9): pic(5).Visible = True: cmd(12).Picture = img(8)
    i& = Val(M3.MP3TagInfoGet(encTagGenreNo, t$)): lbl(16) = t$ & ":": cmb(2).Clear
    For l& = 0 To 255 'Get geners to combo
        s$ = Trim$(M3.MP3GetTagGenreName(l&))
        If LenB(s$) Then
            cmb(2).AddItem s$: cmb(2).ItemData(cmb(2).NewIndex) = l&: n& = n& + 1
        End If
    Next
    If M3.MP3TagExist Then 'Search genre and show it in combo
        For l& = 0 To cmb(2).ListCount - 1
            If i& = cmb(2).ItemData(l&) Then cmb(2).ListIndex = l&: Exit For
        Next
    Else: cmb(2).ListIndex = -1: End If
    'if cmb(2).ListIndex=-1 then cmb(2).ListIndex
    If FormType = frmPlayListCreator Or FormType = frmTagBatchEdit Then cmb(2).ListIndex = 0
    IsLoaded = True: Set oUserform = OwnerForm: bFocus = True: Set FocusCtrl = cmd(12)
    If FormType = frmTaggerForm Then TagDirty sDirtyTag$: cmd(13).Enabled = False
    If Not FormIsOpen Then
        pic(5).Move 8, 8: Height = 4400: Width = 4020: ShowMe , OwnerForm
    End If
End Sub

'// Search CDDB by category and CDID
Friend Function ShowSearchCDDB$()
Dim n&: Caption = "Search CDDB by category and CDID": cmd(4).Default = True: cmb(1).Clear
    For n& = 90 To 100: cmb(1).AddItem IL(n&): Next: bFocus = True: Set FocusCtrl = txt(2)
    Width = 4875: Height = 1100: cmb(1).ListIndex = 0: pic(0).Visible = True: pic(0).Move 8, 8
    sValue$ = "": txt_Change 4: ShowMe 1: ShowSearchCDDB$ = sValue$
End Function

'// More than one possible matches from CDDB, show info in list box, lst(0)
Friend Function ShowCDDBMatches$(Infotext$, ListItems$)
Dim v$(), n&: txt(3) = Replace(Infotext$, vbTab, " "): v$() = Split(ListItems$, vbNewLine)
    Width = c_W: Height = c_H: lst(0).Visible = True: txt(0).Visible = False: lst(0).Clear
    'If Tag = c_Det Then Left = lLeft&: Top = lTop&: Tag = "": txt(3) = sValue$: Exit Function
    For n& = 0 To UBound(v$()) - 1: lst(0).AddItem v$(n&): Next
    pic(4).Move 8, 8: Caption = n& - 2 & " " & IL(120): cmd(3).Visible = True
    pic(4).Visible = True: lst(0).Move lst(1).Left, lst(1).Top: lbl(22) = "": lbl(24) = ""
    sFolder$ = sTmpFolder$: cmd(19).Visible = False: cmd(20).Visible = False
    Erase v$(): lst(0).ListIndex = 2: ShowMe 1
End Function

'// Full text search to the freedb database.
Friend Function ShowFreeSearch$(Optional OwnerForm As Object, Optional SearchText$)
Dim n&: Width = 8500: Height = 3800: pic(1).Move 8, 8, ScaleWidth + 32, ScaleHeight + 32
    Caption = IL(101): cmd(5).Default = True: lst(2).Clear: lst(3).Clear
    For n& = 89 To 100: lst(2).AddItem IL(n&): Next 'Add category
    lst(3).AddItem IL(89): lst(3).AddItem IL(51): lst(3).AddItem IL(50)
    lst(3).AddItem IL(52): lst(3).AddItem IL(88)    'Add fields to search
    cmd(8).Picture = cmd(1).Picture: cmd(5).Picture = cmd(4).Picture
    txt(4) = SearchText$: ReDim PageInf(0): ReDim CDDBQ.Page(0): pic(1).Visible = True
    cmd(16).Picture = cmd(3).Picture: Set oUserform = OwnerForm
    lst(2).Selected(0) = True: lst(3).Selected(0) = True: lst_Click 2: lst_Click 3
    bFocus = True: Set FocusCtrl = txt(4)
    txt(4).TabIndex = 0: ShowMe , OwnerForm: Refresh: If LenB(SearchText$) Then GoToPage
End Function

'// Show Tag and rename
Friend Function ShowTagRename$(Infotext$, sCaption$)
Dim s$, d$, v$(), v2$(), n&, l&, Ms&, i&, b As Boolean: Static FormIsOpen As Boolean: Const c = ":"
    txt(0).Visible = False: lst(1).Visible = True: lst(5).Visible = True
    Width = c_W: Height = c_H: FormType = frmTagRename
    v$() = Split(Infotext$, vbNewLine & vbNewLine): chk(0).Visible = True
    v2$() = Split(v$(1), vbNewLine): lst(0).Visible = False
    pic(4).Visible = True: lst(1).Visible = True: pic(4).Move 8, 8
    lst(5).Visible = True: lst(1).Clear: lst(5).Clear: s$ = M3.TagRenameCheck(sTmpFolder$)
    sInfo$ = s$ & sTmpStr$: sTmpStr$ = "" ': If LenB(s$) Then M3.ShowTextForm sInfo$, sCaption$
    v$(0) = Replace(v$(0), vbNewLine, ", "): v$(0) = Replace(v$(0), vbTab, " ")
    txt(3) = IL(106) & vbNewLine & v$(0): b = False
    For n& = 0 To UBound(v2$()) 'Add tracks from the matched album
        v$() = Split(v2$(n&), vbTab)
        If InStr(1, v2$(n&), IL(46)) = 0 Then
            Ms& = lTrackTimeMs&(n&): s$ = "  " & M3.FormatMs(Ms&, "nn:ss") & "  " & sTmpArr$(n&)
            d$ = v$(0) & vbTab & v$(1)
        Else
            b = True
            If v$(0) = IL(46) Then 'More files in destination folder than tracks
                
                Ms& = lTrackTimeMs&(n&): d$ = ""
                s$ = "  " & M3.FormatMs(Ms&, "nn:ss") & "  " & sTmpArr$(n&)
            Else 'More tracks than files in destination folder
                s$ = "  " & IL(47): d$ = v$(0) & vbTab & v$(1): l& = l& + 1
            End If
        End If
        lst(5).AddItem s$: lst(5).ItemData(i&) = Ms&: i& = i& + 1 'Store lenght in Ms
        If LenB(d$) Then lst(1).AddItem d$
    Next
    lbl(22) = lst(1).ListCount & ", " & IL(45) & ":": n& = 0
    lbl(24) = IL(48) & " (" & lst(5).ListCount - l& & "):"
    If Not FormIsOpen Then
        For l& = 0 To 255 'Get geners to combo
            s$ = Trim$(M3.MP3GetTagGenreName(l&))
            If LenB(s$) Then
                cmb(3).AddItem s$, 0: cmb(3).ItemData(0) = l&: n& = n& + 1
            End If
        Next
        
        For l& = 60 To 68: cmb(4).AddItem IL(l&): Next 'Rename file as
    End If
    'Lable caption and default values Artist, Album, Year, Genre, comment...
    M3.MP3TagInfoGet encTagArtist, s$: lbl(36) = s$ & c: lbl(32) = IL(8) & c
    M3.MP3TagInfoGet encTagAlbum, s$: lbl(28) = s$ & c: chk(0).Caption = IL(1)
    M3.MP3TagInfoGet encTagYear, s$: lbl(38) = s$ & c: chk(1).Caption = IL(0)
    M3.MP3TagInfoGet encTagGenre, s$: lbl(26) = s$ & c: FormIsOpen = True
    M3.MP3TagInfoGet encTagComment, s$: lbl(34) = s$ & c: txt(15) = InfoCDDB.Year
    txt(12) = InfoCDDB.Album: txt(13) = m_Settings.DefComment: txt(14) = InfoCDDB.Artist
    chk(1).Value = IIf(m_Settings.CreatePlaylist, 1, 0)
    lst(5).ForeColor = IIf(b, vbRed, vbBlack): lst(1).ListIndex = UBound(v2$())
    lst(5).ListIndex = lst(1).ListIndex: sFolder$ = sTmpFolder$: sTmpFolder$ = "": chk_Click 0
    cmd(19).Picture = img(3): cmd(20).Picture = img(4): lst_Click 5: i& = M3.GenreFromCDDB%
    For l& = 0 To cmb(3).ListCount - 1 'Search genre and show default in combo
        If i& = cmb(3).ItemData(l&) Then cmb(3).ListIndex = l&: Exit For
    Next
    cmb(4).ListIndex = m_Settings.MP3FileFormat ':bFocus = True: Set FocusCtrl = txt(4)
    Caption = IIf(LenB(sCaption$), sCaption$, Tag): ShowMe 1: ShowTagRename$ = sValue$
End Function

'// Show the file size calculator
Friend Sub ShowFileCalc(Optional OwnerForm As Form)
Dim n&, v$(): v$() = M3.EncBitrates
'M3.EncBitrates return a string array with valid bitrates, Index 0 = " Kbps", 1 = "320"...
    If cmb(5).ListCount = 0 Then
        For n& = 1 To UBound(v$)
            cmb(5).AddItem v$(n&) & v$(0): cmb(5).ItemData(n& - 1) = v$(n&)
        Next
    End If
    Caption = "File size calculator": Set oUserform = OwnerForm
    Width = 5750: Height = 3250: img(2) = img(0): Erase v$()
    cmd(14).Picture = img(4): cmd(15).Picture = img(8)
    pic(6).Move 8, 8: pic(6).Visible = True
    bFocus = True: Set FocusCtrl = cmb(5)
    cmb(5).ListIndex = 3: ShowMe , OwnerForm: Calc
End Sub

'// Let user select CD-Drive
Friend Function ShowDrive$(CurDrv$, Avalible$)
Dim v$(), l&
    Width = 2650: Height = 1100: v$() = Split(Avalible$, "|"): pic(2).Move 8, 8: cmb(0).Clear
    For l& = 0 To UBound(v$())
        cmb(0).AddItem v$(l&): If Left$(v$(l&), 1) = CurDrv$ Then cmb(0).ListIndex = l&
    Next
    'cmd(2).Visible = True: cmb(0).Visible = True: txt(3).Top = 8
    Caption = "Select CD-Drive": lbl(6) = "CD-Drive" & ":": pic(2).Visible = True
    bFocus = True: Set FocusCtrl = cmb(0)
    Erase v$(): ShowMe 1: ShowDrive$ = sValue$: Set frmMX3 = Nothing
End Function

Private Sub chk_Click(Index As Integer) 'Select all or none tracks
Dim n&, l&, m&
    If Index = 0 Then
        lst(1).Visible = False: lst(5).Visible = False: bNotNow = True
        l& = IIf(LenB(sFolder$), lst(5).ListIndex, lst(5).ListCount - 1)
        For n& = 0 To lst(1).ListCount - 1
            lst(1).Selected(n&) = chk(0): lst(5).Selected(n&) = chk(0)
        Next
        If lst(1).ListCount - 1 >= l& Then lst(1).ListIndex = l&
        bNotNow = False: lst(5).ListIndex = l&
        lst(1).Visible = True: lst(5).Visible = True
    ElseIf Index = 3 Then
        bNotNow = True: l& = lst(6).ListIndex:  m& = lst(6).TopIndex
        For n& = 0 To lst(6).ListCount - 1: lst(6).Selected(n&) = chk(3): Next
        lst(6).ListIndex = -1: lst(6).ListIndex = l&: lst(6).TopIndex = m&
        bNotNow = False: cmd(13).Enabled = chk(3)
    End If
End Sub

'// Recalculate in Filesize calculator when bitrate change
Private Sub cmb_Click(Index As Integer)
    If bNotNow Then Exit Sub
    If Index = 5 Then Calc
    If FormType = frmTaggerForm And Index = 2 Then cmd(13).Enabled = TagDirty
    If FormType = frmTagBatchEdit Then TagDirty
End Sub

'// Set focus on default control
Private Sub Form_Activate()
    On Error Resume Next: If bFocus Then FocusCtrl.SetFocus: bFocus = False
End Sub

Private Sub Form_Resize()
    If FormType = frmTextForm Then txt(0).Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub lst_Click(Index As Integer)
Dim l&, s$, n&, m&, b As Boolean, e(3) As Boolean: Static IsHere As Boolean, sPrevF$
    Select Case Index
        Case 2, 3 'Search options & Category listbox, if All is selected
            If IsHere Then Exit Sub Else n& = lst(Index).ListIndex: IsHere = True
            If lst(Index).Selected(0) Then b = True 'All selected
            If n& = 0 Then
                For l& = 1 To lst(Index).ListCount - 1
                    If b Then
                        lst(Index).Selected(l&) = True
                    Else
                        lst(Index).Selected(l&) = lst(Index).ItemData(l&)
                    End If
                Next
            Else
                l& = lst(Index).ListIndex: lst(Index).ItemData(l&) = lst(Index).Selected(l&)
            End If
            e(1) = Len(txt(4)): If n& = 0 Then lst(Index).ListIndex = 0
            For n& = 2 To 3 'Check if enable search button cmd(5)
                For l& = 1 To lst(n&).ListCount - 1
                    If lst(n&).Selected(l&) Then e(n&) = True: If n& = Index Then m& = m& + 1
                Next
            Next
            If m& = lst(Index).ListCount - 1 Then
                lst(Index).Selected(0) = True   'All selected
            Else
                lst(Index).Selected(0) = False  'Not all selected
            End If
            cmd(5).Enabled = e(1) And e(2) And e(3): IsHere = False
        Case 0      'Not aloud to select column in listbox
            If lst(Index).ListIndex < 2 Then lst(Index).ListIndex = 2
        Case 1, 5    'Enabled OK button if files selected
            If IsHere Or bNotNow Then Exit Sub Else IsHere = True
            For l& = 0 To lst(5).ListCount - 1
                If lst(1).ListCount - 1 >= l& Then
                    lst(5).Selected(l&) = lst(1).Selected(l&)
                End If
                If l& <= lst(1).ListCount - 1 Then If lst(1).Selected(l&) Then b = True
                If lst(5).ListIndex = l& Then n& = l&
            Next
            If Index = 1 Then
                lst(5).ListIndex = lst(1).ListIndex: lst(5).TopIndex = lst(1).TopIndex
            ElseIf Index = 5 And n& <= lst(1).ListCount - 1 Then
                l& = lst(5).ListCount - lst(1).ListCount
                lst(1).ListIndex = n&: lst(1).TopIndex = lst(5).TopIndex
            End If
            lst(5).ToolTipText = sFile$
            cmd(17).Enabled = (lst(5).ListIndex > 0) 'Upp/Down Buttons
            cmd(18).Enabled = (lst(5).ListIndex < (lst(5).ListCount - 1))
            IsHere = False: s$ = lst(5).Text: If s$ = "  " & IL(47) Then Exit Sub
            s$ = Right$(s$, Len(s$) - 9): sFile$ = sFolder$ & s$
            b = IIf(sPrevF$ <> sFile$, True, False): sPrevF$ = sFile$
            If LenB(sAlias$) And b Then cmd_Click 20: cmd_Click 20
        Case 4      'Add tracks to txt(11)
            txt(11) = "": n& = lst(4).ListIndex + 1
            For l& = 1 To Albums(n&).nTracks
                s$ = s$ & Format(l&, "00") & "   " & Albums(n&).Track(l&).Time & _
                  "   " & Albums(n&).Track(l&).Title & vbNewLine
            Next
            txt(11) = Left$(s$, Len(s$) - 2): cmd(8).Enabled = True: cmd(16).Enabled = True
        Case 6 'Listbox in file tagger search
            If bNotNow Then Exit Sub
            For n& = 0 To lst(6).ListCount - 1
                If lst(6).Selected(n&) Then b = True: Exit For
            Next
            sFile$ = sFiles$(lst(6).ListIndex): lst(6).ToolTipText = sFile$
            cmd(13).Enabled = b: If LenB(sAlias$) Then cmd_Click 11: cmd_Click 11
    End Select
End Sub

Private Sub lst_DblClick(Index As Integer)
    If Index = 0 Then cmd_Click 3    'View details if double click listbox (lst(0))
    If Index = 5 Then M3.ShowTaggerMP3 Me, sFolder$ & Right$(lst(5).Text, Len(lst(5).Text) - 9)
    If Index = 6 Then M3.ShowTaggerMP3 Me, sFiles$(lst(6).ListIndex)
End Sub

'// Select all text when  get focus
Private Sub txt_GotFocus(Index As Integer)
    If Index <> 0 Then txt(Index).SelStart = 0: txt(Index).SelLength = Len(txt(Index))
End Sub
Private Sub txt_Change(Index As Integer)
    If bNotNow Then Exit Sub
    cmd(4).Enabled = (Len(txt(2)) = 8)             'Search if CDID length is 8 charcters
    cmd(5).Enabled = (Len(txt(4)))                 'Text search CDDB
    If FormType = frmTaggerForm Then cmd(13).Enabled = TagDirty
    If FormType = frmTagBatchEdit Then TagDirty     'Save info if changed
    If FormType = frmGenreFromNumber Or FormType = frmGenreNumberFromName Then
        cmd(21).Enabled = LenB(txt(17))
        If FormType = frmGenreFromNumber Then
            If Val(txt(17)) > 255 Then txt(17) = "255": txt_GotFocus 17
        End If
    End If
End Sub


'// Full text search to the freedb database.
Private Function QueryCDDB(ByVal Text$, Optional ByVal PageNo%) As Boolean
Dim s$, sQ$, t$, v$(), l&, n&, i&, lF&, lT&, NewCDDBQ As PageQ: Static PrevQString$
On Error GoTo QueryCDDBErr
    With CDDBQ
        sQ$ = "http://www.freedb.org/freedb_search.php?words=" & Replace(Text$, " ", "+") 'M3.EncodeUniURL$(Text$)
        If Not lst(3).Selected(0) Then 'Not all fields to search
            For l& = 1 To lst(3).ListCount - 1
                If lst(3).Selected(l&) Then
                    i& = Choose(l&, 51, 50, 52, 88): s$ = s$ & "&fields[]=" & ILOrg(i&)
                End If
            Next
        End If
        If Not lst(2).Selected(0) Then 'Not all categorys to search
            For l& = 1 To lst(2).ListCount - 1
                If lst(2).Selected(l&) Then t$ = t$ & "&cats[]=" & LCase(ILOrg(89 + l&))
            Next
        End If
        sQ$ = sQ$ & s$ & t$ & IIf(LenB(s$), "&allfields=NO", "") & _
          IIf(LenB(t$), "&allcats=NO", "") & "&grouping=none": i& = 0
        If sQ$ <> PrevQString$ Then 'Different query string than last
            PageNo% = 0
        Else 'User have click on search, but nothing have change
            If PageNo% = 0 And .Page(CDDBQ.CurPage).Read Then PageNo% = .CurPage
        End If
        If .Page(PageNo%).Read Then 'Page already read
            .CurPage = PageNo%: Albums = .Page(PageNo%).Albums: QueryCDDB = True: Exit Function
        End If
        If PageNo% = 0 Then 'New query
            CDDBQ = NewCDDBQ: ReDim .Page(0)
            .ToFile = M3.PathApp & "freedbtest.txt": s$ = sQ$
        Else: s$ = sQ$ & "&page=" & PageNo%: End If
    'Download the HTML page and read result of query string
        M3.DownloadFile s$ & "#fdbsr", .ToFile: l& = FreeFile: PrevQString$ = sQ$
        Open .ToFile For Input As #l&: s$ = Input(LOF(l&), #l&): Close #l&
    'Check if any match
        t$ = "class=searchFormGlobal><br>": lF& = InStr(1, s$, t$): lT& = InStr(lF&, s$, ".<")
        t$ = Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$) + 1)
        If t$ = ILOrg(121) Then M3.MsgBoxW IL(121), 48: Exit Function
    'Get # of results and # of pages
        If PageNo% = 0 Then
            v$() = Split(t$, " "): lT& = 1: .nPages = Val(v$(5))
            .nResults = Val(v$(0)): ReDim .Page(.nPages): PageNo% = 1
        End If
    End With
    Do 'Loop all results on current page
        With CDDBQ.Page(PageNo%)
            t$ = "searchU1" & n& + 1: lF& = InStr(lT&, s$, t$)
            If lF& = 0 Then 'Done reading the page
                CDDBQ.CurPage = PageNo%: Albums = .Albums '(PageNo%)
                Erase v$(): .nAlbums = n&: QueryCDDB = True: .Read = True: Exit Function
            End If
            n& = n& + 1: ReDim Preserve .Albums(n&)
            lF& = InStr(lF&, s$, ">"): lT& = InStr(lF& + 1, s$, "<")
        End With
        With CDDBQ.Page(PageNo%).Albums(n&)
        'Artist / Album
            t$ = Mid$(s$, lF& + 1, lT& - lF& - 1): v$() = Split(t$, " / ")
            If UBound(v$()) > 0 Then
                .Artist = v$(0): .Album = v$(1)
            Else: .Artist = t$: End If
        '# of tracks on album
            t$ = "Tracks:</b>": lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, c_BR)
            .nTracks = Val(Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$)))
        'Total play time on album
            t$ = "time:</b>": lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, c_BR)
            t$ = Trim$(Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$)))
            .TimeTot = IIf(Len(t$) = 4, "0", "") & t$
        'Year
            t$ = "Year:</b>": i& = InStr(lT&, s$, t$)
            If i& < lT& + 10 And i& <> 0 Then 'Check if year is missing
                lF& = i&: lT& = InStr(lF&, s$, c_BR)
                .Year = Trim$(Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$)))
            End If
        'Category
            t$ = "Disc-ID:</b>": lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, c_SL)
            v$() = Split(Mid$(s$, lF& + Len(t$), 50), c_SL): .Category = Trim$(v$(0))
        'Disc ID
            t$ = c_SL & .Category & c_SL: lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, """>")
            .DiscID = Trim$(Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$))): ReDim Preserve .Track(.nTracks)
        'Loop all tracks on current album
            For l& = 1 To .nTracks
                t$ = c_T1 & l& & c_T2: lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, c_TD)
            'Track title
                v$() = Split(Mid$(s$, lF& + Len(t$), 50), "</"): .Track(l&).Title = Trim$(v$(0))
                t$ = "nowrap>": lF& = InStr(lT&, s$, t$): lT& = InStr(lF&, s$, c_TD)
            'Track length
                t$ = Trim$(Mid$(s$, lF& + Len(t$), lT& - lF& - Len(t$)))
                .Track(l&).Time = IIf(Len(t$) = 4, "0", "") & t$
            Next
        End With
    Loop
    Exit Function
QueryCDDBErr:
    M3.Log "Error: QueryCDDB(): " & Err.Description, encLogErrorNoMsgBox
    Err.Clear: Exit Function
    Resume
End Function

'// Start new search if PageNo% = 0, or else go to selected page
Private Sub GoToPage(Optional ByVal PageNo%)
Dim n&, s$: If PageNo% = 0 Then ReDim CDDBQ.Page(0): CDDBQ.CurPage = 0

    Enabled = False: ReDim Albums(0): lbl(4) = IL(112) & "... " & IL(122) & _
      IIf(PageNo%, " " & IL(123) & " " & PageNo%, ""): txt(11) = ""
    lbl(4).Refresh: lst(4).Clear: ReDim Album(0): cmd(8).Enabled = False
    cmd(16).Enabled = False: QueryCDDB txt(4), PageNo%
    With CDDBQ
        lbl(3) = .CurPage & c_SL & .nPages: txt(11) = "" ': chk(1).Value = 0
        cmd(6).Enabled = .CurPage > 1: cmd(7).Enabled = .CurPage < .nPages: lst(4).Clear
        lbl(4) = .Page(.CurPage).nAlbums & " " & "albums from page" & " " & .CurPage & ", " & _
          .nResults & " " & "album(s) on" & " " & .nPages & " " & "page(s)"
    End With
    s$ = String(4, " "): Enabled = True
    For n& = 1 To UBound(Albums)     'Loop all albums on current page
        With Albums(n&)
            lst(4).AddItem Replace(.Year, vbNullChar, "  ") & s & .TimeTot & s & .DiscID & _
              vbTab & UCase(.Category) & String(20 - Len(s) - _
              Len(.Category), " ") & vbTab & .Artist & "  " & .Album
        End With
    Next
    
End Sub

'// Check if tag info have change
Private Function TagDirty(Optional RetDirtyTag$) As Boolean
Dim l&, s$: s$ = cmb(2).ListIndex & ";": For l& = 0 To 5: s$ = s$ & txt(l& + 5) & ";": Next
On Error Resume Next
    If FormType = frmTagBatchEdit Then
        If bBatch Then sDirtyTag2$ = s$ Else sDirtyTag$ = s$
    End If
    TagDirty = (s$ <> sDirtyTag$): RetDirtyTag$ = s$: If Err Then Err.Clear
End Function

Private Function FixForm(Index%) As Boolean
Dim l&, n&, s$, t$, v$(), b As Boolean
    b = (Index% = 13): b = (b <> bBatch): bBatch = (Index% = 13)
    If b Then 'Edit text fields and genre combo if tag batch edit mode
        If Not bBatch Then
            s$ = "<" & IL(81) & ">": t$ = sDirtyTag$
        Else: s$ = "<?>": t$ = sDirtyTag2$: End If
        cmb(2).RemoveItem 0: cmb(2).AddItem s$: cmb(2).ItemData(0) = -1
        bNotNow = True: v$() = Split(t$, ";"): cmb(2).ListIndex = Val(v$(0))
        For l& = 1 To 6: txt(l& + 4) = v$(l&): Next: Erase v$(): bNotNow = False
    End If
    lbl(7) = "": lst(6).Visible = False: For l& = 2 To 4: chk(l&).Visible = False: Next
    lbl(54).Visible = False: lbl(49).Visible = True: txt(16).Visible = False
    If FormType = frmPlayListCreator Then n& = 33 Else n& = 33
    
    Select Case Index
        Case 12 'Select folder
            If Len(sFolder$) > 40 Then
                v$() = Split(sFolder$, "\"): l& = UBound(v$()) - 1: s$ = v$(0) & "\...\" & v$(l&)
                If l& > 2 Then s$ = v$(0) & "\...\" & v$(l& - 1) & "\" & v$(l&)
            Else: s$ = sFolder$: End If
            cmd(13).Picture = IIf(FormType = frmPlayListCreator, img(7), cmd(7).Picture)
            lbl(17) = s$: lbl(49) = IL(77): cmd(10).Picture = cmd(4).Picture
            cmd(11).Enabled = False: cmd(13).Enabled = False: lbl(15) = IL(55) & ":"
            chk(4).Visible = True: lbl(54).Visible = True:  Erase v$()
            chk(2).Visible = True: cmd(11).Picture = cmd(7).Picture
            cmd(10).Enabled = LenB(sFolder$): lbl(17).Tag = s$
        Case 10 'Search or go back to Select folder
            If cmd(10).Picture = cmd(6).Picture Then 'Go back to search
                lbl(49) = IL(77): cmd(10).Picture = cmd(4).Picture: lbl(17) = lbl(17).Tag
                cmd(11).Picture = cmd(7).Picture: lbl(54).Visible = True: chk(2).Visible = True
                chk(4).Visible = True: lbl(15) = IL(55) & ":"
                cmd(13).Picture = IIf(FormType = frmPlayListCreator, img(7), cmd(7).Picture)
            Else
                If sFiles$(0) <> vbNullString Then 'MP3-Files found in folder(s)
                    n& = UBound(sFiles$()): cmd(11).Enabled = True: lbl(15) = IL(80) & ":"
                    lbl(17) = n& + 1 & " " & LCase(IL(38)) & " " & IL(69) & ", " & IL(105) & "."
                    lbl(49) = IL(1): lst(6).Visible = True: chk(3).Visible = True
                    cmd(10).Picture = cmd(6).Picture: txt(16).Visible = True
                    cmd(11).Picture = img(IIf(LenB(sAlias$), 5, 4)): lbl(7).Tag = lbl(17)
                Else: lbl(17) = IL(121): End If
            End If
            n& = lst(6).ListCount: cmd(11).Enabled = n&: cmd(13).Enabled = n&
        Case 11 'Go to MP3-Files found,Play or Stop Play MP3 file
            If cmd(10).Picture = cmd(4).Picture Then 'Go back to search
                chk(3).Visible = True: cmd(10).Picture = cmd(6).Picture
                cmd(11).Picture = img(IIf(LenB(sAlias$), 5, 4))
                cmd(13).Picture = IIf(FormType = frmPlayListCreator, img(7), cmd(7).Picture)
                txt(16).Visible = True: lbl(15) = IL(80) & ":": lbl(17) = lbl(7).Tag: bNotNow = False
                lbl(49) = IL(1): lst(6).Visible = True: FixForm = True
            ElseIf cmd(13).Picture = cmd(7).Picture Or cmd(11).Picture = cmd(6).Picture Then
                cmd(10).Picture = cmd(4).Picture: bNotNow = False: cmd_Click 11: FixForm = True
            End If
        Case 13 'Save batch edit/Playlist
            cmd(13).Picture = IIf(FormType = frmPlayListCreator, img(7), img(9))
            cmd(10).Picture = cmd(6).Picture: cmd(11).Picture = cmd(6).Picture
            lbl(49).Visible = False: lbl(15) = IL(55) & ":": lbl(17) = IL(103): lbl(7) = IL(104)
    End Select
End Function

'// Create a playlist of selected files in "Playlist Creator"
Private Sub PlaylistDo()
Dim l&, n&, m&, c&, s$, f$, sF$, sFB$, t$
    sF$ = txt(16): If LenB(sF$) = 0 Then sF$ = IL(75)
    sF$ = sFolder$ & M3.FileReplBad$(sF$) & ".m3u": l& = FreeFile
    If M3.FileExists(sF$) Then 'Create backup if the playlist file exists
        sFB$ = M3.BackupFile(sF$): sFB$ = vbLf & vbLf & IL(79) & ":" & vbLf & sFB$
    End If
    Open sF$ For Output As #l&: n& = lst(6).ListCount - 1: Print #l&, "#EXTM3U": t$ = lbl(17)
    MousePointer = 13: m_Settings.IsWorking = True
    For m& = 0 To n&
        If m_Settings.StopWork Then Exit For
        If lst(6).Selected(m&) Then
            lbl(17) = IL(124) & " " & CInt((((m& + 1)) / n&) * 100) & "%"
            M3.EventRaise IL(124), CInt((((m& + 1)) / n&) * 100)
            f$ = sFiles$(m&): M3.MP3ReadTagV1 f$: s$ = M3.MP3TagInfoGet(encTagTitle)
            If Trim$(s$) = vbNullString Then s$ = " - "
            Print #l&, "#EXTINF:" & M3.MP3LenMs(f$) / 1000 & "," & s$
            Print #l&, Right$(f$, Len(f$) - Len(sFolder$)): c& = c& + 1
        End If
    Next
    If m_Settings.StopWork Then s$ = IL(136) & "! ": n& = 52 Else s$ = "": n& = 36
    m_Settings.IsWorking = False: m_Settings.StopWork = False
    Close #l&: lbl(17) = t$: MousePointer = 0: If c& = 0 Then Exit Sub
    s$ = s$ & IL(72) & " (" & c& & " " & LCase(IL(38)) & ").": M3.EventRaise s$, 0
    If M3.MsgBoxW(s$ & vbLf & sF$ & sFB$ & vbLf & vbLf & IL(73), n&) = vbYes Then M3.ExecuteShell sF$
End Sub

'// Edit selected MP3 files, in "Tag Batch Editor"
Private Sub BatchEditDo()
Dim s$, t$, v$(), v2$(), n&, m&, c&, f&: n& = lst(6).ListCount - 1
    v$() = Split(sDirtyTag2, ";")
    For m& = 0 To 5
        If v$(m& + 1) <> "?" Then
            M3.MP3TagInfoGet m&, t$: s$ = s$ & t$ & ":" & _
              vbTab & IIf(m& <> 5, vbTab, "") & txt(5 + m&) & vbLf
        End If
    Next
    If Val(v$(0)) <> 0 Then 'Category
        M3.MP3TagInfoGet encTagGenre, t$: s$ = s$ & t$ & ":" & _
          vbTab & vbTab & cmb(2).Text & vbLf
    End If
    If LenB(s$) = 0 Then M3.MsgBoxW IL(85): GoTo BatchEditDoExit Else t$ = ""
    For m& = 0 To n&
        If lst(6).Selected(m&) Then c& = c& + 1: t$ = t$ & sFiles$(m&) & ";"
    Next
    If c& = 0 Then M3.MsgBoxW IL(85): GoTo BatchEditDoExit 'No files selected
    If M3.MsgBoxW(IL(83) & " " & c& & " " & IIf(c& > 1, LCase(IL(38)), IL(39)) & _
      "." & vbLf & vbLf & s$ & vbLf & IL(84), _
      vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then GoTo BatchEditDoExit
    v2$() = Split(t$, ";"): n& = UBound(v2$()) - 1
    m& = Val(v$(0)): If m& = -1 Then s$ = "?" Else s$ = cmb(2).ItemData(m&)
    For m& = 0 To n& 'Apply changes
        M3.MP3ReadTagV1 v2$(m&): DoEvents
        If M3.MP3TagInfoLet(v2$(m&), v$(1), v$(2), v$(3), _
          v$(4), s$, v$(5), v$(6)) Then f& = f& + 1
    Next
    If f& <> c& Then m& = 48 Else m& = 64
    M3.MsgBoxW IL(86) & " " & f& & "/" & c& & " " & IIf(c& > 1, LCase(IL(38)), IL(39)) & _
      " " & IL(87) & ".", m&
BatchEditDoExit:
    Erase v$(): Erase v2$()
End Sub

'// Ask if like to save changes made to tag
Function SaveQ() As VbMsgBoxResult
    If FormType = frmTaggerForm And cmd(13).Enabled And Visible Then
    
        SaveQ = MsgBox(txt(5) & " " & txt(6) & " " & txt(7) & "." & vbLf & vbLf & IL(110), 547)
        If SaveQ = vbYes Then cmd_Click 13
        If SaveQ = vbCancel Then WindowState = 0: Visible = True: frm(frmTaggerForm).SetFocus
    End If

End Function

'// Calculate data and display in form "File Size Calculator"
Private Sub Calc()
Dim s$, l&, Sz&, Hz#, BR%: If cmb(5).ListCount = 0 Then Exit Sub
    If cmb(5).ListIndex < 0 Then cmb(5).ListIndex = 3
    If LenB(sFile$) = 0 Then 'No file selected
        s$ = IL(111) & " :-)"
    Else
        If bNewFile Then 'This coz long time request some files from MCI.
            lbl(53) = IL(112) & "... " & IL(113)
            cmb(5).Enabled = False: cmd(15).Enabled = False
            MousePointer = 11: lbl(53).ForeColor = c_Red: Refresh: DoEvents
        End If
        BR% = cmb(5).ItemData(cmb(5).ListIndex) 'Get bitrate from combo
        'Sz& = calculated file size in bytes, l& returns length in ms,-
        'Hz# sampling frequency in Hz -1 if fail, s$ error string MCI.
        Sz& = M3.MP3CalcSize&(sFile$, BR%, False, l&, Hz#, s$)
        If Sz& <> -1 Then
          'Lenght source file milliseconds to mm:ss
            lbl(50) = M3.FormatMs$(l&, "nn:ss", True)
          'Calculated size new file, formated to i.e Kb, Mb...
            lbl(51).ForeColor = IIf(Sz& > SzFile&, c_Red, c_Green) 'Larger than source file:S
            lbl(51) = M3.FormatKMG$(Sz&, True, 2)
          'Formula used for calculation
            lbl(52) = "((" & l& & " / 1000) * " & _
              BR% & " * 1024) / 8 = " & Sz& & " byte"
          'Display source file name and size
            If bNewFile And Hz# <> -1 Then _
              sEncFileInfo$ = sEncFileInfo$ & " (" & Hz# / 1000 & " kHz)"
            s$ = sEncFileInfo$
        Else
            For l& = 50 To 52: lbl(l&) = "": Next 'Error from MCI
            If bNewFile Then M3.MsgBoxW s$ & vbLf & vbLf & sFile$, 16, Caption
        End If
    End If
    lbl(53).ForeColor = IIf(Sz& = -1, c_Red, &H8000000D): lbl(53) = s$
    If bNewFile Then
        cmb(5).Enabled = True: cmd(15).Enabled = True
        If Visible Then SetFocus: cmb(5).SetFocus
        MousePointer = 0: Refresh: bNewFile = False
    End If
End Sub

'// Edit form to sizable border with max button
Private Sub SizableBorder(hWnd&, Optional Value As Boolean = True)
Dim l&: Const CB = 262144, CM = 65536, c2 = (-16)
   l& = GetWindowLong(hWnd&, c2): If Value Then l& = l& Or CB Else l& = l& And Not CB
   SetWindowLong hWnd&, c2, l&: SetWindowPos hWnd&, 0, 0, 0, 0, 0, 39 'Sizable
   l& = GetWindowLong(hWnd&, c2): If Value Then l& = l& Or CM Else l& = l& And Not CM
   SetWindowLong hWnd&, c2, l&: SetWindowPos hWnd&, 0, 0, 0, 0, 0, 39 'MaxButton
End Sub

Private Sub ShowMe(Optional Modal%, Optional OwnerForm As Form)
Dim t&, l&: t& = m_Settings.FrmTop(FormType): l& = m_Settings.FrmLeft(FormType)
On Local Error Resume Next
    If l& + t& = 0 Then 'Center screen
        l& = Screen.Width / 2 - (Width / 2): t& = Screen.Height / 2 - (Height / 2)
    End If
    Me.Move l&, t&: Me.Show Modal%, OwnerForm
    If Err Then
        If Err = 401 Then Show 1 'Can't show non-modal form when modal form is displayed
        Err.Clear
    End If
End Sub


'// Propertys file size calculator

'//Bitrate property file size calculator
Public Property Let CalcBitrate(Kbps As encEncodeBitRates)
Dim x%
    For x% = 0 To cmb(5).ListCount - 1
        If cmb(5).ItemData(x%) = Kbps Then cmb(5).ListIndex = x%: Exit For
    Next
End Property
Public Property Get CalcBitrate() As encEncodeBitRates
    CalcBitrate = cmb(5).ItemData(cmb(5).ListIndex)
End Property

'//File path property, (store path in forms tag property and call cmd_Click to calculate)
Public Property Let CalcFilePath(Path As String)
    Tag = Path: cmd_Click 15
End Property
Public Property Get CalcFilePath$()
    CalcFilePath$ = sFile$
End Property

'// Unload form
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not TimeToUnload Then
        If SaveQ = vbCancel Then Cancel = True: Exit Sub
        On Error Resume Next
        If LenB(sAlias$) Then M3.MP3Stop sAlias$ 'Stop play
        m_Settings.FrmLeft(FormType) = Left
        m_Settings.FrmTop(FormType) = Top
        Cancel = True: Hide: oUserform.SetFocus: Exit Sub
    End If
End Sub
Private Sub Form_Terminate()
    If LenB(sAlias$) Then M3.MP3Stop (sAlias$)   'Stop play
    Dim PQ As PageQ: CDDBQ = PQ: Erase Albums(): Erase sFiles$()
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 17 And txt(17).Tag = "#" Then
        Select Case KeyAscii 'Accepts only numeric input
          Case vbKeyLeft, vbKeyRight, vbKeyUp, vbKeyDown, vbKeyTab, _
            vbKeyBack, vbKeyClear, vbKeyDelete, vbKey0 To vbKey9
          Case Else: KeyAscii = 0: Beep
        End Select
    End If
End Sub
