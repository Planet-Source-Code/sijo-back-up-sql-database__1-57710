VERSION 5.00
Begin VB.UserControl ctrCalendar 
   ClientHeight    =   3345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2760
   LockControls    =   -1  'True
   PropertyPages   =   "ctrCalendar.ctx":0000
   ScaleHeight     =   3345
   ScaleWidth      =   2760
   ToolboxBitmap   =   "ctrCalendar.ctx":0013
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2940
      Left            =   0
      TabIndex        =   1
      Top             =   -90
      Width           =   2730
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1875
         Left            =   15
         ScaleHeight     =   1875
         ScaleWidth      =   2700
         TabIndex        =   0
         Top             =   705
         Width           =   2700
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   41
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":0325
            MousePointer    =   99  'Custom
            TabIndex        =   55
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   40
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":0537
            MousePointer    =   99  'Custom
            TabIndex        =   54
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   39
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":0749
            MousePointer    =   99  'Custom
            TabIndex        =   53
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   38
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":095B
            MousePointer    =   99  'Custom
            TabIndex        =   52
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   37
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":0B6D
            MousePointer    =   99  'Custom
            TabIndex        =   51
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   36
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":0D7F
            MousePointer    =   99  'Custom
            TabIndex        =   50
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   42
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":0F91
            MousePointer    =   99  'Custom
            TabIndex        =   49
            Top             =   1575
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   34
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":11A3
            MousePointer    =   99  'Custom
            TabIndex        =   48
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   33
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":13B5
            MousePointer    =   99  'Custom
            TabIndex        =   47
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   32
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":15C7
            MousePointer    =   99  'Custom
            TabIndex        =   46
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   31
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":17D9
            MousePointer    =   99  'Custom
            TabIndex        =   45
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   30
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":19EB
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   29
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":1BFD
            MousePointer    =   99  'Custom
            TabIndex        =   43
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   35
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":1E0F
            MousePointer    =   99  'Custom
            TabIndex        =   42
            Top             =   1260
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   27
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":2021
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   26
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":2233
            MousePointer    =   99  'Custom
            TabIndex        =   40
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   25
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":2445
            MousePointer    =   99  'Custom
            TabIndex        =   39
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   24
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":2657
            MousePointer    =   99  'Custom
            TabIndex        =   38
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   23
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":2869
            MousePointer    =   99  'Custom
            TabIndex        =   37
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   22
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":2A7B
            MousePointer    =   99  'Custom
            TabIndex        =   36
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   21
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":2C8D
            MousePointer    =   99  'Custom
            TabIndex        =   35
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   20
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":2E9F
            MousePointer    =   99  'Custom
            TabIndex        =   34
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   19
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":30B1
            MousePointer    =   99  'Custom
            TabIndex        =   33
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   18
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":32C3
            MousePointer    =   99  'Custom
            TabIndex        =   32
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   17
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":34D5
            MousePointer    =   99  'Custom
            TabIndex        =   31
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   16
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":36E7
            MousePointer    =   99  'Custom
            TabIndex        =   30
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   15
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":38F9
            MousePointer    =   99  'Custom
            TabIndex        =   29
            Top             =   645
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   28
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":3B0B
            MousePointer    =   99  'Custom
            TabIndex        =   28
            Top             =   960
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   14
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":3D1D
            MousePointer    =   99  'Custom
            TabIndex        =   27
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   13
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":3F2F
            MousePointer    =   99  'Custom
            TabIndex        =   26
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   12
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":4141
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   11
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":4353
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   10
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":4565
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   9
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":4777
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   8
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":4989
            MousePointer    =   99  'Custom
            TabIndex        =   21
            Top             =   345
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   7
            Left            =   2355
            MouseIcon       =   "ctrCalendar.ctx":4B9B
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   6
            Left            =   1995
            MouseIcon       =   "ctrCalendar.ctx":4DAD
            MousePointer    =   99  'Custom
            TabIndex        =   19
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   5
            Left            =   1620
            MouseIcon       =   "ctrCalendar.ctx":4FBF
            MousePointer    =   99  'Custom
            TabIndex        =   18
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   4
            Left            =   1245
            MouseIcon       =   "ctrCalendar.ctx":51D1
            MousePointer    =   99  'Custom
            TabIndex        =   17
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   3
            Left            =   870
            MouseIcon       =   "ctrCalendar.ctx":53E3
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   2
            Left            =   495
            MouseIcon       =   "ctrCalendar.ctx":55F5
            MousePointer    =   99  'Custom
            TabIndex        =   15
            Top             =   45
            Width           =   240
         End
         Begin VB.Label l 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   220
            Index           =   1
            Left            =   135
            MouseIcon       =   "ctrCalendar.ctx":5807
            MousePointer    =   99  'Custom
            TabIndex        =   14
            Top             =   45
            Width           =   240
         End
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   285
         Left            =   15
         TabIndex        =   5
         Top             =   105
         Width           =   360
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   285
         Left            =   2385
         TabIndex        =   4
         Top             =   105
         Width           =   345
      End
      Begin VB.TextBox txtYear 
         Alignment       =   2  'Center
         BackColor       =   &H00DA7C58&
         BorderStyle     =   0  'None
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
         Height          =   195
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "November, 2003"
         Top             =   135
         Width           =   2445
      End
      Begin VB.CommandButton cmdCurrentDate 
         Caption         =   "Current Date"
         Height          =   315
         Left            =   15
         TabIndex        =   2
         Top             =   2610
         Width           =   2700
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00DA7C58&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   15
         ScaleHeight     =   285
         ScaleWidth      =   3000
         TabIndex        =   56
         Top             =   105
         Width           =   3000
      End
      Begin VB.Line Line1 
         X1              =   15
         X2              =   2700
         Y1              =   645
         Y2              =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   150
         TabIndex        =   12
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Mon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   495
         TabIndex        =   11
         Top             =   435
         Width           =   270
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   870
         TabIndex        =   10
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Wed"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   1245
         TabIndex        =   9
         Top             =   435
         Width           =   270
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Thu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   1620
         TabIndex        =   8
         Top             =   435
         Width           =   240
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Fri"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   2040
         TabIndex        =   7
         Top             =   435
         Width           =   180
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Sat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00DA7C58&
         Height          =   165
         Left            =   2385
         TabIndex        =   6
         Top             =   435
         Width           =   210
      End
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   165
      TabIndex        =   13
      Top             =   4080
      Width           =   2730
   End
End
Attribute VB_Name = "ctrCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Dim ADDINDX As Integer
Dim LASTDAY As Integer
Dim SETDAY As Integer
Dim Today As String
Dim SETDATE As Date
Dim SETMONTH As Date
Dim textDate As Date
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."

Private Sub cmdCurrentDate_Click()
    SETDATE = Date
    SETMONTH = Date
    Today = Day(SETDATE)
    Call Calender
    Call CurrentDay
    txtYear.Text = Format(SETMONTH, "MMMM, YYYY")
End Sub

Private Sub cmdNext_Click()
    SETMONTH = SETMONTH + 32
    Call Calender
    Call CurrentDay
    txtYear.Text = Format(SETMONTH, "MMMM, YYYY")
End Sub

Private Sub cmdPrevious_Click()
    SETMONTH = SETMONTH - 32
    Call Calender
    Call CurrentDay
    txtYear.Text = Format(SETMONTH, "MMMM, YYYY")
End Sub

Private Sub l_DblClick(Index As Integer)
    RaiseEvent DblClick
End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo chngtxt
    Dim intLocation As Integer
    Dim CHNGETXT
    Dim n
    
    
    Select Case KeyCode
        Case vbKeyUp
    If ADDINDX < SETDAY Then Exit Sub
            ADDINDX = ADDINDX - 7
                        
                If Len(l(ADDINDX).Caption) > 0 Then
                        For n = l.LBound To l.UBound
                            l(n).ForeColor = &H0&          ' vbBlack
                            l(n).BackColor = &H80000005
                                
                        Next
                        
    
                            l(ADDINDX).ForeColor = &H8000000E               ' vbRed
                            l(ADDINDX).BackColor = &H8000000D         ' vbYellow
                End If
        Case vbKeyDown
    If ADDINDX > LASTDAY Then Exit Sub
            ADDINDX = ADDINDX + 7
                If Len(l(ADDINDX).Caption) > 0 Then

                        For n = l.LBound To l.UBound
                            l(n).ForeColor = &H0&          ' vbBlack
                            l(n).BackColor = &H80000005
                            l(n).Enabled = True

                        Next

                            l(ADDINDX).ForeColor = &H8000000E               ' vbRed
                            l(ADDINDX).BackColor = &H8000000D         ' vbYellow
                End If
    
        Case vbKeyLeft
    If ADDINDX < SETDAY Then Exit Sub
            ADDINDX = ADDINDX - 1
                If Len(l(ADDINDX).Caption) > 0 Then

                         For n = l.LBound To l.UBound
                             l(n).ForeColor = &H0&          ' vbBlack
                             l(n).BackColor = &H80000005
                             l(n).Enabled = True

                         Next

                             l(ADDINDX).ForeColor = &H8000000E               ' vbRed
                             l(ADDINDX).BackColor = &H8000000D         ' vbYellow
                 End If
        Case vbKeyRight
    If ADDINDX > LASTDAY Then Exit Sub
            ADDINDX = ADDINDX + 1
                If Len(l(ADDINDX).Caption) > 0 Then

                        For n = l.LBound To l.UBound
                            l(n).ForeColor = &H0&          ' vbBlack
                            l(n).BackColor = &H80000005
                            l(n).Enabled = True

                        Next

                            l(ADDINDX).ForeColor = &H8000000E               ' vbRed
                            l(ADDINDX).BackColor = &H8000000D         ' vbYellow

                End If
    
        Case vbKeyPageDown
            SETMONTH = SETMONTH + 32
            Call Calender
            Call CurrentDay
            txtYear.Text = Format(SETMONTH, "MMMM, YYYY")
    Picture1.SetFocus
        Case vbKeyPageUp
            SETMONTH = SETMONTH - 32
            Call Calender
            Call CurrentDay
            txtYear.Text = Format(SETMONTH, "MMMM, YYYY")
            Picture1.SetFocus
        End Select
            textDate = Month(SETMONTH) & "/" & l(ADDINDX).Caption & "/" & Year(SETMONTH)
            txtDate = textDate
            SDate = textDate
    Exit Sub
chngtxt:
End Sub

Private Sub UserControl_Initialize()
    Call cmdCurrentDate_Click
    cmdCurrentDate.Caption = Date
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = Frame1.Width
    UserControl.Height = Frame1.Height - 100
End Sub

Private Sub Calender()
On Error Resume Next
Dim m
Dim Y As Integer

Dim Days As Integer
Dim aLOOP As Integer
Dim WDay As String

Dim setMaxDay As Integer
Dim setChkDay As Integer

SETDATE = (SETMONTH - Day(SETMONTH)) + 1

WDay = Format(SETDATE, "DDDD")
If WDay = "Sunday" Then SETDAY = 1
If WDay = "Monday" Then SETDAY = 2
If WDay = "Tuesday" Then SETDAY = 3
If WDay = "Wednesday" Then SETDAY = 4
If WDay = "Thursday" Then SETDAY = 5
If WDay = "Friday" Then SETDAY = 6
If WDay = "Saturday" Then SETDAY = 7
Days = 0



m = Month(SETMONTH)
Y = (Year(SETMONTH) Mod 4)

If m = 1 Or m = 3 Or m = 5 Or m = 7 Or m = 8 Or m = 10 Or m = 12 Then
setMaxDay = 31
ElseIf (m = 4 Or m = 6 Or m = 9 Or m = 11) Then 'For April,June,September,November count 30 days
setMaxDay = 30
ElseIf m = 2 And Y = 0 Then
setMaxDay = 29
ElseIf m = 2 And Y <> 0 Then
setMaxDay = 28
End If
setChkDay = setMaxDay + SETDAY

    For aLOOP = 1 To l.Count
        l(aLOOP).Caption = ""
    Next aLOOP

For aLOOP = SETDAY To setChkDay
Days = Days + 1
    l(aLOOP).Caption = Days

If Days = setMaxDay Then Exit For
Next
LASTDAY = aLOOP
'MsgBox LASTDAY
End Sub
Private Sub CurrentDay()
On Error Resume Next
Dim ChkCurrentDay As Integer
           
           
           Dim n
            For n = l.LBound To l.UBound
                If l(n).Caption = Today Then
                    l(n).ForeColor = &HFF&      ' &H8000000E          ' vbRed
                    l(n).BackColor = &HC0FFC0     '&H8000000D         ' vbYellow
                    'l(n).FontBold = True
                    If l(n) = RestoreDate(Month(SETMONTH) & "/" & l(n).Caption & "/" & Year(SETMONTH)) = True Then
                        l(n).FontBold = True
                        l(n).ForeColor = &HFF0000
                    Else
                        l(n).FontBold = False
                    End If
                        textDate = Month(SETMONTH) & "/" & l(n).Caption & "/" & Year(SETMONTH)
                        txtDate = textDate
                Else
                    l(n).ForeColor = &H0&          ' vbBlack
                    l(n).BackColor = &H80000005
                    If RestoreDate(Month(SETMONTH) & "/" & l(n).Caption & "/" & Year(SETMONTH)) = True Then
                        l(n).FontBold = True
                        l(n).ForeColor = &HFF0000
                    Else
                        l(n).FontBold = False
                    End If
                End If
            Next


End Sub

Private Sub l_Click(Index As Integer)
On Error Resume Next
Dim n
ADDINDX = Index
If Len(l(Index).Caption) > 0 Then

For n = l.LBound To l.UBound
    l(n).ForeColor = &H0&          ' vbBlack
    If RestoreDate(Month(SETMONTH) & "/" & l(n).Caption & "/" & Year(SETMONTH)) = True Then
        l(n).FontBold = True
        l(n).ForeColor = &HFF0000
    Else
        l(n).FontBold = False
    End If
    l(n).BackColor = &H80000005
Next

    l(Index).ForeColor = &HFF& '&H8000000E          ' vbRed
    'l(Index).FontBold = True
    l(Index).BackColor = &HC0FFC0 '&H8000000D         ' vbYellow
End If
textDate = Month(SETMONTH) & "/" & l(Index).Caption & "/" & Year(SETMONTH)
txtDate = textDate
SDate = textDate
Picture1.SetFocus
RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
     
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtDate,txtDate,-1,Text
Public Property Get SDate() As String
Attribute SDate.VB_Description = "Returns/sets the text contained in the control."
    SDate = txtDate.Text
End Property

Public Property Let SDate(ByVal New_SDate As String)
    txtDate.Text() = New_SDate
    PropertyChanged "SDate"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    txtDate.Text = PropBag.ReadProperty("SDate", "")
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("SDate", txtDate.Text, "")
End Sub
Public Function RestoreDate(SDate As String) As Boolean
If Not Date = "" Then
    'MsgBox Date
End If
Dim DIYAINI As New DIYAINI
    With DIYAINI
        .path = App.path & "\Settings.ini"
        .Section = "AvailDate"
        .Key = SDate
        If .Value = "Yes" Then
            RestoreDate = True
        Else
            RestoreDate = False
        End If
    End With
End Function

