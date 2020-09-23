VERSION 5.00
Begin VB.Form FrmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diyas Database Management Wizard."
   ClientHeight    =   8100
   ClientLeft      =   930
   ClientTop       =   1230
   ClientWidth     =   11550
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11550
   Begin Project1.ProgressBar ProgressBar3 
      Height          =   6480
      Left            =   240
      Top             =   1095
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   11430
      ToColor         =   14208459
      Value           =   100
      BorderStyle     =   0
      Orientation     =   2
      BackColor       =   15527148
      Begin Project1.ShapeFrame ShapeFrame1 
         Height          =   1740
         Left            =   6165
         TabIndex        =   6
         Top             =   4290
         Width           =   4410
         _ExtentX        =   7779
         _ExtentY        =   3069
         FillColor       =   16777215
         FillStyle       =   0
         BorderColor     =   13020333
         Begin Project1.XPOption Opt3 
            Height          =   195
            Left            =   795
            TabIndex        =   9
            Top             =   885
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   344
            Caption         =   "Cr&eate a Restore Point."
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
         Begin Project1.XPOption Opt1 
            Height          =   195
            Left            =   795
            TabIndex        =   7
            Top             =   225
            Width           =   2190
            _ExtentX        =   3863
            _ExtentY        =   344
            Caption         =   "View / Al&ter My Database."
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
         Begin Project1.XPOption Opt2 
            Height          =   195
            Left            =   795
            TabIndex        =   8
            Top             =   555
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   344
            Caption         =   "&Restore My Database to an earlier time."
            Value           =   -1  'True
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
         Begin Project1.XPOption Opt4 
            Height          =   195
            Left            =   810
            TabIndex        =   12
            Top             =   1230
            Width           =   2655
            _ExtentX        =   4683
            _ExtentY        =   344
            Caption         =   "Change Settings"
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
      End
      Begin Project1.ProgressBar ProgressBar5 
         Height          =   4830
         Left            =   1320
         Top             =   780
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   8520
         ToColor         =   14208459
         Value           =   100
         BorderStyle     =   1
         Orientation     =   3
         BackColor       =   14208459
         Begin VB.Image Image2 
            Height          =   1920
            Left            =   390
            Picture         =   "FrmMain.frx":9A1A
            Top             =   1350
            Width           =   1920
         End
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmMain.frx":B3EB
         Height          =   810
         Left            =   6150
         TabIndex        =   5
         Top             =   2955
         Width           =   4590
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmMain.frx":B4D6
         Height          =   810
         Left            =   6150
         TabIndex        =   4
         Top             =   2010
         Width           =   4590
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C6ACAD&
         X1              =   5535
         X2              =   5535
         Y1              =   225
         Y2              =   6300
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmMain.frx":B5C1
         Height          =   810
         Left            =   6150
         TabIndex        =   3
         Top             =   1080
         Width           =   4590
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Database Management Wizard Assists you to perform all kind of basic Database Administration tasks."
         Height          =   480
         Left            =   6150
         TabIndex        =   2
         Top             =   465
         Width           =   4590
      End
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   1050
      Left            =   0
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1852
      ToColor         =   14208459
      Value           =   100
      BorderStyle     =   0
      Orientation     =   0
      BackColor       =   15527148
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Database Management Wizard."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   1335
         TabIndex        =   1
         Top             =   315
         Width           =   6570
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   30
         Picture         =   "FrmMain.frx":B6B9
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Database Management Wizard."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C6ACAD&
         Height          =   360
         Left            =   1365
         TabIndex        =   0
         Top             =   330
         Width           =   6585
      End
   End
   Begin Project1.ProgressBar ProgressBar2 
      Height          =   495
      Left            =   0
      Top             =   7605
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   873
      ToColor         =   13020333
      Value           =   100
      BorderStyle     =   0
      Orientation     =   3
      BackColor       =   15527148
      Begin Project1.DiyaButton CmdCancel 
         Height          =   285
         Left            =   9480
         TabIndex        =   11
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":C9C5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.DiyaButton CmdNext 
         Height          =   285
         Left            =   8115
         TabIndex        =   10
         Top             =   90
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Next >"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmMain.frx":C9E1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin Project1.ProgressBar ProgressBar4 
      Height          =   6480
      Left            =   -60
      Top             =   1095
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   11430
      ToColor         =   14208459
      Value           =   100
      BorderStyle     =   0
      BackColor       =   15527148
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   11550
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   11550
      Y1              =   7590
      Y2              =   7590
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    End
End Sub

Private Sub cmdNext_Click()
    If Opt1.Value = True Then
        MsgBox "Unavailable"
    ElseIf Opt2.Value = True Then
        FrmRestore.Left = Me.Left
        FrmRestore.Top = Me.Top
        FrmRestore.Show
        Unload Me
    ElseIf Opt3.Value = True Then
        FrmCrRP.Left = Me.Left
        FrmCrRP.Top = Me.Top
        FrmCrRP.Show
        Unload Me
    ElseIf Opt4.Value = True Then
        FrmOpt.Left = Me.Left
        FrmOpt.Top = Me.Top
        FrmOpt.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
If Command$ = "BackUp" Then
    AutoRestore ("Auto" & Command$)
End If
'ONLY FOR PSC | to create sscb.exe
'Because PSC will delete it from my ZIP file
'It is Windows Cabinet maker found Package & Deployment directory as Makecab.exe
Dim SEXE() As Byte
FileNumber = FreeFile
Open App.path & "\SSCB.EXE" For Binary As FileNumber
    SEXE = LoadResData(101, "CUSTOM")
    Put #FileNumber, , SEXE()
Close #FileNumber
End Sub

Private Sub Form_Paint()
    ProgressBar1.ToColor = vbActiveTitleBar
    ProgressBar2.ToColor = vbActiveTitleBar
    ProgressBar3.ToColor = vbActiveTitleBar
    ProgressBar4.ToColor = vbActiveTitleBar
    ProgressBar5.ToColor = vbActiveTitleBar
End Sub

Private Sub Opt1_ValueChanged(blnValue As Boolean)
If Opt1.Value = True Then
    Opt2.Value = False
    Opt3.Value = False
    Opt4.Value = False
Else
    Opt1.Value = True
End If
End Sub

Private Sub Opt2_ValueChanged(blnValue As Boolean)
If Opt2.Value = True Then
    Opt1.Value = False
    Opt3.Value = False
    Opt4.Value = False
Else
    Opt2.Value = True
End If
End Sub

Private Sub Opt3_ValueChanged(blnValue As Boolean)
If Opt3.Value = True Then
    Opt1.Value = False
    Opt2.Value = False
    Opt4.Value = False
Else
    Opt3.Value = True
End If
End Sub

Private Sub Opt4_ValueChanged(blnValue As Boolean)
If Opt4.Value = True Then
    Opt1.Value = False
    Opt2.Value = False
    Opt3.Value = False
Else
    Opt4.Value = True
End If
End Sub
