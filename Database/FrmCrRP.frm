VERSION 5.00
Begin VB.Form FrmCrRP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diyas Database Management Wizard."
   ClientHeight    =   8100
   ClientLeft      =   930
   ClientTop       =   1230
   ClientWidth     =   11550
   Icon            =   "FrmCrRP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
         Height          =   2295
         Left            =   945
         TabIndex        =   4
         Top             =   1620
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   4048
         FillColor       =   16777215
         FillStyle       =   0
         BorderColor     =   13020333
         Begin VB.TextBox TxtNme 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   300
            TabIndex        =   8
            Top             =   450
            Width           =   6300
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "This restore point cannot be changed after it is created. Before continuing, ensure that you have typed the correct name."
            Height          =   510
            Left            =   315
            TabIndex        =   10
            Top             =   1575
            Width           =   6900
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The current date and time are automatically added to your restore point."
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   1095
            Width           =   5040
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Restore point description:"
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
            Left            =   285
            TabIndex        =   7
            Top             =   165
            Width           =   2700
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmCrRP.frx":9A1A
         Height          =   810
         Left            =   660
         TabIndex        =   3
         Top             =   1080
         Width           =   8070
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmCrRP.frx":9AD1
         Height          =   795
         Left            =   660
         TabIndex        =   2
         Top             =   240
         Width           =   8085
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
         Caption         =   "Create a Restore Point."
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
         Height          =   360
         Left            =   1335
         TabIndex        =   1
         Top             =   315
         Width           =   3495
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   30
         Picture         =   "FrmCrRP.frx":9BCD
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Create a Restore Point."
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
         Width           =   3495
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
      Begin Project1.DiyaButton CmdBack 
         Height          =   285
         Left            =   6825
         TabIndex        =   11
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "< &Back"
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
         MICON           =   "FrmCrRP.frx":AED9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.DiyaButton CmdCancel 
         Height          =   285
         Left            =   9480
         TabIndex        =   6
         Top             =   105
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
         MICON           =   "FrmCrRP.frx":AEF5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.DiyaButton CmdCreate 
         Height          =   285
         Left            =   8175
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "C&reate"
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
         MICON           =   "FrmCrRP.frx":AF11
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
Attribute VB_Name = "FrmCrRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DIYAButton1_Click()

End Sub

Private Sub DIYAButton3_Click()
    
End Sub

Private Sub CmdBack_Click()
    FrmMain.Left = Me.Left
    FrmMain.Top = Me.Top
    FrmMain.Show
    Unload Me
End Sub

Private Sub CmdCancel_Click()
    End
End Sub

Private Sub CmdCreate_Click()
If Not TxtNme.Text = "" Then
    FrmCrRPF.Left = Me.Left
    FrmCrRPF.Top = Me.Top
    FrmCrRPF.Show
    FrmCrRPF.TxtNme.Text = TxtNme.Text
    Unload Me
Else
    MsgBox "Please specify a description for Restore Point", vbInformation, "Database Management"
End If
End Sub

Private Sub Form_Paint()
    ProgressBar1.ToColor = vbActiveTitleBar
    ProgressBar2.ToColor = vbActiveTitleBar
    ProgressBar3.ToColor = vbActiveTitleBar
    ProgressBar4.ToColor = vbActiveTitleBar
End Sub
