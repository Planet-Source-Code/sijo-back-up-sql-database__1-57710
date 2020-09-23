VERSION 5.00
Begin VB.Form FrmOpt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diyas Database Management Wizard."
   ClientHeight    =   8100
   ClientLeft      =   945
   ClientTop       =   1215
   ClientWidth     =   11550
   Icon            =   "FrmOpts.frx":0000
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
      Begin VB.TextBox TxtPath 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   3030
         MousePointer    =   1  'Arrow
         TabIndex        =   4
         Text            =   "C:\BUP\DBS"
         Top             =   1095
         Width           =   6885
      End
      Begin Project1.ProgressBar ProgressBar5 
         Height          =   4830
         Left            =   885
         Top             =   855
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   8520
         ToColor         =   14208459
         Value           =   100
         BorderStyle     =   1
         Orientation     =   3
         BackColor       =   14208459
         Begin VB.Image Image2 
            Height          =   2790
            Left            =   180
            Picture         =   "FrmOpts.frx":9A1A
            Stretch         =   -1  'True
            Top             =   915
            Width           =   2400
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Store Back Up files in :"
         Height          =   195
         Left            =   1155
         TabIndex        =   3
         Top             =   1110
         Width           =   1620
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
         Caption         =   "Change Settings."
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
         Width           =   2595
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   30
         Picture         =   "FrmOpts.frx":EE58
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change Settings."
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
         Width           =   2595
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
      Begin Project1.DiyaButton CmdClose 
         Height          =   285
         Left            =   9480
         TabIndex        =   2
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Close"
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
         MICON           =   "FrmOpts.frx":10164
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
Attribute VB_Name = "FrmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdClose_Click()
    FrmMain.Left = Me.Left
    FrmMain.Top = Me.Top
    FrmMain.Show
    Unload Me
End Sub

Private Sub Form_Paint()
    ProgressBar1.ToColor = vbActiveTitleBar
    ProgressBar2.ToColor = vbActiveTitleBar
    ProgressBar3.ToColor = vbActiveTitleBar
    ProgressBar4.ToColor = vbActiveTitleBar
    ProgressBar5.ToColor = vbActiveTitleBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim DIYAINI As New DIYAINI
        With DIYAINI
            .path = App.path & "\Settings.ini"
            .Section = "DataBase"
            .Key = "BPath"
            .Value = TxtPath.Text
        End With
End Sub

