VERSION 5.00
Begin VB.Form FrmRestore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diyas Database Management Wizard."
   ClientHeight    =   8100
   ClientLeft      =   915
   ClientTop       =   1245
   ClientWidth     =   11550
   Icon            =   "FrmRStore.frx":0000
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
         Height          =   3840
         Left            =   945
         TabIndex        =   4
         Top             =   1620
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   6773
         FillColor       =   16777215
         FillStyle       =   0
         BorderColor     =   13020333
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   2865
            Left            =   3900
            ScaleHeight     =   2835
            ScaleWidth      =   3675
            TabIndex        =   9
            Top             =   660
            Width           =   3705
            Begin VB.ListBox LstRestore 
               Appearance      =   0  'Flat
               Height          =   2565
               ItemData        =   "FrmRStore.frx":9A1A
               Left            =   -15
               List            =   "FrmRStore.frx":9A21
               TabIndex        =   14
               Top             =   285
               Width           =   3705
            End
            Begin VB.CommandButton cmdNext 
               Caption         =   ">>"
               Height          =   285
               Left            =   3330
               TabIndex        =   11
               Top             =   0
               Width           =   345
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00DA7C58&
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   360
               ScaleHeight     =   285
               ScaleWidth      =   2970
               TabIndex        =   12
               Top             =   0
               Width           =   2970
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
                  Height          =   285
                  Left            =   0
                  Locked          =   -1  'True
                  TabIndex        =   13
                  Text            =   "November, 2003"
                  Top             =   30
                  Width           =   2970
               End
            End
            Begin VB.CommandButton cmdPrevious 
               Caption         =   "<<"
               Height          =   285
               Left            =   0
               TabIndex        =   10
               Top             =   0
               Width           =   360
            End
         End
         Begin Project1.ctrCalendar CAL 
            Height          =   2835
            Left            =   360
            TabIndex        =   8
            Top             =   675
            Width           =   2730
            _ExtentX        =   4815
            _ExtentY        =   5001
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SDate           =   "12/08/2004"
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "2. On this list click a restore point."
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
            Left            =   3915
            TabIndex        =   16
            Top             =   375
            Width           =   3720
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "1. On this calendar, click a bold date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   330
            TabIndex        =   15
            Top             =   375
            Width           =   3465
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRStore.frx":9A31
         Height          =   810
         Left            =   660
         TabIndex        =   3
         Top             =   1080
         Width           =   8070
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"FrmRStore.frx":9ACA
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
         Caption         =   "Restore Database."
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
         Width           =   2790
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   30
         Picture         =   "FrmRStore.frx":9B57
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Restore Database."
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
         Width           =   2790
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
         TabIndex        =   7
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
         MICON           =   "FrmRStore.frx":AE63
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
         MICON           =   "FrmRStore.frx":AE7F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Project1.DiyaButton CmdRestore 
         Height          =   285
         Left            =   8175
         TabIndex        =   5
         Top             =   90
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Restore"
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
         MICON           =   "FrmRStore.frx":AE9B
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
Attribute VB_Name = "FrmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CAL_Click()
If CAL.RestoreDate(CAL.SDate) = False Then
    LstRestore.Clear
    LstRestore.AddItem "There is no Restore Point on Selected Date"
Else
    LstRestore.Clear
    Dim DIYAINI As New DIYAINI
    Dim strKeys() As String
    Dim lonKeyCount As Long
    Dim lonCurrentKey As Long
        With DIYAINI
            .path = App.path & "\Settings.ini"
            .Section = "RPoint" & CAL.SDate
            .EnumerateCurrentSection strKeys(), lonKeyCount
            For lonCurrentKey = 1 To lonKeyCount
                .Key = strKeys(lonCurrentKey)
                LstRestore.AddItem .Key '.Value
            Next lonCurrentKey
        End With
End If
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

Private Sub CmdRestore_Click()
    RestoreDB
End Sub

Private Sub Form_Load()
    txtYear.Text = MonthName(Month(CAL.SDate)) & ", " & Year(CAL.SDate)
    CAL_Click
End Sub

Private Sub SIYAButton2_Click()

End Sub

Private Sub Form_Paint()
    ProgressBar1.ToColor = vbActiveTitleBar
    ProgressBar2.ToColor = vbActiveTitleBar
    ProgressBar3.ToColor = vbActiveTitleBar
    ProgressBar4.ToColor = vbActiveTitleBar
End Sub
