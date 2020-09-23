VERSION 5.00
Begin VB.Form FrmCrRPF 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Diyas Database Management Wizard."
   ClientHeight    =   8100
   ClientLeft      =   930
   ClientTop       =   1230
   ClientWidth     =   11550
   Icon            =   "FrmCrRPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11550
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5175
      Top             =   3810
   End
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
      Begin VB.FileListBox FL 
         Height          =   2235
         Left            =   570
         TabIndex        =   8
         Top             =   4005
         Visible         =   0   'False
         Width           =   2190
      End
      Begin Project1.ShapeFrame ShapeFrame1 
         Height          =   2295
         Left            =   1800
         TabIndex        =   2
         Top             =   1620
         Width           =   7440
         _ExtentX        =   13123
         _ExtentY        =   4048
         FillColor       =   16777215
         FillStyle       =   0
         BorderColor     =   13020333
         Begin Project1.XP_ProgressBar XPPG 
            Height          =   270
            Left            =   1755
            TabIndex        =   6
            Top             =   975
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   476
            Color           =   16561022
            Scrolling       =   1
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Please wait Database Management Creating a Restore Point.."
            Height          =   195
            Left            =   1515
            TabIndex        =   5
            Top             =   1620
            Width           =   4380
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Creating Restore Point."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   285
            TabIndex        =   4
            Top             =   165
            Width           =   1995
         End
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
         Caption         =   "Creating Restore Point."
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
         Width           =   3465
      End
      Begin VB.Image Image1 
         Height          =   975
         Left            =   30
         Picture         =   "FrmCrRPF.frx":9A1A
         Top             =   75
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Creating Restore Point."
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
         Width           =   3510
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
      Begin VB.TextBox TxtNme 
         Height          =   345
         Left            =   255
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   75
         Visible         =   0   'False
         Width           =   2760
      End
      Begin Project1.DiyaButton CmdFinish 
         Height          =   285
         Left            =   9480
         TabIndex        =   3
         Top             =   90
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "&Finish"
         ENAB            =   0   'False
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
         MICON           =   "FrmCrRPF.frx":AD26
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
Attribute VB_Name = "FrmCrRPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFinish_Click()
    End
End Sub

Private Sub Form_Activate()
Dim DIYAINI As New DIYAINI
    With DIYAINI
        .path = App.path & "\Settings.ini"
        .Section = "AvailDate"
        .Key = Date
        .Value = "Yes"
        .Section = "RPoint" & Date
        .Key = Time & " " & TxtNme.Text
        .Value = Month(Date) & Day(Date) & Year(Date) & TxtNme.Text & Hour(Time) & Minute(Time) & Second(Time) & Right(Time, 2)
    End With
    BackUpDB
    CmdFinish.Enabled = True
End Sub

Private Sub Form_Load()
    FL.path = App.path
End Sub

Private Sub Form_Paint()
    ProgressBar1.ToColor = vbActiveTitleBar
    ProgressBar2.ToColor = vbActiveTitleBar
    ProgressBar3.ToColor = vbActiveTitleBar
    ProgressBar4.ToColor = vbActiveTitleBar
End Sub

Private Sub Timer1_Timer()
If XPPG.Value < XPPG.Max + 5 Then
    XPPG.Value = XPPG.Value + 1
End If
End Sub

Private Sub BackUpDB()
KillReport
MakeEnvironment ' Only For Source code viewers
'With Dlg
'.DialogTitle = "Back Up Database"
'.Filter = "ZIP Files|*.ZIP"
'.ShowSave
'TMPL = Mid(Dlg.FileTitle, 1, Len(Dlg.FileTitle) - 4)
'DDT = Left(Dlg.FileName, 3)
FNME = Month(Date) & Day(Date) & Year(Date) & TxtNme.Text & Hour(Time) & Minute(Time) & Second(Time) & Right(Time, 2) & ".CAB"
Dim DIYAINI As New DIYAINI
    With DIYAINI
        .path = App.path & "\Settings.ini"
        .Section = "DataBase"
        .Key = "BPath"
        BUPPath = .Value
    End With
    '------
        Dim SZIP As String
        SZIP = ".Option EXPLICIT" _
        & vbCrLf & ".Set MaxDiskSize = CDRom" _
        & vbCrLf & ".Set ReservePerCabinetSize = 44" _
        & vbCrLf & ".Set RptFileName=" & App.path & "\MKCB.rpt" _
        & vbCrLf & ".Set DiskDirectoryTemplate =" & Chr(34) & BUPPath & Chr(34) _
        & vbCrLf & ".Set CompressionType = MSZIP" _
        & vbCrLf & ".Set CompressionLevel = 7" _
        & vbCrLf & ".Set CompressionMemory = 21" _
        & vbCrLf & ".Set CabinetNameTemplate =" & Chr(34) & FNME & Chr(34) _
        & vbCrLf & ".Set Cabinet=on" _
        & vbCrLf & ".Set Compress=on" _
        '& vbCrLf & "Settings.ini"
        XPPG.Max = FL.ListCount - 1
        For a = 0 To FL.ListCount - 1
            XPPG.Value = a
            FL.ListIndex = a
            SZIP = SZIP & vbCrLf & Chr(34) & FL.FileName & Chr(34)
        Next
    Open App.path & "\DBZIP.SED" For Output As #1
        Print #1, SZIP
    Close #1
    Shell App.path & "\SSCB.exe /f DBZIP.SED", vbHide
    MsgBox "Restore Point has been successfully created.", vbInformation
End Sub

Private Sub MakeEnvironment()
On Error Resume Next
MkDir "C:\BUP"
MsgBox "A Dir Created at C: as BUP"
MkDir "C:\BUP\DBS"
MsgBox "A Dir Created at C:\BUP as DBS"
End Sub
