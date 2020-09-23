VERSION 5.00
Begin VB.UserControl XPOption 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2520
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   1035
   ScaleWidth      =   2520
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   375
      Top             =   570
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      BorderStyle     =   3  'Dot
      DrawMode        =   6  'Mask Pen Not
      FillColor       =   &H00404040&
      Height          =   255
      Left            =   240
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image imgNone 
      Height          =   195
      Left            =   0
      Picture         =   "XPOption.ctx":0000
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgDown 
      Height          =   195
      Left            =   840
      Picture         =   "XPOption.ctx":00D8
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCheckedDown 
      Height          =   195
      Left            =   1620
      Picture         =   "XPOption.ctx":01B3
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgCheckedOver 
      Height          =   195
      Left            =   1230
      Picture         =   "XPOption.ctx":0310
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgChecked 
      Height          =   195
      Left            =   1425
      Picture         =   "XPOption.ctx":053E
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image imgOver 
      Height          =   195
      Left            =   1035
      Picture         =   "XPOption.ctx":0696
      Top             =   390
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "XP Option Button"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1230
   End
   Begin VB.Image imgCheckBox 
      Height          =   195
      Left            =   0
      Picture         =   "XPOption.ctx":0776
      Top             =   0
      Width           =   195
   End
End
Attribute VB_Name = "XPOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_Value = 0
'Property Variables:
Dim m_Value As Boolean


Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type

Dim NewCur As POINTAPI
Dim OldCur As POINTAPI
Dim Mousedown As Boolean

Event ValueChanged(blnValue As Boolean)
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    Label1.Enabled = UserControl.Enabled
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Value() As Boolean
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Boolean)
    m_Value = New_Value
    PropertyChanged "Value"
    If m_Value = True Then
    imgCheckBox = imgChecked
    Else
    imgCheckBox = imgNone
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Label1,Label1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Label1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Label1.Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub imgCheckBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub imgCheckBox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub


Private Sub imgCheckBox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub


Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub Timer1_Timer()
If Mousedown = True Then Exit Sub
GetCursorPos NewCur
If NewCur.X = OldCur.X And NewCur.Y = OldCur.Y Then

Else
If m_Value = True Then
imgCheckBox = imgChecked
Else
imgCheckBox = imgNone
End If
Timer1.Enabled = False
End If

End Sub

Private Sub UserControl_GotFocus()
'Shape1.Visible = True
Shape1.Top = Label1.Top - 2
Shape1.Left = Label1.Left - 2
Shape1.Width = Label1.Width + 4
Shape1.Height = Label1.Height + 4
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Value = m_def_Value
End Sub

Private Sub UserControl_LostFocus()
Shape1.Visible = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mousedown = True
If m_Value = True Then
imgCheckBox = imgCheckedDown
Else
imgCheckBox = imgDown
End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Mousedown = True Then Exit Sub

If m_Value = True Then
imgCheckBox = imgCheckedOver
Else
imgCheckBox = imgOver
End If
GetCursorPos OldCur
Timer1.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Mousedown = False
m_Value = Not m_Value
RaiseEvent ValueChanged(m_Value)
If m_Value = True Then
imgCheckBox = imgChecked
Else
imgCheckBox = imgNone
End If

End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Label1.Caption = PropBag.ReadProperty("Caption", "XP Check Box")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    Set Label1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    If m_Value = True Then
    imgCheckBox = imgChecked
    Else
    imgCheckBox = imgNone
    End If
    Label1.Enabled = UserControl.Enabled
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 195
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", Label1.Caption, "XP Check Box")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("Font", Label1.Font, Ambient.Font)
End Sub

