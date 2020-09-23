VERSION 5.00
Begin VB.UserControl ShapeFrame 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Shape SP 
      Height          =   1470
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "ShapeFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Sub UserControl_Resize()
    SP.Width = UserControl.Width
    SP.Height = UserControl.Height
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SP,SP,-1,FillColor
Public Property Get FillColor() As OLE_COLOR
Attribute FillColor.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    FillColor = SP.FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    SP.FillColor() = New_FillColor
    PropertyChanged "FillColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SP,SP,-1,FillStyle
Public Property Get FillStyle() As Integer
Attribute FillStyle.VB_Description = "Returns/sets the fill style of a shape."
    FillStyle = SP.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As Integer)
    SP.FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SP,SP,-1,BorderColor
Public Property Get BorderColor() As Long
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = SP.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As Long)
    SP.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SP,SP,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = SP.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    SP.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=SP,SP,-1,Shape
Public Property Get Shape() As Integer
Attribute Shape.VB_Description = "Returns/sets a value indicating the appearance of a control."
    Shape = SP.Shape
End Property

Public Property Let Shape(ByVal New_Shape As Integer)
    SP.Shape() = New_Shape
    PropertyChanged "Shape"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SP.FillColor = PropBag.ReadProperty("FillColor", &H0&)
    SP.FillStyle = PropBag.ReadProperty("FillStyle", 1)
    SP.BorderColor = PropBag.ReadProperty("BorderColor", -2147483640)
    SP.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    SP.Shape = PropBag.ReadProperty("Shape", 4)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("FillColor", SP.FillColor, &H0&)
    Call PropBag.WriteProperty("FillStyle", SP.FillStyle, 1)
    Call PropBag.WriteProperty("BorderColor", SP.BorderColor, -2147483640)
    Call PropBag.WriteProperty("BorderStyle", SP.BorderStyle, 1)
    Call PropBag.WriteProperty("Shape", SP.Shape, 4)
End Sub

