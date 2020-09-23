VERSION 5.00
Begin VB.UserControl UserControl1 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ScaleHeight     =   1470
   ScaleWidth      =   1500
   Begin VB.Label lblDeg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   900
      TabIndex        =   0
      Top             =   60
      Width           =   90
   End
   Begin VB.Shape Center 
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   780
      Shape           =   3  'Circle
      Top             =   840
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   60
      X2              =   1320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   60
      X2              =   1320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Border 
      Height          =   495
      Left            =   120
      Top             =   840
      Width           =   1215
   End
   Begin VB.Shape Outer 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1020
      Width           =   135
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim ANGLE As Integer
Const Pi As Single = 3.1416
Dim Circ As Double
Dim Unit As Double
Dim Radian As Double

Sub MoveBorder()
    With Border
        .Top = 0
        .Left = 0
        .Width = Width
        .Height = Height
    End With
End Sub

Sub CreateCenter()
    With Center
        .Top = (Height / 2) - (.Height / 2)
        .Left = (Width / 2) - (.Width / 2)
    End With

End Sub

Sub MoldOuter()
    Dim W As Integer
    W = Width - 120
    With Outer
        .Width = W
        .Height = W
        .Top = (Height / 2) - (.Height / 2)
        .Left = (Width / 2) - (.Width / 2)
    End With
End Sub

Public Sub Rotate(B As Double)
    Dim X1 As Integer
    Dim X2 As Integer
    Dim Y1 As Integer
    Dim Y2 As Integer
    B = B
    
    
    X1 = 650
    Y1 = Outer.Top
    
    X2 = ((X1 - 650) * Cos(B)) - ((Y1 - 650) * Sin(B))
    Y2 = ((X1 - 650) * Sin(B)) + ((Y1 - 650) * Cos(B))
    
    Line1.X1 = X2 + 650
    Line1.Y1 = Y2 + 650
    
    Line2.X1 = -(X2 - 650)
    Line2.Y1 = -(Y2 - 650)
    
End Sub

Private Sub lblDeg_Change()
    With lblDeg
        .Top = (Height / 2) - (.Height / 2)
        .Left = ((Width / 2) - (.Width / 2)) + 15
    End With
End Sub

Private Sub UserControl_Resize()
    MoveBorder
    CreateCenter
    MoldOuter
    If Width <> 1300 Then Width = 1300
    If Height <> 1300 Then Height = 1300
    SetLines
    Radian = 360 / (Pi * 2)
    lblDeg = "0" & Chr(186)
End Sub


Sub SetLines()
    Line2.X2 = 650
    Line2.Y2 = 650
    Line1.Y2 = 650
    Line1.X2 = 650
    If ANGLE = 0 Then
        Line1.X1 = 650
        Line1.Y1 = Outer.Top
        Line2.X1 = 650
        Line2.Y1 = (Outer.Top + Outer.Height) - 30
    End If
    lblDeg = ANGLE
End Sub


Public Property Get Value() As Double
    Value = Value1
End Property
Public Property Let Value(newValue As Double)
    lblDeg = newValue & Chr(186)
    PropertyChanged "Value"
    Rotate (newValue / Radian)
End Property


Public Property Get TextColor() As OLE_COLOR
    TextColor = lblDeg.ForeColor
End Property

Public Property Let TextColor(newTextColor As OLE_COLOR)
    lblDeg.ForeColor = newTextColor
    PropertyChanged "TextColor"
End Property



Public Property Get IndicatorColor() As OLE_COLOR
    IndicatorColor = Line1.BorderColor
End Property

Public Property Let IndicatorColor(newIndicatorColor As OLE_COLOR)
    Line1.BorderColor = newIndicatorColor
    PropertyChanged "IndicatorColor"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lblDeg = PropBag.ReadProperty("Value", 0) & Chr(186)
    lblDeg.ForeColor = PropBag.ReadProperty("textColor", vbBlack)
    Line1.BorderColor = PropBag.ReadProperty("IndicatorColor", vbRed)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Value", Val(lblDeg), 0)
    Call PropBag.WriteProperty("TextColor", lblDeg.ForeColor, vbWhite)
    Call PropBag.WriteProperty("IndicatorColor", Line1.BorderColor, vbRed)
End Sub
