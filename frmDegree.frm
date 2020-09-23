VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Degrees"
   ClientHeight    =   1470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   2775
   StartUpPosition =   3  'Windows Default
   Begin Project1.UserControl1 UserControl11 
      Height          =   1305
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   2302
      TextColor       =   -2147483630
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "0"
      Top             =   600
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "º"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   540
      Width           =   75
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MsgBox UserControl11.Value
End Sub

Private Sub Text1_Change()
    On Error GoTo err
    If Val(Text1) > 359 Then Text1 = Val(Text1) - 360
    If Val(Text1) < 0 Then Text1 = 360 - (-(Val(Text1)))
    UserControl11.Value = Val(Text1)
    Exit Sub
err:
    MsgBox "Error"
    End
    Exit Sub
    
End Sub
