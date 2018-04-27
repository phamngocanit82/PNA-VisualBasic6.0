VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.UserControl PNATextBox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000080FF&
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1815
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   25
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   ToolboxBitmap   =   "PNATextBox.ctx":0000
   Begin VB.Shape shape 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1815
   End
   Begin MSForms.TextBox textBox 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      VariousPropertyBits=   750798875
      BackColor       =   16777215
      MousePointer    =   99
      Size            =   "2566;450"
      BorderColor     =   -2147483644
      SpecialEffect   =   0
      MouseIcon       =   "PNATextBox.ctx":0312
      FontName        =   "Arial"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "PNATextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--
Public Property Let Enable(ByVal value As Boolean)
   textBox.Enabled = value
   Enabled = value
   PropertyChanged "Enable"
End Property
Public Property Get Enable() As Boolean
   Enable = textBox.Enabled
End Property
Public Property Let Text(ByVal value As String)
   textBox.Text = value
   PropertyChanged "Text"
End Property
Public Property Get Text() As String
   Text = textBox.Text
End Property
Public Property Let PasswordChar(ByVal value As String)
   textBox.PasswordChar = value
   PropertyChanged "PasswordChar"
End Property
Public Property Get PasswordChar() As String
   PasswordChar = textBox.PasswordChar
End Property
Public Property Set Font(ByVal value As Font)
   Set textBox.Font = value
   PropertyChanged "Font"
End Property
Public Property Get Font() As Font
   Set Font = textBox.Font
End Property
Public Property Let ForeColor(ByVal value As OLE_COLOR)
   textBox.ForeColor = value
   PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = textBox.ForeColor
End Property
'--
Private Sub textBox_GotFocus()
   textBox.SelStart = 0
   textBox.SelLength = Len(textBox.Text)
   shape.BorderColor = &H80FF&
   textBox.BackColor = &HFFFFFF
End Sub
Private Sub textBox_LostFocus()
   shape.BorderColor = &HC000&
   textBox.BackColor = &HE0E0E0
End Sub
'--
Private Sub UserControl_Initialize()
   textBox.Text = "pham ngoc an"
   shape.BorderColor = &HC000&
   textBox.Top = 0: textBox.Left = -6
   textBox.Width = ScaleWidth + 6: textBox.Height = ScaleHeight
   shape.Top = 0: shape.Left = 0
   shape.Width = ScaleWidth: shape.Height = ScaleHeight
End Sub
Private Sub UserControl_Resize()
   Height = 300
   textBox.Width = ScaleWidth + 6: textBox.Height = ScaleHeight
   shape.Width = ScaleWidth: shape.Height = ScaleHeight
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      textBox.Enabled = .ReadProperty("Enable", True)
      textBox.Text = .ReadProperty("Text", "")
      textBox.PasswordChar = .ReadProperty("PasswordChar", "")
      Set textBox.Font = .ReadProperty("Font", textBox.Font)
      textBox.ForeColor = .ReadProperty("ForeColor", textBox.ForeColor)
      Enable = textBox.Enabled
   End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enable", textBox.Enabled
      .WriteProperty "Text", textBox.Text
      .WriteProperty "PasswordChar", textBox.PasswordChar
      .WriteProperty "Font", textBox.Font
      .WriteProperty "ForeColor", textBox.ForeColor
   End With
End Sub
