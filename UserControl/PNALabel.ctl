VERSION 5.00
Begin VB.UserControl PNALabel 
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "PNALabel.ctx":0000
End
Attribute VB_Name = "PNALabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--
Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function SetRect Lib "user32" (mRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As Rect, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
'--
Private mText As String
Private mColor As Long
'--
Public Property Let Enabled(ByVal value As Boolean)
   UserControl.Enabled = value
   UserControl_Paint
   PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property
Public Property Let Text(ByVal value As String)
   mText = value
   UserControl_Resize
   PropertyChanged "Text"
End Property
Public Property Get Text() As String
   Text = mText
End Property
Public Property Set Font(ByVal value As Font)
   Set UserControl.Font = value
   UserControl_Resize
   PropertyChanged "Font"
End Property
Public Property Get Font() As Font
   Set Font = UserControl.Font
End Property
Public Property Let ForeColor(ByVal value As OLE_COLOR)
   UserControl.ForeColor = value
   UserControl_Paint
   PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = UserControl.ForeColor
End Property
Public Property Let BackColor(ByVal value As OLE_COLOR)
   UserControl.BackColor = value
   UserControl_Paint
   PropertyChanged "BackColor"
End Property
Public Property Get BackColor() As OLE_COLOR
   BackColor = UserControl.BackColor
End Property
'--
Private Sub UserControl_Initialize()
   mText = ""
End Sub
Private Sub UserControl_InitProperties()
   BackColor = Parent.BackColor
End Sub
Private Sub UserControl_Paint()
   Dim mRect As Rect, mRs As Recordset
   Set mRs = New Recordset
   Set mRs = returnRecord
   mRs.MoveNext
   mText = mRs.Fields(1).value
   Cls
   UserControl.ScaleMode = 3
   SetRect mRect, 3, 3, ScaleWidth, ScaleHeight - 1
   DrawTextEx hdc, StrConv(mText, vbUnicode), Len(mText), mRect, &H0, ByVal 0&
End Sub
Private Sub UserControl_Resize()
   Height = (UserControl.TextHeight("A") + 6) * Screen.TwipsPerPixelY
   Width = (6 + UserControl.TextWidth(mText)) * Screen.TwipsPerPixelX
   UserControl_Paint
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      UserControl.Enabled = .ReadProperty("Enabled", True)
      mText = .ReadProperty("Text", "")
      Set Font = .ReadProperty("Font", Font)
      UserControl.ForeColor = .ReadProperty("ForeColor", UserControl.ForeColor)
      UserControl.BackColor = .ReadProperty("BackColor", UserControl.BackColor)
      UserControl_Paint
   End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enabled", UserControl.Enabled
      .WriteProperty "Text", mText
      .WriteProperty "Font", Font
      .WriteProperty "ForeColor", UserControl.ForeColor
      .WriteProperty "BackColor", UserControl.BackColor
      UserControl_Paint
   End With
End Sub
