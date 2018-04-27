VERSION 5.00
Begin VB.UserControl PNACheckBox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
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
   MousePointer    =   99  'Custom
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "PNACheckBox.ctx":0000
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "PNACheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type IPPoint
   x As Long
   y As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, mRect As Rect) As Long
Private Declare Function GetCursorPos Lib "user32" (mPoint As IPPoint) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetRect Lib "user32" (mRect As Rect, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExW" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As Rect, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As Rect) As Long
'--
Private mText As String
Private mCheck As Boolean
Private mValue As Integer
Private mMoveDown As Boolean
Private mMove As Boolean
Private mFocus As Boolean
'--
Private Function checkMove() As Boolean
   Dim mRect As Rect, mPoint As IPPoint
   GetWindowRect hwnd, mRect
   mRect.Right = mRect.Left + 20
   GetCursorPos mPoint
   If mPoint.x >= mRect.Left And mPoint.x <= mRect.Right And mPoint.y >= mRect.Top And mPoint.y <= mRect.Bottom Then
      checkMove = True
      Exit Function
   End If
   checkMove = False
End Function
'--
Public Property Let Enabled(ByVal value As Boolean)
   UserControl.Enabled = value
   If UserControl.Enabled Then
      If mCheck Then
         mValue = 105
      Else
         mValue = 101
      End If
   Else
      If mCheck Then
         mValue = 108
      Else
         mValue = 104
      End If
   End If
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
Public Property Let Check(ByVal value As Boolean)
   mCheck = value
   If UserControl.Enabled Then
      If mCheck Then
         mValue = 105
      Else
         mValue = 101
      End If
   Else
      If mCheck Then
         mValue = 108
      Else
         mValue = 104
      End If
   End If
   UserControl_Paint
   PropertyChanged "Check"
End Property
Public Property Get Check() As Boolean
   Check = mCheck
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
   mText = "": mCheck = False
   mValue = 101: mMoveDown = False: mMove = False
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
   UserControl.PaintPicture LoadResPicture(mValue, 0), 3, (ScaleHeight - 13) / 2
   SetRect mRect, 20, 3, ScaleWidth, ScaleHeight - 1
   If UserControl.Enabled Then
      SetTextColor hdc, UserControl.ForeColor
      DrawTextEx hdc, StrConv(mText, vbUnicode), Len(mText), mRect, &H0, ByVal 0&
   Else
      SetTextColor hdc, RGB(150, 150, 150)
      DrawTextEx hdc, StrConv(mText, vbUnicode), Len(mText), mRect, &H0, ByVal 0&
   End If
   mRect.Left = 18: mRect.Top = 1
   mRect.Right = 22 + UserControl.TextWidth(mText)
   If mFocus Then DrawFocusRect hdc, mRect
End Sub
Private Sub UserControl_GotFocus()
   mFocus = True
   UserControl_Paint
End Sub
Private Sub UserControl_LostFocus()
   mMove = False
   mMoveDown = False
   mFocus = False
   If mCheck Then
      mValue = 105
   Else
      mValue = 101
   End If
   UserControl_Paint
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mMoveDown = False Then
      mMoveDown = True
   End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mMove = False Then
      mMove = True
      Timer.Enabled = True
   End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If checkMove And mMoveDown Then
      If mCheck = False Then
         mValue = 105
         mCheck = True
      Else
         mValue = 101
         mCheck = False
      End If
      mMoveDown = False
   End If
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 32
         If mCheck Then
            mValue = 107
         Else
            mValue = 103
         End If
         UserControl_Paint
      Case 39, 40
         SendKeys "+{Tab}"
      Case 37, 38
         SendKeys "+{Tab}"
   End Select
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If mCheck = False Then
         mValue = 105
         mCheck = True
      Else
         mValue = 101
         mCheck = False
      End If
      UserControl_Paint
   End If
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
   Height = (UserControl.TextHeight("A") + 6) * Screen.TwipsPerPixelY
   Width = (25 + UserControl.TextWidth(mText)) * Screen.TwipsPerPixelX
   UserControl_Paint
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      UserControl.Enabled = .ReadProperty("Enabled", True)
      mText = .ReadProperty("Text", "")
      mCheck = .ReadProperty("Check", False)
      Set Font = .ReadProperty("Font", Font)
      UserControl.ForeColor = .ReadProperty("ForeColor", UserControl.ForeColor)
      UserControl.BackColor = .ReadProperty("BackColor", UserControl.BackColor)
      UserControl_Paint
   End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enabled", UserControl.Enabled
      .WriteProperty "Check", mCheck
      .WriteProperty "Text", mText
      .WriteProperty "Font", Font
      .WriteProperty "ForeColor", UserControl.ForeColor
      .WriteProperty "BackColor", UserControl.BackColor
      UserControl_Paint
   End With
End Sub
'--
Private Sub Timer_Timer()
   If checkMove Then
      If mMoveDown Then
         If mCheck Then
            mValue = 107
         Else
            mValue = 103
         End If
         Set MouseIcon = LoadResPicture(102, 1)
      Else
         If mCheck Then
            mValue = 106
         Else
            mValue = 102
         End If
         Set MouseIcon = LoadResPicture(101, 1)
      End If
      mMove = True
   Else
      mMove = False
      mMoveDown = False
      Timer.Enabled = False
      If UserControl.Enabled Then
         If mCheck Then
            mValue = 105
         Else
            mValue = 101
         End If
      Else
         If mCheck Then
            mValue = 108
         Else
            mValue = 104
         End If
      End If
      Set MouseIcon = Nothing
   End If
   UserControl_Paint
End Sub
