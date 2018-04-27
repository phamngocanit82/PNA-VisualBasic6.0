VERSION 5.00
Begin VB.UserControl PNAButton 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   ClientHeight    =   3600
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "PNAButton.ctx":0000
   Begin VB.Timer Timer 
      Interval        =   100
      Left            =   30
      Top             =   0
   End
End
Attribute VB_Name = "PNAButton"
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
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As IPPoint) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As Rect, ByVal hBrush As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'--
Private mText As String
Private mFocus As Boolean
Private mMoveDown As Boolean
Private mMove As Boolean
'--
Private Sub setRgnCornor()
   Dim mRect As Rect
   Dim mRgn1 As Long, mRgn2 As Long, mRgn As Long
   GetWindowRect hwnd, mRect
   mRgn = CreateRectRgn(0, 0, mRect.Right, mRect.Bottom)
   mRgn2 = CreateRectRgn(0, 0, 0, 0)
   mRgn1 = CreateRectRgn(0, 0, 1, 1)
   CombineRgn mRgn2, mRgn, mRgn1, 4
   CombineRgn mRgn, mRgn2, mRgn1, 4
   
   mRgn2 = CreateRectRgn(0, ScaleHeight, 0, ScaleHeight)
   mRgn1 = CreateRectRgn(0, ScaleHeight, 1, ScaleHeight - 1)
   CombineRgn mRgn2, mRgn, mRgn1, 4
   CombineRgn mRgn, mRgn2, mRgn1, 4
   
   mRgn2 = CreateRectRgn(ScaleWidth, 0, ScaleWidth, 0)
   mRgn1 = CreateRectRgn(ScaleWidth, 0, ScaleWidth - 1, 1)
   CombineRgn mRgn2, mRgn, mRgn1, 4
   CombineRgn mRgn, mRgn2, mRgn1, 4
   
   mRgn2 = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth, ScaleHeight)
   mRgn1 = CreateRectRgn(ScaleWidth, ScaleHeight, ScaleWidth - 1, ScaleHeight - 1)
   CombineRgn mRgn2, mRgn, mRgn1, 4
   CombineRgn mRgn, mRgn2, mRgn1, 4
   SetWindowRgn hwnd, mRgn, True
End Sub
Private Function checkMove() As Boolean
   Dim mRect As Rect, mPoint As IPPoint
   GetWindowRect hwnd, mRect
   GetCursorPos mPoint
   If mPoint.x >= mRect.Left - 1 And mPoint.x <= mRect.Right - 1 And mPoint.y >= mRect.Top - 1 And mPoint.y <= mRect.Bottom - 1 Then
      checkMove = True
      Exit Function
   End If
   checkMove = False
End Function
Private Sub drawPattern()
   Dim mRect As Rect, mPoint As IPPoint
   Dim mBrush As Long, mPen As Long
   SetRect mRect, 0, 0, ScaleWidth, ScaleHeight
   If Enabled Then
      If mMoveDown Then
         'MoveDown
         mBrush = CreateSolidBrush(RGB(226, 225, 217))
         SelectObject hdc, mBrush
         FillRect hdc, mRect, mBrush
         mPen = CreatePen(0, 1, RGB(209, 204, 193))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 2: mPoint.y = 2
         MoveToEx hdc, 2, 2, mPoint
         LineTo hdc, ScaleWidth - 2, 2
         mPoint.x = 2: mPoint.y = ScaleHeight - 4
         MoveToEx hdc, 2, 2, mPoint
         LineTo hdc, 2, ScaleHeight - 4
         mPen = CreatePen(0, 1, RGB(220, 216, 207))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 2: mPoint.y = 3
         MoveToEx hdc, 2, 3, mPoint
         LineTo hdc, ScaleWidth - 2, 3
         mPoint.x = 3: mPoint.y = ScaleHeight - 4
         MoveToEx hdc, 3, 2, mPoint
         LineTo hdc, 3, ScaleHeight - 4
      
         mPen = CreatePen(0, 1, RGB(234, 233, 227))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 4
         MoveToEx hdc, 3, ScaleHeight - 4, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 4
         mPen = CreatePen(0, 1, RGB(242, 241, 238))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 3, ScaleHeight - 3, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
      Else
         If mMove Or mFocus Then
            'Focus Move
            mBrush = CreateSolidBrush(RGB(246, 246, 243))
            SelectObject hdc, mBrush
            FillRect hdc, mRect, mBrush
            mPen = CreatePen(0, 1, RGB(236, 235, 230))
            SelectObject hdc, mPen
            mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 5
            MoveToEx hdc, 4, ScaleHeight - 5, mPoint
            LineTo hdc, ScaleWidth - 4, ScaleHeight - 5
         Else
            mBrush = CreateSolidBrush(RGB(246, 246, 243))
            SelectObject hdc, mBrush
            FillRect hdc, mRect, mBrush
            mPen = CreatePen(0, 1, RGB(234, 231, 224))
            SelectObject hdc, mPen
            mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 4
            MoveToEx hdc, ScaleWidth - 3, 4, mPoint
            LineTo hdc, ScaleWidth - 3, ScaleHeight - 4
            mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 4
            MoveToEx hdc, 3, ScaleHeight - 4, mPoint
            LineTo hdc, ScaleWidth - 3, ScaleHeight - 4
            
            mPen = CreatePen(0, 1, RGB(236, 235, 230))
            SelectObject hdc, mPen
            mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 3
            MoveToEx hdc, ScaleWidth - 4, 3, mPoint
            LineTo hdc, ScaleWidth - 4, ScaleHeight - 3
            mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 5
            MoveToEx hdc, 2, ScaleHeight - 5, mPoint
            LineTo hdc, ScaleWidth - 3, ScaleHeight - 5
            
            mPen = CreatePen(0, 1, RGB(214, 208, 197))
            SelectObject hdc, mPen
            mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 3
            MoveToEx hdc, 3, ScaleHeight - 3, mPoint
            LineTo hdc, ScaleWidth - 4, ScaleHeight - 3
         End If
      End If
         
   Else
      'Disable
      mBrush = CreateSolidBrush(RGB(245, 244, 234))
      SelectObject hdc, mBrush
      FillRect hdc, mRect, mBrush
   End If
   DeleteObject mBrush
End Sub
Private Sub DrawText()
   Dim mRect As Rect
   SetRect mRect, 1, (ScaleHeight - TextHeight(mText)) / 2 + 1, ScaleWidth, ScaleHeight
   If UserControl.Enabled = False Then
      SetTextColor hdc, RGB(161, 161, 146)
   End If
   DrawTextEx hdc, StrConv(mText, vbUnicode), Len(mText), mRect, &H1, ByVal 0&
   If mFocus And mMoveDown = False Then
      SetRect mRect, 5, 5, ScaleWidth - 4, ScaleHeight - 4
      DrawFocusRect hdc, mRect
   End If
End Sub
Private Sub drawBorder()
   Dim mBrush As Long, mPen As Long, mPoint As IPPoint
   Dim mRect As Rect
   If UserControl.Enabled Then
      mPen = CreatePen(0, 1, RGB(224, 220, 201))
      SelectObject hdc, mPen
      mPoint.x = ScaleWidth - 1: mPoint.y = 1
      MoveToEx hdc, 1, 0, mPoint
      LineTo hdc, ScaleWidth - 1, 0
      
      mPen = CreatePen(0, 1, RGB(240, 238, 224))
      SelectObject hdc, mPen
      mPoint.x = ScaleWidth - 1: mPoint.y = ScaleHeight - 1
      MoveToEx hdc, ScaleWidth - 1, 1, mPoint
      LineTo hdc, ScaleWidth - 1, ScaleHeight - 1
      
      mPen = CreatePen(0, 1, RGB(248, 247, 240))
      SelectObject hdc, mPen
      mPoint.x = ScaleWidth - 1: mPoint.y = ScaleHeight - 1
      MoveToEx hdc, 1, ScaleHeight - 1, mPoint
      LineTo hdc, ScaleWidth - 1, ScaleHeight - 1
      
      mPen = CreatePen(0, 1, RGB(226, 222, 203))
      SelectObject hdc, mPen
      mPoint.x = 0: mPoint.y = ScaleHeight - 1
      MoveToEx hdc, 0, 1, mPoint
      LineTo hdc, 0, ScaleHeight - 1
      '4 cornor
      SetPixel hdc, 1, 1, RGB(230, 225, 208): SetPixel hdc, ScaleWidth - 2, 1, RGB(232, 228, 210)
      SetPixel hdc, ScaleWidth - 2, ScaleHeight - 2, RGB(244, 244, 234): SetPixel hdc, 1, ScaleHeight - 2, RGB(236, 233, 216)
      '8 cornor
      SetPixel hdc, 3, 1, RGB(37, 87, 131): SetPixel hdc, 1, 3, RGB(37, 87, 131)
      SetPixel hdc, ScaleWidth - 4, 1, RGB(37, 87, 131): SetPixel hdc, ScaleWidth - 2, 3, RGB(37, 87, 131)
      SetPixel hdc, ScaleWidth - 2, ScaleHeight - 4, RGB(37, 87, 131): SetPixel hdc, ScaleWidth - 4, ScaleHeight - 2, RGB(37, 87, 131)
      SetPixel hdc, 1, ScaleHeight - 4, RGB(37, 87, 131): SetPixel hdc, 3, ScaleHeight - 2, RGB(37, 87, 131)
      '4 Line out
      mPen = CreatePen(0, 1, RGB(0, 60, 116))
      SelectObject hdc, mPen
      mPoint.x = ScaleWidth - 4: mPoint.y = 1
      MoveToEx hdc, 4, 1, mPoint
      LineTo hdc, ScaleWidth - 4, 1
      mPoint.x = ScaleWidth - 2: mPoint.y = ScaleHeight - 4
      MoveToEx hdc, ScaleWidth - 2, 4, mPoint
      LineTo hdc, ScaleWidth - 2, ScaleHeight - 4
      mPoint.x = 3: mPoint.y = ScaleHeight - 2
      MoveToEx hdc, ScaleWidth - 5, ScaleHeight - 2, mPoint
      LineTo hdc, 3, ScaleHeight - 2
      mPoint.x = 1: mPoint.y = ScaleHeight - 4
      MoveToEx hdc, 1, 4, mPoint
      LineTo hdc, 1, ScaleHeight - 4
      '4 Line in
   If mFocus Then
         'Focus
         mPen = CreatePen(0, 1, RGB(206, 231, 255))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = 2
         MoveToEx hdc, 3, 2, mPoint
         LineTo hdc, ScaleWidth - 3, 2
         mPen = CreatePen(0, 1, RGB(188, 212, 246))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = 3
         MoveToEx hdc, 3, 3, mPoint
         LineTo hdc, ScaleWidth - 3, 3
         mPoint.x = 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 3, 3, mPoint
         LineTo hdc, 3, ScaleHeight - 3
         mPoint.x = 2: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 2, 3, mPoint
         LineTo hdc, 2, ScaleHeight - 3
      
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, ScaleWidth - 3, 3, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
         mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, ScaleWidth - 4, 3, mPoint
         LineTo hdc, ScaleWidth - 4, ScaleHeight - 3
      
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 4
         MoveToEx hdc, 3, ScaleHeight - 4, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 4
         mPen = CreatePen(0, 1, RGB(105, 130, 238))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 3, ScaleHeight - 3, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
      End If
      If mMove Then
         'Move
         mPen = CreatePen(0, 1, RGB(255, 240, 207))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = 2
         MoveToEx hdc, 3, 2, mPoint
         LineTo hdc, ScaleWidth - 3, 2
         mPen = CreatePen(0, 1, RGB(253, 216, 137))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = 3
         MoveToEx hdc, 3, 3, mPoint
         LineTo hdc, ScaleWidth - 3, 3
         mPoint.x = 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 3, 3, mPoint
         LineTo hdc, 3, ScaleHeight - 3
         mPoint.x = 2: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 2, 3, mPoint
         LineTo hdc, 2, ScaleHeight - 3
      
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, ScaleWidth - 3, 3, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
         mPoint.x = ScaleWidth - 4: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, ScaleWidth - 4, 3, mPoint
         LineTo hdc, ScaleWidth - 4, ScaleHeight - 3
      
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 4
         MoveToEx hdc, 3, ScaleHeight - 4, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 4
         mPen = CreatePen(0, 1, RGB(229, 151, 0))
         SelectObject hdc, mPen
         mPoint.x = ScaleWidth - 3: mPoint.y = ScaleHeight - 3
         MoveToEx hdc, 3, ScaleHeight - 3, mPoint
         LineTo hdc, ScaleWidth - 3, ScaleHeight - 3
      End If
      DeleteObject mPen
   Else
      'Disable
      SetRect mRect, 0, 0, ScaleWidth, ScaleHeight
      mBrush = CreateSolidBrush(RGB(201, 199, 186))
      FrameRect hdc, mRect, mBrush
      DeleteObject mBrush
   End If
End Sub
Private Sub drawInnerBorder()
   Dim mRgn As Long
   If UserControl.Enabled Then
      SetPixel hdc, 0, 0, RGB(255, 0, 0): SetPixel hdc, 1, 2, RGB(122, 149, 168)
      SetPixel hdc, 2, 1, RGB(122, 149, 168): SetPixel hdc, 2, 2, RGB(77, 117, 152)
      SetPixel hdc, 2, 3, RGB(157, 168, 172): SetPixel hdc, 3, 2, RGB(157, 168, 172)
   
      SetPixel hdc, 0, ScaleHeight - 1, RGB(255, 0, 0): SetPixel hdc, 1, ScaleHeight - 3, RGB(122, 149, 168)
      SetPixel hdc, 2, ScaleHeight - 2, RGB(122, 149, 168): SetPixel hdc, 2, ScaleHeight - 3, RGB(77, 117, 152)
      SetPixel hdc, 2, ScaleHeight - 4, RGB(157, 168, 172): SetPixel hdc, 3, ScaleHeight - 3, RGB(157, 168, 172)
   
      SetPixel hdc, ScaleWidth - 1, 0, RGB(255, 0, 0): SetPixel hdc, ScaleWidth - 2, 2, RGB(122, 149, 168)
      SetPixel hdc, ScaleWidth - 3, 1, RGB(122, 149, 168): SetPixel hdc, ScaleWidth - 3, 2, RGB(77, 117, 152)
      SetPixel hdc, ScaleWidth - 3, 3, RGB(157, 168, 172): SetPixel hdc, ScaleWidth - 4, 2, RGB(157, 168, 172)
   
      SetPixel hdc, ScaleWidth - 1, ScaleHeight - 1, RGB(255, 0, 0): SetPixel hdc, ScaleWidth - 2, ScaleHeight - 3, RGB(122, 149, 168)
      SetPixel hdc, ScaleWidth - 3, ScaleHeight - 2, RGB(122, 149, 168): SetPixel hdc, ScaleWidth - 3, ScaleHeight - 3, RGB(77, 117, 152)
      SetPixel hdc, ScaleWidth - 3, ScaleHeight - 4, RGB(157, 168, 172): SetPixel hdc, ScaleWidth - 4, ScaleHeight - 3, RGB(157, 168, 172)
   Else
      'Disable
      SetPixel hdc, 0, 0, RGB(255, 0, 0): SetPixel hdc, 0, 1, RGB(216, 214, 202)
      SetPixel hdc, 1, 0, RGB(216, 214, 202): SetPixel hdc, 1, 1, RGB(216, 214, 202)
      SetPixel hdc, 1, 2, RGB(234, 233, 222): SetPixel hdc, 2, 1, RGB(234, 233, 222)
   
      SetPixel hdc, 0, ScaleHeight - 1, RGB(255, 0, 0): SetPixel hdc, 0, ScaleHeight - 2, RGB(216, 214, 202)
      SetPixel hdc, 1, ScaleHeight - 1, RGB(216, 214, 202): SetPixel hdc, 1, ScaleHeight - 2, RGB(216, 214, 202)
      SetPixel hdc, 1, ScaleHeight - 3, RGB(234, 233, 222): SetPixel hdc, 2, ScaleHeight - 2, RGB(234, 233, 222)
   
      SetPixel hdc, ScaleWidth - 1, 0, RGB(255, 0, 0): SetPixel hdc, ScaleWidth - 1, 1, RGB(216, 214, 202)
      SetPixel hdc, ScaleWidth - 2, 0, RGB(216, 214, 202): SetPixel hdc, ScaleWidth - 2, 1, RGB(216, 214, 202)
      SetPixel hdc, ScaleWidth - 2, 2, RGB(234, 233, 222): SetPixel hdc, ScaleWidth - 3, 1, RGB(234, 233, 222)
   
      SetPixel hdc, ScaleWidth - 1, ScaleHeight - 1, RGB(255, 0, 0): SetPixel hdc, ScaleWidth - 1, ScaleHeight - 2, RGB(216, 214, 202)
      SetPixel hdc, ScaleWidth - 2, ScaleHeight - 1, RGB(216, 214, 202): SetPixel hdc, ScaleWidth - 2, ScaleHeight - 2, RGB(216, 214, 202)
      SetPixel hdc, ScaleWidth - 2, ScaleHeight - 3, RGB(234, 233, 222): SetPixel hdc, ScaleWidth - 3, ScaleHeight - 2, RGB(234, 233, 222)
   End If
End Sub
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
   UserControl_Paint
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
'--
Private Sub UserControl_Paint()
   Dim mRs As Recordset
   Set mRs = New Recordset
   Set mRs = returnRecord
   mRs.MoveNext
   mText = mRs.Fields(1).value
   Cls
   UserControl.ScaleMode = 3
   setRgnCornor
   drawPattern
   DrawText
   drawInnerBorder
   drawBorder
End Sub
Private Sub UserControl_LostFocus()
   mMove = False
   mMoveDown = False
   mFocus = False
   UserControl_Paint
End Sub
Private Sub UserControl_GotFocus()
   mFocus = True
   UserControl_Paint
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
   Height = (UserControl.TextHeight("A") + 13) * Screen.TwipsPerPixelY
   Width = (25 + UserControl.TextWidth(mText)) * Screen.TwipsPerPixelX
   UserControl_Paint
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mMoveDown = False Then
      If mMove Then mMove = False
      Set MouseIcon = LoadResPicture(102, 1)
      mMoveDown = True
      Timer.Enabled = True
      UserControl_Paint
   End If
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   If mMove = False And mMoveDown = False Then
      Set MouseIcon = LoadResPicture(101, 1)
      mMove = True
      Timer.Enabled = True
      UserControl_Paint
   End If
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   mMoveDown = False
   Timer.Enabled = False
   UserControl_Paint
End Sub
Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13
         If mMove Then mMove = False
         If Timer.Enabled Then Timer.Enabled = False
         mMoveDown = True
         UserControl_Paint
      Case 39, 40
         SendKeys "+{Tab}"
      Case 37, 38
         SendKeys "+{Tab}"
   End Select
End Sub
Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    mMoveDown = False
    UserControl_Paint
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      UserControl.Enabled = .ReadProperty("Enabled", True)
      mText = .ReadProperty("Text", "")
      Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
      UserControl.ForeColor = .ReadProperty("ForeColor", UserControl.ForeColor)
   End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enabled", UserControl.Enabled
      .WriteProperty "Text", mText
      .WriteProperty "Font", UserControl.Font
      .WriteProperty "ForeColor", UserControl.ForeColor
   End With
End Sub
Private Sub Timer_Timer()
   If checkMove = False Then
      mMove = False
      mMoveDown = False
      Timer.Enabled = False
      Set MouseIcon = Nothing
      UserControl_Paint
   End If
End Sub
