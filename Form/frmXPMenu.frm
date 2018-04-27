VERSION 5.00
Begin VB.Form frmXPMenu 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   414
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   447
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMenuBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   3165
      Left            =   1560
      ScaleHeight     =   211
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      Top             =   780
      Visible         =   0   'False
      Width           =   2340
      Begin VB.PictureBox picIcon 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0080C0FF&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   720
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   690
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picPopup 
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FF00&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   1
         Top             =   1770
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Timer tmrHover 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   870
         Top             =   930
      End
   End
   Begin VB.Timer tmrActive 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   360
      Top             =   3690
   End
End
Attribute VB_Name = "frmXPMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenuClass As clsXPMenu
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public upY As Single
Private Sub Form_Click()
    Dim selectedItem As Long
    selectedItem = XPMenuClass.GetHilightedItem(upY)
    If XPMenuClass.IsTextItem(CInt(selectedItem)) Then
        XPMenuClass.KillAllMenus
        HandleClick XPMenuClass.GetMenuName(), CInt(selectedItem), XPMenuClass.GetItemText(CInt(selectedItem))
    End If
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim getHilight As Long
    getHilight = XPMenuClass.GetHilightedItem(y)
    If getHilight = XPMenuClass.GetHilightNum Then Exit Sub
    XPMenuClass.setHilightedItem CInt(getHilight)
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    upY = y
End Sub
Private Sub tmrActive_Timer()
    Dim frm As Form
    For Each frm In Forms
        If frm.Tag = "XPMenu" And GetActiveWindow() = frm.hwnd Then Exit Sub
    Next frm
    XPMenuClass.KillPopupMenus
    XPMenuClass.UnloadMenu
End Sub
Private Sub tmrHover_Timer()
    Dim pt As POINTAPI
    GetCursorPos pt
    Dim hw As Long
    hw = WindowFromPoint(pt.x, pt.y)
    If hw <> Me.hwnd Then
        If XPMenuClass.PopupShown() = False Then
            XPMenuClass.setHilightedItem -1
        End If
    End If
End Sub


