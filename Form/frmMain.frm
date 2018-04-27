VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "PhamNgocanProject"
   ClientHeight    =   8115
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   ScaleHeight     =   8115
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   3810
      TabIndex        =   10
      Top             =   5100
      Width           =   1245
   End
   Begin PhamNgocAnProject.PNATextBox PNATextBox2 
      Height          =   300
      Left            =   2130
      TabIndex        =   9
      Top             =   5190
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      Enable          =   -1  'True
      Text            =   "pham ngoc an"
      PasswordChar    =   ""
      Font            =   "frmMain.frx":0000
      ForeColor       =   -2147483640
   End
   Begin PhamNgocAnProject.PNATextBox PNATextBox1 
      Height          =   300
      Left            =   360
      TabIndex        =   8
      Top             =   5220
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      Enable          =   -1  'True
      Text            =   "pham ngoc an"
      PasswordChar    =   ""
      Font            =   "frmMain.frx":0030
      ForeColor       =   -2147483640
   End
   Begin PhamNgocAnProject.PNACheckBox PNACheckBox2 
      Height          =   315
      Left            =   3840
      TabIndex        =   7
      Top             =   4080
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Enabled         =   -1  'True
      Check           =   0   'False
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16761024
   End
   Begin PhamNgocAnProject.PNARadioButton PNARadioButton2 
      Height          =   315
      Left            =   2550
      TabIndex        =   6
      Top             =   4560
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Enabled         =   -1  'True
      Selected        =   0   'False
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16761024
   End
   Begin PhamNgocAnProject.PNARadioButton PNARadioButton1 
      Height          =   315
      Left            =   480
      TabIndex        =   5
      Top             =   3870
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Enabled         =   -1  'True
      Selected        =   0   'False
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16761024
   End
   Begin PhamNgocAnProject.PNALabel PNALabel1 
      Height          =   315
      Left            =   2130
      TabIndex        =   4
      Top             =   4080
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      Enabled         =   -1  'True
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BackColor       =   16761024
   End
   Begin PhamNgocAnProject.PNASuperGrid PNASuperGrid1 
      Height          =   3015
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5318
      Enabled         =   -1  'True
      Caption         =   ""
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      BackColor       =   12648384
   End
   Begin PhamNgocAnProject.PNACheckBox PNACheckBox1 
      Height          =   315
      Left            =   540
      TabIndex        =   2
      Top             =   4500
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      Enabled         =   -1  'True
      Check           =   0   'False
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      BackColor       =   16761024
   End
   Begin PhamNgocAnProject.PNAButton PNAButton2 
      Height          =   420
      Left            =   3510
      TabIndex        =   1
      Top             =   3420
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   741
      Enabled         =   -1  'True
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
   End
   Begin PhamNgocAnProject.PNAButton PNAButton1 
      Height          =   420
      Left            =   510
      TabIndex        =   0
      Top             =   3270
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   741
      Enabled         =   -1  'True
      Text            =   "Phan VAn Hoài"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   810
      Top             =   8130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0060
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0794
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public XPMenu As New clsXPMenu
Public XPMenu2 As New clsXPMenu
Public XPM_EFNet As New clsXPMenu
Public XPM_DALNet As New clsXPMenu

Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Sub Command1_Click()
    Dim mRs As Recordset
    Set mRs = New Recordset
    Set mRs = returnRecord
    mRs.MoveNext
    'Set XPM_DALNet = New clsXPMenu
    XPM_DALNet.Init "DALNet"
    XPM_DALNet.AddItem 0, "Server1 (blah blah blah)", False, False
    XPM_DALNet.AddItem 0, "Server2 (asdfasdf:5636)", False, False
    XPM_DALNet.AddItem 0, "Server3 (dalnet)", False, False
    XPM_DALNet.AddItem 0, "", False, True
    XPM_DALNet.AddItem 0, "Random DALNet Server", False, False
    
    'Set XPM_EFNet = New clsXPMenu
    XPM_EFNet.Init "EFNet"
    XPM_EFNet.AddItem 0, "Prison (irc.prison.net)", False, False
    XPM_EFNet.AddItem 0, "Lagged (irc.lagged.org)", False, False
    XPM_EFNet.AddItem 0, "Another one.... (unknown)", False, False
    XPM_EFNet.AddItem 0, "", False, True
    XPM_EFNet.AddItem 0, "Random EFNet Server", False, False
       
    'Set XPMenu2 = New clsXPMenu
    XPMenu2.Init "Servers"
    XPMenu2.AddItem 0, "DALNet", True, False, XPM_DALNet
    XPMenu2.AddItem 0, "EFNet", True, False, XPM_EFNet
    
    'Set XPMenu = New clsXPMenu
    XPMenu.Init "Connect", ImageList1
    XPMenu.AddItem 0, mRs.Fields(1).value, True, False, XPMenu2
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 1, "Connect", False, False
    XPMenu.AddItem 2, "Disconnect", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 3, "Change Profile", False, False
    XPMenu.AddItem 0, "", False, True
    XPMenu.AddItem 0, "Exit", False, False
    
    Dim pos As POINTAPI
    GetCursorPos pos
    XPMenu.ShowMenu pos.x, pos.y
End Sub
Private Sub Form_Load()
   Set PNASuperGrid1.DataSource = returnRecord
   PNASuperGrid1.CaptionFont = "Arial"
   PNASuperGrid1.Caption = returnRecord.Fields(1).value
   PNASuperGrid1.ColoumsCaption(1) = returnRecord.Fields(1).value
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next
    End
End Sub



