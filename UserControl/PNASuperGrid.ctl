VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.UserControl PNASuperGrid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4245
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   1890
   ScaleWidth      =   4245
   ToolboxBitmap   =   "PNASuperGrid.ctx":0000
   Begin MSDataGridLib.DataGrid dataGrid 
      Height          =   1845
      Left            =   20
      TabIndex        =   0
      Top             =   20
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3254
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648384
      ForeColor       =   192
      HeadLines       =   1
      RowHeight       =   18
      RowDividerStyle =   3
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PNASuperGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--
Public Event RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'--
Private mRow As Long
'--
Public Property Let Enabled(ByVal value As Boolean)
   dataGrid.Enabled = value
   UserControl.Enabled = value
   PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
   Enabled = UserControl.Enabled
End Property
Public Property Let Caption(ByVal value As String)
   dataGrid.Caption = value
   PropertyChanged "Caption"
End Property
Public Property Get Caption() As String
   Caption = dataGrid.Caption
End Property
Public Property Set CaptionFont(ByVal value As Font)
   Set dataGrid.HeadFont = value
   PropertyChanged "CaptionFont"
End Property
Public Property Get CaptionFont() As Font
   Set CaptionFont = dataGrid.HeadFont
End Property
Public Property Set Font(ByVal value As Font)
   Set dataGrid.Font = value
   PropertyChanged "Font"
End Property
Public Property Get Font() As Font
   Set Font = dataGrid.Font
End Property
Public Property Let ForeColor(ByVal value As OLE_COLOR)
   dataGrid.ForeColor = value
   PropertyChanged "ForeColor"
End Property
Public Property Get ForeColor() As OLE_COLOR
   ForeColor = dataGrid.ForeColor
End Property
Public Property Let BackColor(ByVal value As OLE_COLOR)
   dataGrid.BackColor = value
End Property
Public Property Get BackColor() As OLE_COLOR
   BackColor = dataGrid.BackColor
End Property
Public Property Set DataSource(ByVal mRs As Recordset)
   Set dataGrid.DataSource = mRs
End Property
Public Property Let ColoumsCaption(ByVal index As Long, value As String)
   dataGrid.Columns(index).Caption = value
   dataGrid.Columns(index).Width = (Len(value) * dataGrid.Font.Size) * Screen.TwipsPerPixelX
End Property
Public Property Get Rows() As Long
   Rows = mRow
End Property
Public Property Get ColoumnsValue(ByVal value As Long) As Variant
   ColoumnsValue = dataGrid.Columns(value).value
End Property
'--
Private Sub dataGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   mRow = dataGrid.Row
   RaiseEvent RowColChange(LastRow, LastCol)
End Sub
'--
Private Sub UserControl_Resize()
   dataGrid.Width = Width - 30
   dataGrid.Height = Height - 30
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   With PropBag
      dataGrid.Enabled = .ReadProperty("Enabled", True)
      dataGrid.Caption = .ReadProperty("Caption", "")
      Set dataGrid.HeadFont = .ReadProperty("CaptionFont", Font)
      Set dataGrid.Font = .ReadProperty("Font", Font)
      dataGrid.ForeColor = .ReadProperty("ForeColor", ForeColor)
      dataGrid.BackColor = .ReadProperty("BackColor", BackColor)
      UserControl.Enabled = dataGrid.Enabled
   End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   With PropBag
      .WriteProperty "Enabled", dataGrid.Enabled
      .WriteProperty "Caption", dataGrid.Caption
      .WriteProperty "CaptionFont", dataGrid.HeadFont
      .WriteProperty "Font", dataGrid.Font
      .WriteProperty "ForeColor", dataGrid.ForeColor
      .WriteProperty "BackColor", dataGrid.BackColor
   End With
End Sub
