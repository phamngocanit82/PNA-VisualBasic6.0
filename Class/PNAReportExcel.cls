VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PNAReportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--
Dim mXMLDoc As MSXML2.DOMDocument
Dim mElement As MSXML2.IXMLDOMElement
Dim mElement2 As MSXML2.IXMLDOMElement
Dim mCollection As Collection

Dim mExcel As Excel.Application
Dim mWorkSheet As Excel.Worksheet
Dim mRow As Long
Dim mCol As Long
'--
Private Sub Class_Initialize()
   mRow = 0
   mCol = 65
   If mExcel Is Nothing Then
      Set mXMLDoc = New MSXML2.DOMDocument
      Set mCollection = New Collection
      Set mExcel = GetObject("", "Excel.Application")
      mExcel.DisplayAlerts = False
   End If
End Sub
'--
Public Function setRow(ByVal NewmRow As Long) As Long
   mRow = NewmRow
   setRow = mRow
End Function
Public Function getRow() As Long
   getRow = mRow
End Function
Public Function increaseRow() As Long
   mRow = mRow + 1
   increaseRow = mRow
End Function
'--
Public Function setCol(ByVal NewmCol As Long) As Long
   mCol = NewmCol
   setCol = mCol
End Function
Public Function getCol() As Long
   getCol = mCol
End Function
Public Function increaseCol() As Long
   mCol = mCol + 1
   increaseCol = mCol
End Function
'--
Public Property Let setCell(ByVal NewColumn As Long, ByVal NewRow As Long, ByVal NewValue As Variant)
   mWorkSheet.Cells(NewRow, NewColumn) = NewValue
End Property
Public Property Let setRange(ByVal sRange As String, ByVal NewValue As Variant)
   mWorkSheet.Range(sRange) = NewValue
End Property
Public Function setMerge(ByVal sRange As String)
   mWorkSheet.Range(sRange).Merge
End Function
Public Property Let setHorizontalAlignment(ByVal sRange As String, ByVal NewHorizontalAlignment As Long)
   mWorkSheet.Range(sRange).HorizontalAlignment = NewHorizontalAlignment
End Property
Public Property Let setVerticalAlignment(ByVal sRange As String, ByVal NewVerticalAlignment As Long)
   mWorkSheet.Range(sRange).VerticalAlignment = NewVerticalAlignment
End Property
Public Property Let setFontBold(ByVal sRange As String, ByVal NewBold As Boolean)
   mWorkSheet.Range(sRange).Font.Bold = NewBold
End Property
Public Property Let setFontSize(ByVal sRange As String, ByVal NewSize As Integer)
   mWorkSheet.Range(sRange).Font.Name = "Arial"
   mWorkSheet.Range(sRange).Font.Size = NewSize
End Property
Public Property Let setRowHeight(ByVal sRange As String, ByVal NewHeight As Integer)
   mWorkSheet.Range(sRange).RowHeight = NewHeight
End Property
Public Property Let setColumnWidth(ByVal sRange As String, ByVal NewWidth As Long)
   mWorkSheet.Columns(sRange).ColumnWidth = NewWidth
End Property
Public Property Let setPatterns(ByVal sRange As String, ByVal NewColor As ColorConstants)
   mWorkSheet.Range(sRange).Interior.Color = NewColor
End Property
'--
Public Function setBorderLeft(ByVal sRange As String, ByVal NewXlLineStyle As String, ByVal NewXlBorderWeight As String)
   mWorkSheet.Range(sRange).Borders(xlEdgeLeft).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlEdgeLeft).Weight = getExcelFormat(NewXlBorderWeight)
End Function
Public Function setBorderRight(ByVal sRange As String, ByVal NewXlLineStyle As String, ByVal NewXlBorderWeight As String)
   mWorkSheet.Range(sRange).Borders(xlEdgeRight).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlEdgeRight).Weight = getExcelFormat(NewXlBorderWeight)
End Function
Public Function setBorderTop(ByVal sRange As String, ByVal NewXlLineStyle As String, ByVal NewXlBorderWeight As String)
   mWorkSheet.Range(sRange).Borders(xlEdgeTop).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlEdgeTop).Weight = getExcelFormat(NewXlBorderWeight)
End Function
Public Function setBorderBottom(ByVal sRange As String, ByVal NewXlLineStyle As String, ByVal NewXlBorderWeight As String)
   mWorkSheet.Range(sRange).Borders(xlEdgeBottom).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlEdgeBottom).Weight = getExcelFormat(NewXlBorderWeight)
End Function
Public Function setBorderInside(ByVal sRange As String, ByVal NewXlLineStyle As String, ByVal XlBorderWeight As String)
   mWorkSheet.Range(sRange).Borders(xlInsideHorizontal).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlInsideHorizontal).Weight = getExcelFormat(XlBorderWeight)
   mWorkSheet.Range(sRange).Borders(xlInsideVertical).LineStyle = getExcelFormat(NewXlLineStyle)
   mWorkSheet.Range(sRange).Borders(xlInsideVertical).Weight = getExcelFormat(XlBorderWeight)
End Function
'--
Public Function setNumberFormat(ByVal sColumns As String)
   mWorkSheet.Columns(sColumns).NumberFormat = "#,##0.00"
End Function
'--
Public Function HBreakPage(ByVal sRange As String)
   mWorkSheet.HPageBreaks.Add mExcel.Application.Range(sRange)
End Function
'--
Public Function SetAutoFit(ByVal sColumns As String)
  mWorkSheet.Columns(sColumns).AutoFit
End Function
'--
Public Function setPicture()
   mWorkSheet.Shapes.AddPicture App.Path & "\SMOKES.BMP", 1, 1, 120, 10, 20, 20
End Function
'--
Public Function setFreeze(ByVal sRange As String)
   mWorkSheet.Range(sRange).Select
   mExcel.Application.ActiveWindow.FreezePanes = True
End Function
'--
Public Function getExcelFormat(ByVal s As String) As Long
   getExcelFormat = Val(mCollection(s))
End Function
'--
Private Sub PageSetupFormat()
On Error Resume Next
   mWorkSheet.Range("A1:" & Chr(getCol) & getRow).Font.Name = "Arial"
   CallByName mWorkSheet.PageSetup, "LeftMargin", VbLet, mExcel.InchesToPoints(0.05)
   CallByName mWorkSheet.PageSetup, "RightMargin", VbLet, mExcel.InchesToPoints(0.05)
   CallByName mWorkSheet.PageSetup, "TopMargin", VbLet, mExcel.InchesToPoints(0.25)
   CallByName mWorkSheet.PageSetup, "BottomMargin", VbLet, mExcel.InchesToPoints(0.5)
   CallByName mWorkSheet.PageSetup, "HeaderMargin", VbLet, mExcel.InchesToPoints(0.25)
   CallByName mWorkSheet.PageSetup, "FooterMargin", VbLet, mExcel.InchesToPoints(0.25)
   mWorkSheet.PageSetup.PrintGridlines = False
   mWorkSheet.PageSetup.PrintHeadings = False
   mWorkSheet.PageSetup.CenterHorizontally = True
   mWorkSheet.PageSetup.CenterVertically = False
   mWorkSheet.PageSetup.PaperSize = xlPaperA4
   mWorkSheet.PageSetup.Orientation = xlLandscape
   mWorkSheet.PageSetup.CenterFooter = "&P of &N"
   mWorkSheet.PageSetup.Zoom = False
   mWorkSheet.PageSetup.FitToPagesWide = 1
   mWorkSheet.PageSetup.FitToPagesTall = 32767
End Sub
'--
Public Function setExcelReport() As Boolean
On Error GoTo Set_ExcelReport
   Dim mWorkBook As Excel.Workbook
   Dim i As Integer
   mXMLDoc.Load App.Path & "\excelFormat.xml"
   Set mElement = mXMLDoc.getElementsByTagName("definition")(0)
   For i = 0 To mElement.childNodes.Length - 1
      Set mElement2 = mElement.childNodes(i)
      mCollection.Add mElement2.getAttribute("value"), mElement2.getAttribute("name")
   Next
   Set mWorkBook = mExcel.Workbooks.Add
   While mWorkBook.Worksheets.Count > 1
       mWorkBook.ActiveSheet.Delete
   Wend
   Set mWorkSheet = mWorkBook.Worksheets(1)
   mExcel.Caption = "Pham Ngoc An"
   mWorkSheet.Name = "PhamNgocAn"
   setExcelReport = True
Exit Function
Set_ExcelReport:
   setExcelReport = False
End Function
'--
Public Function showExcelReport() As Boolean
On Error GoTo Show_ExcelReport
   PageSetupFormat
   mExcel.Visible = True
   Set mElement2 = Nothing
   Set mElement = Nothing
   Set mXMLDoc = Nothing
   Set mCollection = Nothing
   Set mWorkSheet = Nothing
   showExcelReport = True
Exit Function
Show_ExcelReport:
   mExcel.Quit
   Set mExcel = Nothing
   showExcelReport = False
End Function



