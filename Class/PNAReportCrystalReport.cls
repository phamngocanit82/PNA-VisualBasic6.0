VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PNAReportCrystalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As CRAXDRT.Report
Dim CRApp As CRAXDRT.Application
'--
Private Sub Class_Initialize()
   Set CRReport = New CRAXDRT.Report
   Set CRApp = New CRAXDRT.Application
End Sub
'--
Public Property Let setFormulaFields(ByVal fieldName As String, ByVal newText As String)
   CRReport.FormulaFields.GetItemByName(fieldName).Text = newText
End Property
'--
Public Property Let setParamater(ByVal index As Long, ByVal value As String)
   CRReport.ParameterFields(1).AddCurrentValue value
End Property
'--
Public Function reportFile(ByVal filePath As String, Optional mRs As ADODB.Recordset) As Boolean
On Error GoTo Report_File
   Set CRReport = CRApp.OpenReport(filePath, 1)
   If Not mRs Is Nothing Then
      Dim s() As String, i As Integer
      s = Split(ConnectString, ";")
      CRReport.DiscardSavedData
      CRReport.Database.Tables(1).ConnectionProperties.DeleteAll
      For i = LBound(s) To (UBound(s) - 1)
          CRReport.Database.Tables(1).ConnectionProperties.Add Left(s(i), InStr(s(i), "=") - 1), Right(s(i), Len(s(i)) - InStr(s(i), "="))
      Next i
      CRReport.Database.Tables(1).SetDataSource mRs
   End If
   Set mRs = Nothing
   reportFile = True
Exit Function
Report_File:
   reportFile = False
End Function
'--
Public Function showReport(RptViewer As CrystalActiveXReportViewer) As Boolean
On Error GoTo Show_Report
   With RptViewer
      .ReportSource = CRReport
      .ViewReport
      .Zoom (100)
   End With
   Set CRReport = Nothing
   Set CRApp = Nothing
   showReport = True
Exit Function
Show_Report:
   showReport = False
End Function
 





