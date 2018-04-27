Attribute VB_Name = "PNAVariable"
Option Explicit
Public Const ConnectString As String = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;PassWord=;Initial Catalog=studentManage;Data Source=(Local)"
Public Function returnRecord() As Recordset
   Dim mConn As ADODB.Connection
   Dim mRs As ADODB.Recordset
   Set mConn = New Connection
   Set mRs = New Recordset
   mConn.Open ConnectString, "sa", ""
   mRs.Open "SELECT * FROM Student ", mConn, adOpenStatic
   Set returnRecord = mRs
End Function
'Public Function testCrystal()
'   Dim mCrystal As PNAReportCrystalReport
'   mCrystal = New PNAReportCrystalReport
'   mCrystal.reportFile App.Path & "\Report.rpt", returnRecord
'   mCrystal.setFormulaFields("num") = "'" & "stt" & "'"
'   mCrystal.setFormulaFields("id") = "'" & "id" & "'"
'   mCrystal.setFormulaFields("name") = "'" & "name" & "'"
'   mCrystal.setFormulaFields("idClass") = "'" & "idClass" & "'"
'   mCrystal.setFormulaFields("number") = "CStr(RecordNumber,0)"
'   mCrystal.setParamater(1) = "003"
'   mCrystal.showReport frmMain.CrystalActiveXReportViewer1
'End Function
'Public Function testExcel()
'   Dim mExcel As PNAReportExcel
'   Set mExcel = New PNAReportExcel
'   mExcel.setExcelReport
'   Dim mRs As Recordset
'   Set mRs = New Recordset
'   Set mRs = returnRecord
'   Dim i As Integer
'   For i = 0 To mRs.RecordCount - 1
'      mExcel.setRange("A" & i + 2) = mRs.Fields("id").value
'      mExcel.setRange("B" & i + 2) = mRs.Fields("name").value
'      mExcel.setRange("C" & i + 2) = mRs.Fields("idClass").value
'   Next
'   mExcel.setNumberFormat ("C")
'   mExcel.HBreakPage "A2"
'   mExcel.setPicture
'   mExcel.setFreeze "A3"
'   mExcel.setPatterns("B4") = vbCyan
'   mExcel.SetAutoFit "A:C"
'   mExcel.showExcelReport
'End Function
