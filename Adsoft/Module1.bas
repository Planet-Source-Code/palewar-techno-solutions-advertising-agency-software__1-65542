Attribute VB_Name = "Module1"
Global cn As New ADODB.Connection
Global datalocation As String

Global ServiceTax As Single     'containing service tax
Global Commission As Single     'containing commission

Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Global Const conSwNormal = 1

Sub main()
Dim cnstr As String 'Connection String
On Error GoTo errorHandler
ServiceTax = 0.1224   'service tax is 12.24%, so we r converting it into (.1224).
Commission = 0.15   'commission is 15%, so we r converting it into (.15).
frmmdi.Show
frmSplash.Show
'extract Data Location from Registry
datalocation = GetSetting("AdSoft", "startup", "datalocation")

If datalocation = "" Or Dir(datalocation + "mazda.mdb") = "" Then
                 
    MsgBox "Database File Not Found in Data Location.", vbCritical
    AskDataLocation 'user defined procedure
    Exit Sub
End If

'make database connection
cnstr = "Driver=Microsoft Access Driver (*.mdb);DBQ=" & datalocation & "mazda.mdb;password=26378" 'Opens the conection, sets the ConnectionString to a DSN-Less connection. "Specifies the drives used;specifies the database;specifiec the password to use"
'this setting gives error while adding new record in notesmaster table
'cnstr = "provider=Microsoft.jet.oledb.4.0;Jet OLEDB:Database Password=26378;data source=" & datalocation & "mazda.mdb" 'Opens the conection, sets the ConnectionString to a DSN-Less connection. "Specifies the drives used;specifies the database;specifiec the password to use"
cn.Open cnstr

Exit Sub
errorHandler:
MsgBox "Error No. " & Err.Number & vbCrLf & Err.Description & "." & vbCrLf & "Please select a valid Data Location.", vbCritical, "Database Connection Failed!!"
'Unload frmmdi
'DeleteSetting "AdSoft", "startup", "datalocation"

End Sub
Function finyear(cur_date As Date) As Integer 'returns financial year
If Month(cur_date) > 3 Then
    finyear = Year(cur_date)
Else
    finyear = Year(cur_date) - 1
End If
End Function
Sub restore()
Dim WithFiles As Long
On Error GoTo errorHandler
src = BrowseForFolder(frmmdi.hwnd, "Restore From Backup Folder :", WithFiles, RecycleBin)

If src = "" Then
MsgBox "Please select a valid Backup Folder.", vbCritical, "Restore Failed!!"
Exit Sub
End If


If Right(src, 1) = "\" Then
    src = src + "mazda.mdb"
Else
    src = src + "\mazda.mdb"
End If


FileCopy src, datalocation + "\mazda.mdb"
MsgBox "Database restored successfully", vbInformation, "Restore Complete!!"
Unload frmmdi

Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical, "Error No. " & Err.Number

End Sub

Sub backuP()
Dim WithFiles As Long
On Error GoTo errorHandler

cn.Close
Set cn = Nothing

dest = BrowseForFolder(frmmdi.hwnd, "Select Backup Folder :", WithFiles, RecycleBin)

If dest = "" Then
MsgBox "Please select a valid Backup Folder.", vbCritical, "Backup Failed!!"
Exit Sub
End If


If Right(dest, 1) = "\" Then
    dest = dest + "mazda.mdb"
Else
    dest = dest + "\mazda.mdb"
End If

FileCopy datalocation + "\mazda.mdb", dest
MsgBox "Database copied in Backup Folder", vbInformation, "Backup Complete!!"
Unload frmmdi

Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical, "Error No. " & Err.Number

End Sub

Sub AskDataLocation() 'show browse for folder box for data location
Dim WithFiles As Long
'show browse for folder dialog
returnvaluE = BrowseForFolder(frmmdi.hwnd, "Specify Folder containing mazda.mdb:", WithFiles, RecycleBin)
'show dialog again if user selects nothing
If returnvaluE = "" Then AskDataLocation
'append \ in the end of path
If Right(Trim(returnvaluE), 1) <> "/" Or Right(Trim(returnvaluE), 1) <> "\" Then returnvaluE = Trim(returnvaluE) + "\"
'save data location in registry
SaveSetting "AdSoft", "startup", "datalocation", returnvaluE
'quit the application so that application uses data location when restarted
MsgBox "Data Location Saved." & vbCrLf & "Click OK to quit the Software.", vbInformation, "Restart Software!!"
Unload frmmdi
End Sub

Sub ShowReport(repname As String, frmDate As Date, toDate As Date)
On Error GoTo eror
Dim rec As New ADODB.Recordset 'recordset to act as report source
Dim recBill As New ADODB.Recordset
Dim recPub As New ADODB.Recordset
Dim Rep As New CRAXDRT.Report   'creating instance of crystal report dsr
Dim frmrep As New frmreports ' new instance of report form
Dim repApplication As New CRAXDRT.Application
Dim mystr As String
Select Case repname
    Case Is = "sales"
            Set Rep = repsales
            'open recordset using entered date
'***********************************************Anish***************************************
            If rec.State = 1 Then rec.Close
            rec.Open "SELECT billmaster.finyear,billmaster.billno,First(billmaster.ref) AS FirstOfref, " _
            & "First(billmaster.billdate) AS FirstOfbilldate, " _
            & "First(billmaster.client) AS FirstOfclient, Sum(billdetail.amount) AS SumOfamount " _
            & "FROM billdetail INNER JOIN billmaster ON " _
            & "(billdetail.billmaster_id = billmaster.billmaster_id) " _
            & "GROUP BY billmaster.finyear, billmaster.billno " _
            & "HAVING First(billmaster.cancelled)=No and " _
            & "First(billmaster.billdate) >=#" & CStr(frmDate) & "# " _
            & "and First(billmaster.billdate)<=#" & CStr(toDate) & "#", cn, adOpenStatic, adLockReadOnly

'*******************************************************************************************
'            rec.Open "SELECT billmaster.finyear, billmaster.billno, First(billmaster.ref) " _
'                    & "AS FirstOfref, First(billmaster.billdate) AS FirstOfbilldate, " _
'                    & "First(billmaster.client) AS FirstOfclient, Sum(billdetail.amount) " _
'                    & "AS SumOfamount FROM billdetail INNER JOIN billmaster ON " _
'                    & "(billdetail.billno = billmaster.billno) AND (billdetail.finyear " _
'                    & "= billmaster.finyear) GROUP BY billmaster.finyear, " _
'                    & "billmaster.billno HAVING First(billmaster.cancelled)=" _
'                    & "No and First(billmaster.billdate) >=#" & CStr(frmDate) & "# and First" _
'                    & "(billmaster.billdate)<=#" & CStr(toDate) & "#", cn, adOpenStatic, adLockReadOnly

 
            
            'displaying report
            Rep.DiscardSavedData        'it helps if the report load more than 1 times. it display always new records according to selected date.
            For i = 1 To Rep.FormulaFields.Count
                If Rep.FormulaFields(i).Name = "{@ServiceTaxValue}" Then Rep.FormulaFields(i).Text = ServiceTax
            Next

            For i = 1 To Rep.FormulaFields.Count
                 If Rep.FormulaFields(i).Name = "{@Commission}" Then Rep.FormulaFields(i).Text = Commission
            Next
            
            Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr
            Unload frmrep
            frmrep.CRViewer1.ReportSource = Rep
            frmrep.CRViewer1.PrintReport
            frmrep.Caption = "Sales Report"
            Set rec = Nothing
                        
            
    Case Is = "roregister"
            Dim recRoMaster As New ADODB.Recordset      'getting records from Romaster
            Dim recBillRecord As New ADODB.Recordset    'getting records from billmaster and billdetail
            Dim recPubRecord As New ADODB.Recordset     'getting records from publicationbill
            Dim recRoDetail As New ADODB.Recordset      'getting records from Rodetail
            
            Dim RepRelOrder As New repRelOrd            'instance of Report "repRelord"
            
            Dim RepSubReport1 As New CRAXDRT.Report     'instance for subreport Bill
            Dim RepSubReport2 As New CRAXDRT.Report     'instance for subreport Rodetail
            
  'RO master
            recRoMaster.Open "select ROMASTER_ID,(RONO & '/' & right(FINYEAR,2) & '/' & right(FINYEAR+1,2)) AS RNo, format(RODATE,'dd-mm-yy')as Rdate from romaster where Rodate between #" & CStr(frmDate) & "# and #" & CStr(toDate) & "# ", cn, adOpenForwardOnly, adLockReadOnly
                
  'publication bill
            recPubRecord.Open "SELECT romaster_id,publication, billdate, billno, grossamount, commission from publicationbill", cn, adOpenStatic, adLockReadOnly
            
  'Bill record
            recBillRecord.Open "SELECT BM.ROMASTER_ID as Roid,BM.BILLDATE,  (BM.BILLNO & '/' & right(BM.FINYEAR,2) & '/' & right(BM.FINYEAR +1,2)) as billnumber ,BD.AMOUNT as billamount,BM.CANCELLED as billcancelled,BM.REF  FROM BILLMASTER BM, BILLDETAIL BD WHERE BM.BILLMASTER_ID=BD.BILLMASTER_ID and bm.ref='r'", cn, adOpenStatic, adLockReadOnly
                                                'rec("billno") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2)
  'Ro detail
            recRoDetail.Open "select RO.*, Rm.client,  Rm.cancelled from RoDetail RO, Romaster RM where RO.romaster_id=Rm.romaster_id", cn, adOpenForwardOnly, adLockReadOnly
             
             
            Set RepSubReport1 = RepRelOrder.Subreport1.OpenSubreport    'opening subreport bill
            RepSubReport1.Database.SetDataSource recBillRecord          'assigning recordset
                
            Set RepSubReport2 = RepRelOrder.Subreport2.OpenSubreport    'opening subreport rodetail
            RepSubReport2.Database.SetDataSource recRoDetail            'assigning recordset
            
            RepRelOrder.Database.SetDataSource recRoMaster, , 1
            RepRelOrder.Database.SetDataSource recPubRecord, , 2
            
            'assigning value in formula of report
            For i = 1 To RepSubReport1.FormulaFields.Count
               If RepSubReport1.FormulaFields(i).Name = "{@ServiceTaxValue}" Then RepSubReport1.FormulaFields(i).Text = ServiceTax
            Next

            For i = 1 To RepSubReport1.FormulaFields.Count
                If RepSubReport1.FormulaFields(i).Name = "{@Commission}" Then RepSubReport1.FormulaFields(i).Text = Commission
            Next

            frmrep.CRViewer1.ReportSource = RepRelOrder
            frmrep.CRViewer1.PrintReport
            frmrep.CRViewer1.Zoom 100
            frmrep.Caption = "Release Order Register"
            
    Case Is = "pubregister"
            Set Rep = repPubReg
            If rec.State = 1 Then rec.Close
            Dim str As String
           
            str = "select publicationbill.*, (romaster.rono & '/' & right(romaster.finyear,2) & '/0' & right(romaster.finyear,2) +1) as ReleaseNo from publicationbill, romaster " _
                & "Where publicationbill.romaster_id = romaster.romaster_id " _
                & " and billdate between #" & CStr(frmDate) & "# and #" & CStr(toDate) & "#"
             
'            str = "select publicationbill.*, romaster.rono from publicationbill, romaster " _
                & "Where publicationbill.romaster_id = romaster.romaster_id " _
                & " and billdate between #" & CStr(frmDate) & "# and #" & CStr(toDate) & "#"
            'str = "select * from publicationbill where billdate between #" & CStr(frmDate) & "# and #" & CStr(toDate) & "#"
            rec.Open str, cn, adOpenStatic, adLockReadOnly
            'displaying report
            
            Rep.DiscardSavedData        'it helps if the report load more than 1 times. it display always new records according to selected date.
            Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr
            Unload frmrep
            
            frmrep.CRViewer1.ReportSource = Rep
            frmrep.CRViewer1.PrintReport
            frmrep.CRViewer1.Zoom 100
            frmrep.Caption = "Publication Register"
            
            Set rec = Nothing
            
End Select


Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub
