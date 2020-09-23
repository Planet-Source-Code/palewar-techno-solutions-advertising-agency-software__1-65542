VERSION 5.00
Begin VB.Form frmReleaseOrder 
   Caption         =   "Select Date"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   2715
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Cancel Current RO"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Print this RO"
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox txtdatefrom 
      Height          =   315
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   1
      Top             =   480
      Width           =   1380
   End
   Begin VB.TextBox txtdateto 
      Height          =   315
      Left            =   1530
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1185
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1200
      TabIndex        =   7
      Top             =   0
      Width           =   870
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   6
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   450
      TabIndex        =   5
      Top             =   1185
      Width           =   735
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Ex. 26-03-1978)"
      Height          =   195
      Left            =   1515
      TabIndex        =   4
      Top             =   810
      Width           =   1170
   End
End
Attribute VB_Name = "frmReleaseOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdcancel_Click()
Unload Me
End Sub
Private Sub cmdprint_Click()
On Error GoTo eror
Dim datefrom As Date, dateto As Date
Dim rec As New ADODB.Recordset          'recordset to act as report source
Dim Rep As New reppub           'creating instance of crystal report dsr
Dim frmrep As New frmreports            'new instance of report form

'converting entered date to mm-dd-yyyy format for use in code
datefrom = Split(txtdatefrom, "-")(1) & "/" & Split(txtdatefrom, "-")(0) & "/" & Split(txtdatefrom, "-")(2)
dateto = Split(txtdateto, "-")(1) & "/" & Split(txtdateto, "-")(0) & "/" & Split(txtdateto, "-")(2)

'validating dates
If IsDate(datefrom) = False Then
    MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
    txtdatefrom.SetFocus
    Exit Sub
End If

If IsDate(dateto) = False Then
    MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
    txtdateto.SetFocus
    Exit Sub
End If

''open recordset using entered date
'rec.Open "SELECT BM.REFNO, RD.RELEASEDATE, RD.CAPTION, BM.BILLDATE, BM.BILLNO, " _
'        & "BD.AMOUNT, (VAL(BD.AMOUNT) *VAL( .15) *VAL( .05)) AS SERVICE_TAX, " _
'        & "(VAL(BD.AMOUNT) +VAL( SERVICE_TAX)) AS TOTAL_AMOUNT, PB.PUBLICATION, PB.BILLDATE," _
'        & "PB.BILLNO, PB.GROSSAMOUNT, PB.COMMISSION, (PB.GROSSAMOUNT-PB.COMMISSION) AS NETAMOUNT " _
'        & "FROM BILLDETAIL AS BD, BILLMASTER AS BM, PUBLICATIONBILL AS PB, " _
'        & "RODETAIL AS RD " _
'        & "Where(RD.RONO & "/" & Right(RD.finyear, 2) & "/0" & Right(RD.finyear, 2) + 1) = PB.RONOYEAR " _
'        & "AND PB.RONOYEAR=BM.REFNO " _
'        & "AND BM.BILLNO=BD.BILLNO " _
'        & "AND BM.REF='R' " _
'        & "AND RD.RELEASEDATE BETWEEN #01-01-02# AND #12-02-03# "

rec.Open "SELECT BM.REFNO, RD.RELEASEDATE, RD.CAPTION, BM.BILLDATE, BM.BILLNO,BD.AMOUNT, (VAL(BD.AMOUNT) *VAL( .15) *VAL( .05)) AS SERVICE_TAX,(VAL(BD.AMOUNT) +VAL( SERVICE_TAX)) AS TOTAL_AMOUNT, PB.PUBLICATION, PB.BILLDATE,PB.BILLNO, PB.GROSSAMOUNT, PB.COMMISSION, (PB.GROSSAMOUNT-PB.COMMISSION) AS NETAMOUNT FROM BILLDETAIL AS BD, BILLMASTER AS BM, PUBLICATIONBILL AS PB,RODETAIL AS RD Where(RD.RONO '/' Right(RD.finyear, 2)  '/0' Right(RD.finyear, 2) + 1) = PB.RONOYEAR AND PB.RONOYEAR=BM.REFNO AND BM.BILLNO=BD.BILLNO AND BM.REF='R' AND RD.RELEASEDATE BETWEEN #01-01-02# AND #12-02-03# ", cn, adOpenStatic, adLockReadOnly






''rec.Open "SELECT billmaster.finyear, billmaster.billno, First(billmaster.ref) " _
''& "AS FirstOfref, First(billmaster.billdate) AS FirstOfbilldate, " _
''& "First(billmaster.client) AS FirstOfclient, Sum(billdetail.amount) " _
''& "AS SumOfamount FROM billdetail INNER JOIN billmaster ON " _
''& "(billdetail.billno = billmaster.billno) AND (billdetail.finyear " _
''& "= billmaster.finyear) GROUP BY billmaster.finyear, " _
''& "billmaster.billno HAVING First(billmaster.cancelled)=" _
''& "No and First(billmaster.billdate) >=#" & CStr(datefrom) & "# and First" _
''& "(billmaster.billdate)<=#" & CStr(dateto) & "#", cn, adOpenStatic, adLockReadOnly

'displaying report
Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr

frmrep.CRViewer1.ReportSource = Rep
frmrep.CRViewer1.PrintReport
frmrep.Caption = "Release Order Register"

Unload Me

Exit Sub
eror:
MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
lbldate = "(Ex. " & Format(Date, "dd-mm-yyyy") & ")"
End Sub
Private Sub txtdatefrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtdatefrom) = 2 Or Len(txtdatefrom) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtdatefrom) = 2 Or Len(txtdatefrom) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtdateto_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtdateto) = 2 Or Len(txtdateto) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtdateto) = 2 Or Len(txtdateto) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

