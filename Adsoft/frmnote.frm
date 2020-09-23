VERSION 5.00
Begin VB.Form frmnote 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debit\Credit Note"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   8745
   Begin VB.CommandButton cmdprint 
      Caption         =   "&Print"
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
      Left            =   5205
      TabIndex        =   21
      ToolTipText     =   "Print this RO"
      Top             =   6495
      Width           =   1215
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   5
      Left            =   6735
      MaxLength       =   15
      TabIndex        =   12
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   4
      Left            =   6735
      MaxLength       =   15
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   5
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   13
      Top             =   4920
      Width           =   5655
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   3
      Left            =   6735
      MaxLength       =   15
      TabIndex        =   16
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   4
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   15
      Top             =   5400
      Width           =   5655
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   3
      Left            =   1080
      MaxLength       =   255
      TabIndex        =   11
      Top             =   4440
      Width           =   5655
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   0
      Left            =   6720
      MaxLength       =   15
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   1
      Left            =   6720
      MaxLength       =   15
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   1
      Left            =   1065
      MaxLength       =   255
      TabIndex        =   7
      Top             =   3480
      Width           =   5655
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   405
      Index           =   2
      Left            =   6720
      MaxLength       =   15
      TabIndex        =   10
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   2
      Left            =   1065
      MaxLength       =   255
      TabIndex        =   9
      Top             =   3960
      Width           =   5655
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   0
      Left            =   1065
      MaxLength       =   255
      TabIndex        =   5
      Top             =   3000
      Width           =   5655
   End
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   6735
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   17
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtto 
      Height          =   375
      Left            =   1305
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtrefdate 
      Height          =   315
      Left            =   1305
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1245
      Width           =   2295
   End
   Begin VB.OptionButton optdebit 
      Caption         =   "Debit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6765
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.OptionButton optcredit 
      Caption         =   "Credit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5370
      TabIndex        =   1
      Top             =   720
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtrefsuffix 
      Height          =   315
      Left            =   3585
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   26
      Top             =   720
      Width           =   1125
   End
   Begin VB.TextBox txtrefprefix 
      Height          =   315
      Left            =   1305
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   25
      Text            =   "MAZDA/NGP/"
      Top             =   720
      Width           =   1185
   End
   Begin VB.CommandButton cmdexit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
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
      Left            =   6645
      TabIndex        =   22
      Top             =   6495
      Width           =   1215
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "&Delete"
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
      Left            =   3765
      TabIndex        =   20
      Top             =   6495
      Width           =   1215
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "&Update"
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
      Left            =   2325
      TabIndex        =   19
      Top             =   6495
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "&New"
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
      Left            =   885
      TabIndex        =   18
      Top             =   6495
      Width           =   1215
   End
   Begin VB.ComboBox cmbrefno 
      Height          =   315
      Left            =   2475
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1125
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   38
      Top             =   5040
      Width           =   135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   720
      TabIndex        =   37
      Top             =   5520
      Width           =   135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   720
      TabIndex        =   36
      Top             =   4560
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No."
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
      Left            =   345
      TabIndex        =   35
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
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
      Left            =   3360
      TabIndex        =   34
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   6960
      TabIndex        =   33
      Top             =   2520
      Width           =   675
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   705
      TabIndex        =   32
      Top             =   3600
      Width           =   135
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   705
      TabIndex        =   31
      Top             =   4080
      Width           =   135
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   705
      TabIndex        =   30
      Top             =   3120
      Width           =   135
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   29
      Top             =   6000
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To,"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   345
      TabIndex        =   28
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Dated"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   345
      TabIndex        =   27
      Top             =   1245
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debit\Credit Note"
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
      Left            =   3277
      TabIndex        =   24
      Top             =   0
      Width           =   2190
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ref:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   345
      TabIndex        =   23
      Top             =   720
      Width           =   345
   End
End
Attribute VB_Name = "frmnote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbrefno_Click()
Dim rec As New ADODB.Recordset
Dim srno As Long

EmptyControls

If cmbrefno.ListIndex < 0 Then Exit Sub

rec.Open "select * from notesmaster where srno=" & cmbrefno.ItemData(cmbrefno.ListIndex), cn, adOpenStatic, adLockReadOnly

If rec.BOF And rec.EOF Then
    Exit Sub
End If

txtrefdate = Format(rec("refdate"), "dd-mm-yyyy")

txtto = rec("to")

If rec("notetype") = "c" Then
    optcredit.Value = True
Else
    optdebit.Value = True
End If

srno = rec("srno")

rec.Close
rec.Open "select * from notesdetail where srno=" & srno, cn, adOpenStatic, adLockReadOnly

i = 0

While rec.EOF = False
    txtparticulars(i) = rec("particular")
    txtamount(i) = rec("amount")
    i = i + 1
    rec.MoveNext
Wend

End Sub

Private Sub cmddelete_Click()
On Error GoTo eror

Dim msg As Integer
If cmbrefno.ListIndex < 0 Then Exit Sub 'no credit/debit note selected
msg = MsgBox("Do you really want to delete Note " & txtrefprefix & cmbrefno & txtrefsuffix & "?", vbYesNo, "Confirm Delete")

If msg = vbYes Then
    cn.Execute "delete * from notesmaster where srno=" & cmbrefno.ItemData(cmbrefno.ListIndex)
    cn.Execute "delete * from notesdetail where srno=" & cmbrefno.ItemData(cmbrefno.ListIndex)
    MsgBox "Note deleted from the Database.", vbInformation, "Deleted!!"
    addRefNo
End If

Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
Dim rdate As Date 'ref date
Dim notetype As String 'credit or debit
Dim newsrno As Long 'srno of notesmaster to be entered into notes detail table
Dim temprec As New ADODB.Recordset
 
 If cmdnew.Caption = "&Save" Then
 
 If IsValid = False Then Exit Sub
 
 rdate = Split(txtrefdate, "-")(1) & "/" & Split(txtrefdate, "-")(0) & "/" & Split(txtrefdate, "-")(2)
 
 On Error GoTo errorHandler
 If optcredit.Value = True Then notetype = "c"
 If optdebit.Value = True Then notetype = "d"
     
 'saving in master table
 
 cn.BeginTrans 'beginning transaction
 cn.Execute "insert into notesmaster(refno,finyear,notetype,refdate,to)values('" & cmbrefno & "','" & finyear(rdate) & "','" & notetype & "','" & rdate & "','" & Trim(txtto) & "')"
 'saving in detail table
 
' temprec.Close
 temprec.Open "select max(srno) as maxsrno from notesmaster", cn, adOpenStatic, adLockReadOnly
 
 newsrno = temprec("maxsrno")
 
 For i = 0 To txtparticulars.UBound
     If Trim(txtparticulars(i)) <> "" Then
        cn.Execute "insert into notesdetail (srno,particular,amount)values ('" & newsrno & "','" & Trim(txtparticulars(i)) & "','" & Val(txtamount(i)) & "')"
     End If
 Next
 cn.CommitTrans 'committing transaction
 
 MsgBox "Note Saved into Database.", vbInformation
 
 addRefNo
 cmbrefno.ListIndex = cmbrefno.ListCount - 1
  
 cmdprint.Enabled = True
 cmddelete.Enabled = True
 cmdupdate.Enabled = True
 cmdnew.Caption = "&New"
  
Exit Sub

errorHandler:

MsgBox "Error No. " & Err.Number & vbCrLf & Err.Description, vbCritical
'deleting records in case of error
'cn.Execute "delete * from notesmaster where srno=" & newsrno
'cn.Execute "delete * from notesdetail where srno=" & newsrno
cn.RollbackTrans
Else

 temprec.Open "select max(refno) as maxno from notesmaster where finyear=" & finyear(Date), cn, adOpenStatic, adLockReadOnly
 
 If (temprec.BOF And temprec.EOF) Or IsNull(temprec("maxno")) Then
    cmbrefno = 1
 Else
    cmbrefno = temprec("maxno") + 1
 End If
    
    EmptyControls
    txtrefdate = Format(Date, "dd-mm-yyyy")
    cmdprint.Enabled = False
    cmddelete.Enabled = False
    cmdupdate.Enabled = False
    cmdnew.Caption = "&Save"

End If

End Sub

Private Sub cmdprint_Click()
On Error GoTo eror

Dim rec As New ADODB.Recordset
Dim Rep As New repnotes  'creating instance of crystal reports dsr
Dim frmrep As New frmreports

If cmbrefno.ListIndex < 0 Then Exit Sub

'opening recordset for ro report
rec.Open "SELECT notesmaster.*, notesdetail.* FROM notesDetail RIGHT JOIN notesMaster ON (notesDetail.srno = notesMaster.srno) where notesmaster.srno=" & cmbrefno.ItemData(cmbrefno.ListIndex), cn, adOpenStatic, adLockReadOnly

Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr

frmrep.CRViewer1.ReportSource = Rep
frmrep.CRViewer1.PrintReport
frmrep.Caption = "DEBIT/CREDIT NOTE"

Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdupdate_Click()
Dim rdate As Date, rtype As String
On Error GoTo eror

Dim msg As Integer
If cmbrefno.ListIndex < 0 Then Exit Sub 'no credit/debit note selected

If IsValid = False Then Exit Sub
rdate = Split(txtrefdate, "-")(1) & "/" & Split(txtrefdate, "-")(0) & "/" & Split(txtrefdate, "-")(2)

If optcredit.Value = True Then rtype = "c"
If optdebit.Value = True Then rtype = "d"

msg = MsgBox("Do you want to modify Note " & txtrefprefix & cmbrefno & txtrefsuffix & "?", vbYesNo, "Confirm Update")

If msg = vbYes Then
    'updating master table
   
    cn.Execute "update notesmaster set refno=" & cmbrefno & ",finyear=" & finyear(rdate) & ",refdate='" & rdate & "',to='" & Trim(txtto) & "',notetype='" & rtype & "' where srno=" & cmbrefno.ItemData(cmbrefno.ListIndex)
    'deleting related records from detail table
    cn.Execute "delete * from notesdetail where srno=" & cmbrefno.ItemData(cmbrefno.ListIndex)
    'inserting records in detail table again for the same entry
    For i = 0 To txtparticulars.UBound
         If Trim(txtparticulars(i)) <> "" Then
            cn.Execute "insert into notesdetail (srno,particular,amount) values ('" & cmbrefno.ItemData(cmbrefno.ListIndex) & "','" & Trim(txtparticulars(i)) & "','" & Val(txtamount(i)) & "')"
        End If
    Next
    
    MsgBox "Note modified in the Database.", vbInformation, "Updated!!"
End If

Exit Sub
eror:
MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
addRefNo
End Sub

Function IsValid() As Boolean
 On Error GoTo errorHandler

 If Val(cmbrefno) <= 0 Then
 IsValid = False
 MsgBox "Please enter a valid Ref No.", vbCritical, "Invalid Entry"
 cmbrefno.SetFocus
 Exit Function
 End If
 
 If optcredit.Value = False And optdebit.Value = False Then
 IsValid = False
 MsgBox "Please select Credit or Debit.", vbCritical, "Invalid Entry"
 Exit Function
 End If
 
 
 If IsDate(Split(txtrefdate, "-")(1) & "-" & Split(txtrefdate, "-")(0) & "-" & Split(txtrefdate, "-")(2)) = False Then
 IsValid = False
 MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
 SetFocus
 Exit Function
 End If
 
 If Trim(txtto) = "" Then
 MsgBox "Please enter a Client Name.", vbCritical, "Incomplete Entry"
 txtto.SetFocus
 IsValid = False
 Exit Function
 End If

 If Trim(txtparticulars(0)) = "" Then
 MsgBox "Please enter Particulars.", vbCritical, "Incomplete Entry"
 txtparticulars(0).SetFocus
 IsValid = False
 Exit Function
 End If

 If Val(txttotal) <= 0 Then
 MsgBox "Please enter Amount.", vbCritical, "Incomplete Entry"
 IsValid = False
 Exit Function
 End If

IsValid = True

Exit Function

errorHandler:
IsValid = False
MsgBox "Invalid Entry!!" & vbCrLf & Err.Description, vbCritical

End Function

Sub EmptyControls()

optcredit.Value = False
optdebit.Value = False

txtrefdate = ""
txtto = ""
txttotal = ""


For i = 0 To txtparticulars.UBound
    txtparticulars(i) = ""
    txtamount(i) = ""
Next

End Sub

Private Sub txtamount_Change(Index As Integer)
Dim ttl As Single

For i = 0 To txtamount.UBound
    ttl = ttl + Val(txtamount(i))
Next

txttotal = ttl
End Sub

Private Sub txtamount_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub 'allows backspace and .
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtrefdate_Change()
Dim rdate As String
If Len(txtrefdate) <> 10 Then Exit Sub

rdate = Split(txtrefdate, "-")(1) & "/" & Split(txtrefdate, "-")(0) & "/" & Split(txtrefdate, "-")(2)

If IsDate(rdate) = True Then
    txtrefsuffix = "/" & Right(finyear(CDate(rdate)), 2) & "/" & Right(finyear(CDate(rdate)) + 1, 2)
Else
    txtrefsuffix = ""
End If
End Sub

Private Sub txtrefdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtrefdate) = 2 Or Len(txtrefdate) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtrefdate) = 2 Or Len(txtrefdate) = 5) Then
    KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
    Beep
    Exit Sub
End If
        
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Sub addRefNo()
Dim rec As New ADODB.Recordset

cmbrefno.Clear

'adding names in combo
rec.Open "select srno,refno,finyear from notesmaster order by finyear,refno", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If

While rec.EOF = False
    cmbrefno.AddItem rec("refno")
    cmbrefno.ItemData(cmbrefno.ListCount - 1) = rec("srno")
    rec.MoveNext
Wend

'select last item
If cmbrefno.ListCount > 0 Then cmbrefno.ListIndex = cmbrefno.ListCount - 1

End Sub

