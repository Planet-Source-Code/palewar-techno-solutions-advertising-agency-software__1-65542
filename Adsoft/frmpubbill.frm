VERSION 5.00
Begin VB.Form frmpubbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Publication Bills"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7305
   Begin VB.ComboBox cmbrono 
      Height          =   315
      Left            =   5160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1680
   End
   Begin VB.TextBox txtfolino 
      Height          =   390
      Left            =   5160
      MaxLength       =   5
      TabIndex        =   8
      Top             =   3480
      Width           =   1680
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
      Left            =   1042
      TabIndex        =   9
      Top             =   4155
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
      Left            =   2377
      TabIndex        =   10
      Top             =   4155
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
      Left            =   3712
      TabIndex        =   11
      Top             =   4155
      Width           =   1215
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
      Left            =   5047
      TabIndex        =   12
      Top             =   4155
      Width           =   1215
   End
   Begin VB.TextBox txtnet 
      Height          =   390
      Left            =   1897
      Locked          =   -1  'True
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3465
      Width           =   1680
   End
   Begin VB.TextBox txtcommission 
      Height          =   390
      Left            =   5197
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2760
      Width           =   1680
   End
   Begin VB.TextBox txtpublication 
      Height          =   390
      Left            =   1897
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1410
      Width           =   4965
   End
   Begin VB.TextBox txtbillno 
      Height          =   390
      Left            =   1897
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2070
      Width           =   1680
   End
   Begin VB.TextBox txtgross 
      Height          =   390
      Left            =   1897
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2760
      Width           =   1680
   End
   Begin VB.TextBox txtbilldate 
      Height          =   390
      Left            =   5197
      MaxLength       =   10
      TabIndex        =   4
      Top             =   2070
      Width           =   1680
   End
   Begin VB.ComboBox cmbbillid 
      Height          =   315
      Left            =   1897
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   855
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ro.No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FolioNo."
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
      Index           =   4
      Left            =   3960
      TabIndex        =   21
      Top             =   3480
      Width           =   765
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
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
      Index           =   3
      Left            =   420
      TabIndex        =   20
      Top             =   3525
      Width           =   1035
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commission"
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
      Left            =   3975
      TabIndex        =   19
      Top             =   2835
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
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
      Left            =   420
      TabIndex        =   18
      Top             =   1455
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Publications Bills"
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
      Left            =   2587
      TabIndex        =   17
      Top             =   105
      Width           =   2130
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
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
      Left            =   420
      TabIndex        =   16
      Top             =   2115
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amount"
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
      Index           =   2
      Left            =   420
      TabIndex        =   15
      Top             =   2805
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   3975
      TabIndex        =   14
      Top             =   2100
      Width           =   750
   End
   Begin VB.Label Label2 
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
      Height          =   255
      Left            =   420
      TabIndex        =   13
      Top             =   900
      Width           =   615
   End
End
Attribute VB_Name = "frmpubbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset   'retrieve SrNo
Dim valBillID As Long           'maximum BillId
Dim valrono As String           'release order no

'************************************Anish************************************************
Dim arrRoId() As Long           'array, contains the Release OrderId ID

Private Sub cmbbillid_Click() 'retrieve values from PublicationBill
Dim rec As New ADODB.Recordset
Dim i As Long
On Error GoTo eror
If cmbbillid.Text = "" Then Exit Sub
    If rec.State = 1 Then rec.Close
    rec.Open "select * from publicationbill where bill_id=" & cmbbillid.Text & "", cn, adOpenForwardOnly, adLockReadOnly
        If (rec.BOF = True) Or (rec.EOF = True) Then
            Exit Sub
        End If
'***************************************Anish**********************************************
'display values in controls
For i = 0 To UBound(arrRoId) - 1
    If arrRoId(i) = rec("RoMaster_Id") Then
        cmbrono.ListIndex = i
        Exit For
    End If
Next
'******************************************************************************************

        txtpublication = rec!publication
        txtbilldate = Format(rec!billdate, "dd-mm-yyyy")
        txtgross = rec!grossamount
        txtbillno = rec!billno
        txtcommission = rec!commission
        txtfolino = rec!FolioNo
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmddelete_Click() 'Delete record in database
On Error GoTo eror
If Not cmbbillid.Text = "" Then
    If MsgBox("Do you want to Delete this record", vbQuestion + vbYesNo, "confirmation") = vbYes Then
        cn.Execute "delete * from publicationbill where bill_id=" & cmbbillid.Text & ""
        MsgBox "The Record is deleted", vbInformation, "Information"
        cmbbillid.Clear
        Call addSrNo
        If cmbbillid.ListCount > 0 Then
            cmbbillid.ListIndex = cmbbillid.ListCount - 1
        End If
    End If
End If
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
Dim rodate As Date
Dim valroyear As String
On Error GoTo eror
If cmdnew.Caption = "&Save" Then
        If IsValid = False Then Exit Sub
        rodate = Split(txtbilldate, "-")(1) & "/" & Split(txtbilldate, "-")(0) & "/" & Split(txtbilldate, "-")(2)

'        valroyear = cmbrono
'****************************************Anish*********************************************
        'Saving in Publicationbill
        cn.Execute "insert into publicationbill (BILLNO, ROMASTER_ID, BILLDATE, PUBLICATION, GROSSAMOUNT, COMMISSION, FOLIONO) values ('" & Val(txtbillno) & "'," & arrRoId(cmbrono.ListIndex) & ",'" & rodate & "','" & txtpublication & "'," & Val(txtgross) & "," & Val(txtcommission) & "," & Val(txtfolino) & ")"
        
'******************************************************************************************
'        cn.Execute "insert into publicationbill values (" & valBillID & ",'" & Val(txtbillno) & "','" & Trim(valroyear) & "','" & rodate & "','" & txtpublication & "'," & Val(txtgross) & "," & Val(txtcommission) & "," & Val(txtfolino) & ")"
        MsgBox "The Publication Information is Saved into Database", vbInformation, "Information"
        cmdnew.Caption = "&New"
        cmdupdate.Enabled = True
        cmddelete.Enabled = True
        cmbbillid.Enabled = True
        cmbbillid.Clear
        Call addSrNo
        cmbbillid.ListIndex = cmbbillid.ListCount - 1
Else

        Call BlankControl   'Clear the control
        cmbbillid.ListIndex = -1
        cmbbillid.Enabled = False
        cmdnew.Caption = "&Save"
        cmdupdate.Enabled = False
        cmddelete.Enabled = False
        cmbrono.SetFocus
End If
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub
Function IsValid() As Boolean   'Check the validation
If cmbrono.Text = "" Then
    IsValid = False
    MsgBox "Please select the Ro No.", vbCritical, "Incomplete Entry"
    cmbrono.SetFocus
    Exit Function
End If

If txtpublication = "" Then
    IsValid = False
    MsgBox "Please enter the Publication", vbCritical, "Incomplete Entry"
    txtpublication.SetFocus
    Exit Function
End If

If IsDate(Split(txtbilldate, "-")(1) & "-" & Split(txtbilldate, "-")(0) & "-" & Split(txtbilldate, "-")(2)) = False Then
    IsValid = False
    MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
    txtbilldate.SetFocus
    Exit Function
End If
 
If txtbillno = "" Then
    IsValid = False
    MsgBox "Please enter the Bill No.", vbCritical, "Incomplete Entry"
    txtbillno.SetFocus
    Exit Function
End If

If txtgross = "" Then
    IsValid = False
    MsgBox "Please enter the Gross Amount", vbCritical, "Incomplete Entry"
    txtgross.SetFocus
    Exit Function
End If

If txtcommission = "" Then
    IsValid = False
    MsgBox "Please enter the Commission", vbCritical, "Incomplete Entry"
    txtcommission.SetFocus
    Exit Function
End If

If txtfolino = "" Then
    IsValid = False
    MsgBox "Please enter the FolioNo", vbCritical, "Incomplete Entry"
    txtfolino.SetFocus
    Exit Function
End If

IsValid = True
End Function

Private Sub BlankControl()  'Clear the control
    cmbrono.ListIndex = -1
    txtpublication = ""
    txtbillno = ""
    txtbilldate = ""
    txtgross = ""
    txtcommission = ""
    txtfolino = ""
    txtnet = ""
End Sub

Private Sub cmdupdate_Click() 'Update the record
Dim billdate As Date
On Error GoTo eror
If IsValid = False Then Exit Sub    'Check the validation
    billdate = Split(txtbilldate, "-")(1) & "/" & Split(txtbilldate, "-")(0) & "/" & Split(txtbilldate, "-")(2)

'****************************************Anish*******************************************
    'update record in publicationbill table
    cn.Execute "update publicationbill set billno=" & Val(txtbillno) & ", romaster_id=" & arrRoId(cmbrono.ListIndex) & ", billdate='" & billdate & "',publication='" & txtpublication & "', grossamount=" & txtgross & ",commission=" & txtcommission & ",foliono=" & txtfolino & " where bill_id=" & cmbbillid.Text & ""
    
'****************************************************************************************
    
    'cn.Execute "update publicationbill set billno=" & txtbillno & ",ronoyear='" & cmbrono.Text & "', billdate='" & billdate & "',publication='" & txtpublication & "', grossamount=" & txtgross & ",commission=" & txtcommission & ",foliono=" & txtfolino & " where bill_id=" & cmbbillid.Text & ""
    MsgBox "Entry modified in Database.", vbInformation
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call addSrNo                     'retrieve SrNo from publicationbill
Call addRoNo                     'retrieve RoNo from romaster
If cmbbillid.ListCount > 0 Then
    cmbbillid.ListIndex = 0
End If
End Sub
Private Sub addRoNo()   'retrieve RoNo from romaster
Dim rsRoNo As New ADODB.Recordset
On Error GoTo eror
cmbrono.Clear


'*********************************Anish*************************************************
ReDim arrRoId(0)
rsRoNo.Open "select romaster_id,rono,finyear from romaster where cancelled=false order by rono,finyear", cn, adOpenForwardOnly, adLockReadOnly

If rsRoNo.BOF = True And rsRoNo.EOF = True Then
    Exit Sub
End If

'adding RoNo in combo
While rsRoNo.EOF = False
    arrRoId(UBound(arrRoId)) = rsRoNo("romaster_id")
    valrono = rsRoNo!RoNo & "/" & Right(rsRoNo!finyear, 2) & "/" & Right(rsRoNo!finyear + 1, 2)
    cmbrono.AddItem valrono
    valrono = ""
    ReDim Preserve arrRoId(UBound(arrRoId) + 1)
    rsRoNo.MoveNext
Wend
'*******************************************************************************************
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub
Private Sub addSrNo()   'retrieve SrNo from publicationbill
On Error GoTo eror
If rs.State = 1 Then rs.Close
rs.Open "select Bill_id from publicationbill order by Bill_id", cn, adOpenForwardOnly, adLockReadOnly
    If (rs.BOF = True) Or (rs.EOF = True) Then
        Exit Sub
    End If
    'adding SrNo in combo
    While rs.EOF = False
         cmbbillid.AddItem rs!Bill_id
         rs.MoveNext
    Wend
Exit Sub
eror:
MsgBox Err.Description, vbCritical
End Sub
Private Sub txtbilldate_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtbilldate) = 2 Or Len(txtbilldate) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtbilldate) = 2 Or Len(txtbilldate) = 5) Then
    KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
    Beep
    Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then  'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtbillno_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then         'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtcommission_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then         'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtfolino_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub                   'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then          'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub
Private Sub txtcommission_Change()              'calculate the amount
If Not (txtgross = "" And txtcommission = "") Then
    txtnet = Val(txtgross) - Val(txtcommission)
End If
End Sub
Private Sub txtgross_Change()                   'calculate the amount
If Not (txtgross = "" And txtcommission = "") Then
    txtnet = Val(txtgross) - Val(txtcommission)
End If
End Sub

Private Sub txtgross_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub                   'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then          'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub

Private Sub txtnet_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub  'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then          'allows numbers
    KeyAscii = 0
    Beep
End If
End Sub
