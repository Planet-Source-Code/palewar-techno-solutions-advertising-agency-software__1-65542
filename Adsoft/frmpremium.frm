VERSION 5.00
Begin VB.Form frmpremium 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Premium"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5850
   Begin VB.TextBox txtpercent 
      Height          =   375
      Left            =   2453
      TabIndex        =   9
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cmbname 
      Height          =   315
      Left            =   2453
      TabIndex        =   8
      Top             =   1440
      Width           =   2775
   End
   Begin VB.ComboBox cmbpublication 
      Height          =   315
      Left            =   2453
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2775
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
      Left            =   338
      TabIndex        =   3
      Top             =   2760
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
      Left            =   1680
      TabIndex        =   4
      Top             =   2760
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
      Left            =   3000
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      Cancel          =   -1  'True
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
      Left            =   4320
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Percent"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   653
      TabIndex        =   10
      Top             =   2040
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Premium Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   653
      TabIndex        =   7
      Top             =   1440
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Publication"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   653
      TabIndex        =   2
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Premium"
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
      Left            =   2348
      TabIndex        =   0
      Top             =   0
      Width           =   1140
   End
End
Attribute VB_Name = "frmpremium"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrpubid() As Long
Dim arrid() As Long
Private Sub cmbname_Click()
Dim rec As New ADODB.Recordset
If cmbname = "" Then Exit Sub
rec.Open "select premium_percent from premium where premium_id=" & arrid(cmbname.ListIndex), cn, adOpenStatic, adLockReadOnly
txtpercent = rec!premium_percent
rec.Close
Set rec = Nothing
End Sub

Private Sub cmbpublication_Click()
If cmbpublication = "" Then Exit Sub
If cmdnew.Caption = "&New" Then
txtpercent = ""
RefreshCombo
End If
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim msg As Integer
If cmbname = "" Then Exit Sub
msg = MsgBox("Do you really want to delete " & cmbname & " Premium?", vbYesNo, "Confirm Delete")
If msg = vbYes Then
cn.Execute "delete * from premium where premium_id=" & arrid(cmbname.ListIndex)
MsgBox "Premium deleted from the Database.", vbInformation, "Deleted!!"
txtpercent = ""
RefreshCombo
End If
End Sub

Private Sub cmdnew_Click()
If cmdnew.Caption = "&New" Then 'new record
 cmddelete.Enabled = False
 cmdupdate.Enabled = False
 cmbpublication.ListIndex = -1
 cmbname = ""
 txtpercent = ""
 cmbname.SetFocus
 cmdnew.Caption = "&Add"

Else 'add record
 If IsValid = False Then Exit Sub
 cn.Execute "insert into premium (publication,premium_name,premium_percent) values (" & arrpubid(cmbpublication.ListIndex) & ",'" & Trim(cmbname) & "','" & txtpercent & "')"
 MsgBox "Premium Added into Database.", vbInformation
 RefreshCombo
 
 cmddelete.Enabled = True
 cmdupdate.Enabled = True
 cmbname.Enabled = True
 cmdnew.Caption = "&New"
End If

End Sub

Private Sub cmdupdate_Click()
If cmbname = "" Then Exit Sub
If IsValid = False Then Exit Sub
 cn.Execute "update premium set premium_name='" & Trim(cmbname) & "',publication=" & arrpubid(cmbpublication.ListIndex) & ",premium_percent='" & txtpercent & "' where premium_id=" & arrid(cmbname.ListIndex)
 MsgBox "Premium modified in Database.", vbInformation
 RefreshCombo
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50
addpubcombo
End Sub

Sub RefreshCombo() ' this adds names in combobox
Dim rec As New ADODB.Recordset
cmbname.Clear
Erase arrid
ReDim arrid(0)
'adding names in combo
If cmbpublication = "" Then Exit Sub
rec.Open "select * from premium where publication=" & arrpubid(cmbpublication.ListIndex), cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
Exit Sub
'MsgBox "Please add a Premium.", vbCritical, "No Premium Found"
'cmdupdate.Enabled = False
'cmddelete.Enabled = False
Exit Sub
End If

cmbname.AddItem rec!premium_name
arrid(0) = rec!premium_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arrid(UBound(arrid) + 1)
cmbname.AddItem rec!premium_name
arrid(UBound(arrid)) = rec!premium_id
rec.MoveNext
Wend

cmbname.ListIndex = 0
End Sub

Function IsValid() As Boolean
 If cmbpublication = "" Then
 MsgBox "Please select a Publication.", vbCritical, "Blank Publication"
 cmbpublication.SetFocus
 IsValid = False
 Exit Function
 End If
 
 If Trim(cmbname) = "" Then
 MsgBox "Please enter a Premium Name.", vbCritical, "Blank Entry"
 cmbname.SetFocus
 IsValid = False
 Exit Function
 End If
  
If IsNumeric(txtpercent) = False Or Val(txtpercent) <= 0 Or Val(txtpercent) > 250 Then
 MsgBox "Please enter a valid Premium Percent.", vbCritical, "Invalid Entry"
 txtpercent.SetFocus
 IsValid = False
 Exit Function
 End If
  
IsValid = True

End Function
Sub addpubcombo()
Dim rec As New ADODB.Recordset
cmbpublication.Clear
'adding names in combo
rec.Open "select publication_id,name from publication", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
frmpublication.Show
frmpublication.ZOrder 0
Exit Sub
End If

Erase arrpubid
ReDim arrpubid(0)
cmbpublication.AddItem rec!Name
arrpubid(0) = rec!publication_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arrpubid(UBound(arrpubid) + 1)
cmbpublication.AddItem rec!Name
arrpubid(UBound(arrpubid)) = rec!publication_id
rec.MoveNext
Wend
cmbpublication.ListIndex = 0
End Sub
