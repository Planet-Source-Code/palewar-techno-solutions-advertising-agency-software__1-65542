VERSION 5.00
Begin VB.Form frmpublication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Publication"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   Icon            =   "frmpublication.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6225
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Height          =   1455
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   5175
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   2085
         MaxLength       =   100
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
      Begin VB.ComboBox cmbday 
         Height          =   315
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbmonth 
         Height          =   315
         Left            =   2985
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cmbyear 
         Height          =   315
         Left            =   4005
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lbldate 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Card Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.ComboBox cmbname 
      Height          =   315
      Left            =   2640
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
      Left            =   525
      TabIndex        =   8
      Top             =   3435
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
      Left            =   1845
      TabIndex        =   10
      Top             =   3435
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
      Left            =   3165
      TabIndex        =   12
      Top             =   3435
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
      Left            =   4485
      TabIndex        =   13
      Top             =   3435
      Width           =   1215
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
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
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
      Left            =   2122
      TabIndex        =   0
      Top             =   0
      Width           =   1980
   End
End
Attribute VB_Name = "frmpublication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrid() As Long 'stores id field of the table

Private Sub cmbname_Click()
Dim rec_current As New ADODB.Recordset 'Current Record
Dim cur_date As Date
If cmbname = "" Then Exit Sub
rec_current.Open "select effective_date from publication where publication_id=" & arrid(cmbname.ListIndex), cn, adOpenStatic, adLockReadOnly
txtname = cmbname
cur_date = rec_current!effective_date

rec_current.Close
Set rec_current = Nothing

'setting date in combo boxes from where current rate card is effective
cmbday.ListIndex = DatePart("d", cur_date) - 1
cmbmonth.ListIndex = DatePart("m", cur_date) - 1
cmbyear = DatePart("yyyy", cur_date)
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim msg As Integer
Dim rec As New ADODB.Recordset
rec.Open "select publication from ratecard where publication=" & arrid(cmbname.ListIndex), cn, adOpenStatic, adLockReadOnly
If rec.BOF = True And rec.EOF = True Then ' no check for related records in premium table
msg = MsgBox("Do you really want to delete Publication " & cmbname & "?", vbYesNo, "Confirm Delete")
If msg = vbYes Then 'related records in premium table are automatically deleted by cascade delete option in database
cn.Execute "delete * from publication where publication_id=" & arrid(cmbname.ListIndex)
MsgBox "Publication deleted from the Database.", vbInformation, "Deleted!!"
RefreshCombo
End If
Else
MsgBox "Publication " & cmbname & " has related entries in Database.", vbCritical, "Deletion Failed!!"
End If
End Sub

Private Sub cmdnew_Click()
If cmdnew.Caption = "&New" Then 'new record
 cmddelete.Enabled = False
 cmdupdate.Enabled = False
 cmbname.ListIndex = -1
 cmbday.ListIndex = -1
 cmbmonth.ListIndex = -1
 cmbyear.ListIndex = -1
 cmbname.Enabled = False
 txtname = ""
 txtname.SetFocus
 cmdnew.Caption = "&Add"

Else 'add record
 If IsValid = False Then Exit Sub
 cn.Execute "insert into publication (name,effective_date) values ('" & Trim(txtname) & "','" & cmbmonth & "/" & cmbday & "/" & cmbyear & "')"
 MsgBox "Publication Added into Database.", vbInformation
 RefreshCombo
 
 cmddelete.Enabled = True
 cmdupdate.Enabled = True
 cmbname.Enabled = True
 cmdnew.Caption = "&New"
End If

End Sub

Private Sub cmdupdate_Click()
If IsValid = False Then Exit Sub
 cn.Execute "update publication set name='" & Trim(txtname) & "',effective_date= '" & cmbmonth & "/" & cmbday & "/" & cmbyear & "' where publication_id=" & arrid(cmbname.ListIndex)
 MsgBox "Publication modified in Database.", vbInformation
 RefreshCombo
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50

'adding days in combo box

For i = 1 To 31
    If i < 10 Then
        cmbday.AddItem "0" & i
    Else
        cmbday.AddItem i
    End If
Next

'adding months in combo box
For i = 1 To 12
    cmbmonth.AddItem MonthName(i, True)
Next

'adding year in combo box
For i = 1997 To Year(Date)
    cmbyear.AddItem i
Next

RefreshCombo
End Sub

Sub RefreshCombo() ' this adds names in combobox
Dim rec As New ADODB.Recordset
cmbname.Clear
Erase arrid
ReDim arrid(0)
'adding names in combo
rec.Open "select publication_id,name from publication", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
MsgBox "Please add a Publication.", vbCritical, "No Publication Found"
cmdupdate.Enabled = False
cmddelete.Enabled = False
Exit Sub
End If

cmbname.AddItem rec!Name
arrid(0) = rec!publication_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arrid(UBound(arrid) + 1)
cmbname.AddItem rec!Name
arrid(UBound(arrid)) = rec!publication_id
rec.MoveNext
Wend

cmbname.ListIndex = 0
End Sub

Function IsValid() As Boolean
Dim tempdate As String
 
 If Trim(txtname) = "" Then
 MsgBox "Please enter a Publication Name.", vbCritical, "Blank Publication"
 txtname.SetFocus
 IsValid = False
 Exit Function
 End If
 
 tempdate = cmbmonth & "/" & cmbday & "/" & cmbyear
 
 If IsDate(tempdate) = False Then
 MsgBox "Please select a valid date.", vbCritical, "Invalid Date"
 cmbday.SetFocus
 IsValid = False
 Exit Function
 End If

IsValid = True

End Function
