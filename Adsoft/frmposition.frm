VERSION 5.00
Begin VB.Form frmposition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6225
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   2640
      MaxLength       =   100
      TabIndex        =   7
      Top             =   1200
      Width           =   2775
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
      TabIndex        =   3
      Top             =   1995
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
      TabIndex        =   4
      Top             =   1995
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
      TabIndex        =   5
      Top             =   1995
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
      TabIndex        =   6
      Top             =   1995
      Width           =   1215
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
      Left            =   840
      TabIndex        =   8
      Top             =   1320
      Width           =   1635
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Position"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
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
      Left            =   2595
      TabIndex        =   0
      Top             =   0
      Width           =   1020
   End
End
Attribute VB_Name = "frmposition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrid() As Long 'stores id field of the table

Private Sub cmbname_Click()
If cmbname = "" Then Exit Sub
txtname = cmbname
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim msg As Integer
Dim rec As New ADODB.Recordset
rec.Open "select positions from ratecard where positions=" & arrid(cmbname.ListIndex), cn, adOpenStatic, adLockReadOnly
If rec.BOF = True And rec.EOF = True Then
msg = MsgBox("Do you really want to delete Position " & cmbname & "?", vbYesNo, "Confirm Delete")
If msg = vbYes Then
cn.Execute "delete * from positions where position_id=" & arrid(cmbname.ListIndex)
MsgBox "Position deleted from the Database.", vbInformation, "Deleted!!"
RefreshCombo
End If
Else
MsgBox "Position " & cmbname & " has related entries in Database.", vbCritical, "Deletion Failed!!"
End If
End Sub

Private Sub cmdnew_Click()
If cmdnew.Caption = "&New" Then 'new record
 cmddelete.Enabled = False
 cmdupdate.Enabled = False
 cmbname.ListIndex = -1
 cmbname.Enabled = False
 txtname = ""
 txtname.SetFocus
 cmdnew.Caption = "&Add"

Else 'add record
 If IsValid = False Then Exit Sub
 cn.Execute "insert into positions (name) values ('" & Trim(txtname) & "')"
 MsgBox "Position Added into Database.", vbInformation
 RefreshCombo
 
 cmddelete.Enabled = True
 cmdupdate.Enabled = True
 cmbname.Enabled = True
 cmdnew.Caption = "&New"
End If

End Sub

Private Sub cmdupdate_Click()
If IsValid = False Then Exit Sub
 cn.Execute "update positions set name='" & Trim(txtname) & "' where Position_id=" & arrid(cmbname.ListIndex)
 MsgBox "Position modified in Database.", vbInformation
 RefreshCombo
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50

RefreshCombo
End Sub

Sub RefreshCombo() ' this adds names in combobox
Dim rec As New ADODB.Recordset
cmbname.Clear
Erase arrid
ReDim arrid(0)
'adding names in combo
rec.Open "select position_id,name from positions", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
MsgBox "Please add a Position.", vbCritical, "No Position Found"
cmdupdate.Enabled = False
cmddelete.Enabled = False
Exit Sub
End If

cmbname.AddItem rec!Name
arrid(0) = rec!position_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arrid(UBound(arrid) + 1)
cmbname.AddItem rec!Name
arrid(UBound(arrid)) = rec!position_id
rec.MoveNext
Wend

cmbname.ListIndex = 0
End Sub

Function IsValid() As Boolean
 
 If Trim(txtname) = "" Then
 MsgBox "Please enter a Position Name.", vbCritical, "Blank Position"
 txtname.SetFocus
 IsValid = False
 Exit Function
 End If
  
IsValid = True

End Function
