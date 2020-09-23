VERSION 5.00
Begin VB.Form frmratecard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rate Card"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6015
   Begin VB.TextBox txtunit 
      Height          =   375
      Left            =   3975
      MaxLength       =   15
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
   End
   Begin VB.OptionButton optunit 
      Height          =   255
      Index           =   1
      Left            =   3615
      TabIndex        =   6
      Top             =   3000
      Width           =   255
   End
   Begin VB.OptionButton optunit 
      Height          =   255
      Index           =   0
      Left            =   2175
      TabIndex        =   4
      Top             =   3000
      Width           =   255
   End
   Begin VB.ComboBox cmbunit 
      Height          =   315
      ItemData        =   "frmratecard.frx":0000
      Left            =   2535
      List            =   "frmratecard.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtrate 
      Height          =   375
      Left            =   2175
      MaxLength       =   5
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.ComboBox cmbposition 
      Height          =   315
      Left            =   2175
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ComboBox cmbpublication 
      Height          =   315
      Left            =   2175
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   3015
   End
   Begin VB.TextBox txtinfo 
      Height          =   975
      Left            =   2175
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
   End
   Begin VB.ComboBox cmbedition 
      Height          =   315
      Left            =   2175
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   3015
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
      Left            =   420
      TabIndex        =   11
      Top             =   4875
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
      Left            =   1740
      TabIndex        =   12
      Top             =   4875
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
      Left            =   3060
      TabIndex        =   13
      Top             =   4875
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
      Left            =   4380
      TabIndex        =   14
      Top             =   4875
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Info"
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
      Left            =   745
      TabIndex        =   19
      Top             =   3600
      Width           =   1395
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   745
      TabIndex        =   18
      Top             =   3009
      Width           =   1155
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   745
      TabIndex        =   17
      Top             =   1902
      Width           =   1155
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   745
      TabIndex        =   16
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   745
      TabIndex        =   15
      Top             =   2478
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Edition"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   745
      TabIndex        =   10
      Top             =   1311
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Card"
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
      Left            =   2355
      TabIndex        =   9
      Top             =   0
      Width           =   1290
   End
End
Attribute VB_Name = "frmratecard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim arrpubid() As Long      'Publication id
Dim arreditionid() As Long  'Edition id
Dim arrpositionid() As Long 'Position id

Private Sub cmbedition_Click()
cmbposition.Clear
txtrate = ""
txtinfo = ""
optunit(0).Value = True
If cmbedition = "" Then Exit Sub
If cmdnew.Caption = "&New" Then
addPosition arrpubid(cmbpublication.ListIndex), arreditionid(cmbedition.ListIndex)
Else
addPosition
End If
End Sub

Private Sub cmbposition_Click()
Dim rec As New ADODB.Recordset
txtrate = ""
txtinfo = ""
optunit(0).Value = True

If cmdnew.Caption = "&New" Then
    If cmbposition = "" Then Exit Sub
    rec.Open "select rate,unit,info from ratecard where publication=" & arrpubid(cmbpublication.ListIndex) & " and edition=" & arreditionid(cmbedition.ListIndex) & " and positions=" & arrpositionid(cmbposition.ListIndex), cn, adOpenStatic, adLockReadOnly
    txtrate = rec!Rate
    If rec!unit = "CC" Or rec!unit = "Page" Or rec!unit = "Panel" Then
        optunit(0).Value = True
        cmbunit = rec!unit
    Else
        optunit(1).Value = True
        txtunit = rec!unit
    End If
    txtinfo = rec!info
End If
End Sub

Private Sub cmbpublication_Click()
cmbedition.Clear
cmbposition.Clear
txtrate = ""
txtinfo = ""
optunit(0).Value = True

If cmbpublication = "" Then Exit Sub
If cmdnew.Caption = "&New" Then
addEdition arrpubid(cmbpublication.ListIndex)
Else
addEdition
End If
End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmddelete_Click()
Dim msg As Integer
If cmbposition = "" Then Exit Sub
msg = MsgBox("Do you want to delete this Rate Card Entry?", vbYesNo, "Confirm Delete")
If msg = vbYes Then
    cn.Execute "delete * from ratecard where publication=" & _
    arrpubid(cmbpublication.ListIndex) & " and edition=" & _
    arreditionid(cmbedition.ListIndex) & " and positions=" & _
    arrpositionid(cmbposition.ListIndex)
    MsgBox "Entry deleted from the Database.", vbInformation, "Deleted!!"
    addPosition arrpubid(cmbpublication.ListIndex), arreditionid(cmbedition.ListIndex)
End If
End Sub

Private Sub cmdnew_Click()
If cmdnew.Caption = "&New" Then 'new record
 cmddelete.Enabled = False
 cmdupdate.Enabled = False
  txtrate = ""
 txtinfo = ""
 optunit(0).Value = True
 cmbpublication.ListIndex = -1
 cmbedition.Clear
 cmbposition.Clear
 cmdnew.Caption = "&Add"

Else 'add record
 If IsValid = False Then Exit Sub
 On Error GoTo errorHandler
 cn.Execute "insert into ratecard values ('" & arrpubid(cmbpublication.ListIndex) & "','" & arreditionid(cmbedition.ListIndex) & "','" & arrpositionid(cmbposition.ListIndex) & "','" & txtrate & "','" & cmbunit + Trim(txtunit) & "','" & txtinfo & "')"
 MsgBox "Rate Card Entry Added into Database.", vbInformation
 addpublication
 cmddelete.Enabled = True
 cmdupdate.Enabled = True
 cmdnew.Caption = "&New"
End If
Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdupdate_Click()
If cmbposition = "" Then Exit Sub
If IsValid = False Then Exit Sub
 cn.Execute "update ratecard set rate='" & txtrate & "', unit='" & cmbunit + Trim(txtunit) & "', info='" & Trim(txtinfo) & "' where publication=" & arrpubid(cmbpublication.ListIndex) & " and edition=" & arreditionid(cmbedition.ListIndex) & " and positions=" & arrpositionid(cmbposition.ListIndex)
 MsgBox "Entry modified in Database.", vbInformation
End Sub

Private Sub Form_Load()
Me.Top = 50
Me.Left = 50
optunit(0).Value = True
cmbunit.ListIndex = 0
addpublication
End Sub


Function IsValid() As Boolean
 
 If cmbposition = "" Then
 IsValid = False
 Exit Function
 End If
 
 If IsNumeric(txtrate) = False Or Val(txtrate) < 0 Or Val(txtrate) > 10000 Then
 MsgBox "Please enter a valid Rate.", vbCritical, "Invalid Entry"
 txtrate.SetFocus
 IsValid = False
 Exit Function
 End If
  
 If Trim(txtunit) = "" And cmbunit = "" Then
 MsgBox "Please enter a Unit.", vbCritical, "Incomplete Entry"
 cmbunit.SetFocus
 IsValid = False
 Exit Function
 End If
  
IsValid = True

End Function

Private Sub optunit_Click(Index As Integer)
If optunit(0).Value = True Then
cmbunit.Enabled = True
cmbunit = "CC"
txtunit = ""
txtunit.Enabled = False
ElseIf optunit(1).Value = True Then
cmbunit.ListIndex = -1
cmbunit.Enabled = False
txtunit.Enabled = True
End If
End Sub
Sub addpublication()
Dim rec As New ADODB.Recordset
cmbpublication.Clear
Erase arrpubid
ReDim arrpubid(0)
'adding names in combo
rec.Open "select publication_id,name from publication", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
MsgBox "Please add a publication.", vbCritical, "No publication Found"
Exit Sub
End If


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
Sub addEdition(Optional pubid As Long = 0)
Dim rec As New ADODB.Recordset
cmbedition.Clear
Erase arreditionid
ReDim arreditionid(0)
'adding names in combo
If pubid = 0 Then
rec.Open "select edition_id,name from edition", cn, adOpenForwardOnly, adLockReadOnly
Else
rec.Open "select edition_id,name from ratecard,edition where edition=edition_id and publication=" & pubid, cn, adOpenForwardOnly, adLockReadOnly
End If

If rec.BOF = True And rec.EOF = True Then
'MsgBox "Please add an Edition.", vbCritical, "No Edition Found"
Exit Sub
End If


cmbedition.AddItem rec!Name
arreditionid(0) = rec!Edition_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arreditionid(UBound(arreditionid) + 1)
cmbedition.AddItem rec!Name
arreditionid(UBound(arreditionid)) = rec!Edition_id
rec.MoveNext
Wend

cmbedition.ListIndex = 0

End Sub
Sub addPosition(Optional pubid As Long = 0, Optional editionid As Long = 0)
Dim rec As New ADODB.Recordset
cmbposition.Clear
Erase arrpositionid
ReDim arrpositionid(0)
'adding names in combo
If pubid = 0 Then
'******* this does not work if there are no record in ratecard
'rec.Open "select position_id,name from positions where position_id<>any(select positions from ratecard where publication=" & arrpubid(cmbpublication.ListIndex) & " and edition=" & arreditionid(cmbedition.ListIndex) & ")", cn, adOpenForwardOnly, adLockReadOnly
'*************************************************************************************
rec.Open "select position_id,name from positions", cn, adOpenForwardOnly, adLockReadOnly
Else
rec.Open "select position_id,name from ratecard,positions where positions=position_id and publication=" & pubid & " and edition=" & editionid, cn, adOpenForwardOnly, adLockReadOnly
End If

If rec.BOF = True And rec.EOF = True Then Exit Sub

cmbposition.AddItem rec!Name
arrpositionid(0) = rec!position_id
rec.MoveNext
While rec.EOF = False
ReDim Preserve arrpositionid(UBound(arrpositionid) + 1)
cmbposition.AddItem rec!Name
arrpositionid(UBound(arrpositionid)) = rec!position_id
rec.MoveNext
Wend

cmbposition.ListIndex = 0

End Sub
