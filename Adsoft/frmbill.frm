VERSION 5.00
Begin VB.Form frmbill 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BILL"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10590
   Begin VB.TextBox txttotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8880
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   47
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox cmbbillno 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
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
      Left            =   3723
      TabIndex        =   44
      ToolTipText     =   "Print this Bill"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
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
      Left            =   7583
      TabIndex        =   43
      ToolTipText     =   "Close This Screen"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
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
      Left            =   1793
      TabIndex        =   42
      ToolTipText     =   "Make and save new Bill"
      Top             =   6000
      Width           =   1215
   End
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
      Left            =   5653
      TabIndex        =   41
      ToolTipText     =   "Cancel Current Bill"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtunit 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   19
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtunit 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   14
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtunit 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtparticulars 
      Height          =   405
      Index           =   0
      Left            =   720
      MaxLength       =   255
      TabIndex        =   7
      Top             =   3840
      Width           =   4935
   End
   Begin VB.TextBox txtparticulars 
      Height          =   375
      Index           =   2
      Left            =   720
      MaxLength       =   255
      TabIndex        =   17
      Top             =   4800
      Width           =   4935
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   18
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox txtrate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   20
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   2
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   21
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtparticulars 
      Height          =   375
      Index           =   1
      Left            =   720
      MaxLength       =   255
      TabIndex        =   12
      Top             =   4320
      Width           =   4935
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   13
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtrate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   15
      Top             =   4320
      Width           =   1095
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   1
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   16
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox txtamount 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   8880
      MaxLength       =   15
      TabIndex        =   11
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtrate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   10
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtqty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   5760
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   1755
      TabIndex        =   26
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtrodate 
         Height          =   375
         Left            =   4920
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optro 
         Caption         =   "R.O. NO."
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optdm 
         Caption         =   "DM  NO."
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtdmdate 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4920
         MaxLength       =   10
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbrono 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtdmno 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         MaxLength       =   12
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.Image imgtip 
         Height          =   420
         Left            =   3225
         Picture         =   "frmbill.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "Click For More Info"
         Top             =   315
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   30
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3840
         TabIndex        =   29
         Top             =   840
         Width           =   675
      End
   End
   Begin VB.TextBox txtclient 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox txtdate 
      Height          =   375
      Left            =   7920
      MaxLength       =   10
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
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
      Left            =   8040
      TabIndex        =   46
      Top             =   5280
      Width           =   555
   End
   Begin VB.Label lblcancel 
      Caption         =   "C A N C E L L E D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   45
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      TabIndex        =   40
      Top             =   3360
      Width           =   345
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BILL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   360
      TabIndex        =   39
      Top             =   600
      Width           =   960
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
      Left            =   480
      TabIndex        =   38
      Top             =   3960
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
      Left            =   480
      TabIndex        =   37
      Top             =   4920
      Width           =   135
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
      Left            =   480
      TabIndex        =   36
      Top             =   4440
      Width           =   135
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
      Left            =   9120
      TabIndex        =   35
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
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
      Left            =   7920
      TabIndex        =   34
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size/Qty"
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
      Left            =   5760
      TabIndex        =   33
      Top             =   3360
      Width           =   975
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
      Left            =   2640
      TabIndex        =   32
      Top             =   3360
      Width           =   945
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
      Left            =   240
      TabIndex        =   31
      Top             =   3360
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
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
      Left            =   600
      TabIndex        =   25
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6840
      TabIndex        =   24
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Ex. 26-03-1978)"
      Height          =   195
      Left            =   7950
      TabIndex        =   23
      Top             =   2340
      Width           =   1170
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   600
      TabIndex        =   22
      Top             =   2040
      Width           =   975
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3BDF9EA601F9"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Form"
'to convert number to words in bill report it requires
'crxlat32.dll , include it in setup or bill wont be printed
Option Base 1
'************************Anish***********
Dim arrRoId() As Long   'contains the id of rono

Private Sub cmbbillno_Click()
On Error GoTo errorHandler
Dim rec As New ADODB.Recordset
Dim recRefDate As New ADODB.Recordset   'getting refdate
ClearControls
If cmbbillno = "" Then Exit Sub
rec.Open "select * from billmaster where billno=" & Split(cmbbillno, "/")(0) & " and finyear=" & Val("20" & Split(cmbbillno, "/")(1)), cn, adOpenStatic, adLockReadOnly
txtdate = Format(rec("billdate"), "dd-mm-yyyy")
txtclient = rec("client")

If rec("ref") = "r" Then
optro.Value = True
'**********************************Anish******************************
'checking the romaster_id in array and then assign (the position of that in array) in combo.listindex
For i = 1 To UBound(arrRoId) - 1
    If arrRoId(i) = rec("romaster_id") Then
        cmbrono.ListIndex = i - 1
        Exit For
    End If
Next

recRefDate.Open "select rodate from romaster where romaster_id=" & rec("romaster_id") & "", cn, adOpenForwardOnly, adLockReadOnly
txtrodate = recRefDate("rodate")
'txtrodate = rec("refdate")
'*********************************************************************
End If
If rec("ref") = "d" Then
optdm.Value = True
'**************************************Anish***********************************************
Dim recDmValues As New ADODB.Recordset
'getting DM values from DmMaster table
recDmValues.Open "select * from DmMaster where dmId=" & rec("RoMaster_Id") & "", cn, adOpenForwardOnly, adLockReadOnly

txtdmno = recDmValues("DmNo")
txtdmdate = recDmValues("DmDate")
'******************************************************************************************
End If

lblcancel.Visible = rec("cancelled")
cmdcancel.Enabled = Not (rec("cancelled"))
cmdprint.Enabled = Not (rec("cancelled"))


'***********************************Anish**************************************************
Dim recDetail As New ADODB.Recordset
'getting billdetails according to billmasterID
recDetail.Open "select * from billdetail where billmaster_ID=" & rec("billmaster_Id") & "", cn, adOpenForwardOnly, adLockReadOnly

i = 0
While recDetail.EOF = False
    txtparticulars(i) = recDetail("particulars")
    txtqty(i) = recDetail("qty")
    txtrate(i) = recDetail("rate")
    txtunit(i) = recDetail("unit")
    txtamount(i) = recDetail("amount")
    i = i + 1
    recDetail.MoveNext
Wend

'******************************************************************************************
'''rec.Close
'''rec.Open "select * from billdetail where billno=" & Split(cmbbillno, "/")(0) & " and finyear=" & Val("20" & Split(cmbbillno, "/")(1)), cn, adOpenStatic, adLockReadOnly
'''i = 0
'''While rec.EOF = False
'''txtparticulars(i) = rec("particulars")
'''txtqty(i) = rec("qty")
'''txtrate(i) = rec("rate")
'''txtunit(i) = rec("unit")
'''txtamount(i) = rec("amount")
'''i = i + 1
'''rec.MoveNext
'''Wend

Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical
End Sub

Private Sub cmbrono_Click()
Dim rec As New ADODB.Recordset


If cmbrono = "" Or cmdsave.Caption = "&New" Then Exit Sub

ClearControls


On Error GoTo errorHandler
var_rono = Split(cmbrono, "/")(0)
var_finyear = Val("20" & Trim(Split(cmbrono, "/")(1)))

'**************************************Anish************************************************
Dim recRoDet As New ADODB.Recordset
'getting values from romaster....with id
rec.Open "select romaster_Id, rodate,client from romaster where rono=" & var_rono & " and finyear=" & var_finyear, cn, adOpenStatic, adLockReadOnly
txtrodate = Format(rec("rodate"), "dd-mm-yyyy")
txtclient = rec("client")
'getting values from rodetail according to assign romaster_id
recRoDet.Open "select * from rodetail where romaster_Id=" & rec("romaster_Id") & "", cn, adOpenStatic, adLockReadOnly

'*******************************************************************************************

''''rec.Open "select rodate,client from romaster where rono=" & var_rono & " and finyear=" & var_finyear, cn, adOpenStatic, adLockReadOnly
''''txtrodate = Format(rec("rodate"), "dd-mm-yyyy")
''''txtclient = rec("client")
''''rec.Close
''''rec.Open "select * from rodetail where rono=" & var_rono & " and finyear=" & var_finyear, cn, adOpenStatic, adLockReadOnly


'************************
'"AD CAPTION : " & rec("caption") & " PUBLICATION : " & rec("publication") & " " & rec("edition") & " (" & rec("position") & ")" _
    & " RELEASE DATE : " & rec("releasedate")
'************************
rec.Close
i = 0
While recRoDet.EOF = False
    txtparticulars(i) = recRoDet("caption") & " Ad. in " & recRoDet("publication") & " " & recRoDet("edition") & " (" & recRoDet("position") & ") Dated " & Format(recRoDet("releasedate"), "dd-mm-yyyy")
    If recRoDet("premium_percent") <> 0 Then
        txtrate(i) = recRoDet("rate") + recRoDet("rate") * (recRoDet!premium_percent / 100)
    Else
        txtrate(i) = recRoDet("rate")
    End If
    
    If recRoDet("size2") = 0 Then
        txtqty(i) = recRoDet("size1")
    Else
        txtqty(i) = recRoDet("size1") * recRoDet("size2")
    End If
    recRoDet.MoveNext
    i = i + 1
Wend

 Exit Sub
errorHandler:

End Sub

Private Sub cmdcancel_Click()
Dim ask As Integer

    ask = MsgBox("Do you really want to Cancel Bill No. " & cmbbillno & "?", vbYesNo, "Confirm Cancellation")
    If ask = vbYes Then
        cn.Execute "update billmaster set cancelled=true where billno=" & Split(cmbbillno, "/")(0) & " and finyear=" & Val("20" & Split(cmbbillno, "/")(1))
        MsgBox "Bill No. " & cmbbillno & " has been Cancelled, which makes it unusable.", vbInformation, "Bill Cancelled !!"
        cmbbillno = cmbbillno 'causes click event
    End If


End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
On Error GoTo errorHandler
Dim rec As New ADODB.Recordset
Dim Rep As New repbill 'creating instance of crystal reports dsr
Dim frmrep As New frmreports
Dim optValue As String

repbillno = Split(cmbbillno, "/")(0)
repfinyear = "20" & Split(cmbbillno, "/")(1)

'*****************************************Anish************************************
'if rono option is true then assign 'r' otherwise assign 'd'
If optro.Value = True Then
   optValue = "r"
ElseIf optdm.Value = True Then
   optValue = "d"
End If
'displaying bill report
Dim str As String
str = "SELECT  billmaster.*, billdetail.*, " _
    & "(SELECT DMNO FROM DMMASTER WHERE DMID=BILLMASTER.ROMASTER_ID ) AS DMNO,(SELECT DMDATE FROM DMMASTER WHERE DMID=BILLMASTER.ROMASTER_ID) AS DMDATE,  " _
    & "(SELECT (RONO & '/' & right(FINYEAR,2) & '/' & right(FINYEAR+1,2)) FROM ROMASTER WHERE ROMASTER_ID=BILLMASTER.ROMASTER_ID) AS RONO,(select rodate  FROM ROMASTER WHERE ROMASTER_ID=BILLMASTER.ROMASTER_ID) as  rodate " _
    & "From BILLMASTER, billdetail Where " _
    & "BILLMASTER.billmaster_id = [billdetail].[billmaster_id] " _
    & "AND billmaster.billno=" & repbillno & " and billmaster.finyear=" & repfinyear & " and billmaster.ref='" & optValue & "'"

 rec.Open str, cn, adOpenStatic, adLockReadOnly

''rec.Open "SELECT billmaster.*, billdetail.*, romaster.rono " _
''        & "From billmaster, billdetail, romaster " _
''        & "Where billmaster.billmaster_id = billdetail.billmaster_id And " _
''        & "billmaster.romaster_id = romaster.romaster_id " _
''        & "billmaster.billno=" & repbillno & " and billmaster.finyear=" & repfinyear, cn, adOpenStatic, adLockReadOnly
'**********************************************************************************

'rec.Open "SELECT billmaster.*, billdetail.* FROM billDetail RIGHT JOIN billMaster ON (billDetail.finyear = billMaster.finyear) AND (billDetail.billno = billMaster.billno) where billmaster.billno=" & repbillno & " and billmaster.finyear=" & repfinyear, cn, adOpenStatic, adLockReadOnly

Rep.Database.SetDataSource rec  'assinging recordset to instance of cr dsr
'assigning servicetax and commission in report
For i = 1 To Rep.FormulaFields.Count
If Rep.FormulaFields(i).Name = "{@ServiceTaxValue}" Then Rep.FormulaFields(i).Text = ServiceTax
Next

For i = 1 To Rep.FormulaFields.Count
If Rep.FormulaFields(i).Name = "{@Commission}" Then Rep.FormulaFields(i).Text = Commission
Next




frmrep.CRViewer1.ReportSource = Rep
frmrep.CRViewer1.PrintReport
frmrep.Caption = "Bill"
Exit Sub
errorHandler:
MsgBox Err.Description, vbCritical



'With CrystalReport1
'.ReportFileName = App.Path + "\bill.rpt"
'.DataFiles(0) = App.Path + "\mazda.mdb"
'.SelectionFormula = "{billMaster.billno}=" & Split(cmbbillno, "/")(0) & " and {billMaster.finyear}=" & "20" & Split(cmbbillno, "/")(1)
'.PrintReport
'End With

End Sub

Private Sub cmdsave_Click()
 Dim billdate As Date, rec_billno As New ADODB.Recordset
 
 If cmdsave.Caption = "&Save" Then
 
 If IsValid = False Then Exit Sub
 
 billdate = Split(txtdate, "-")(1) & "/" & Split(txtdate, "-")(0) & "/" & Split(txtdate, "-")(2)
 rec_billno.Open "select max(billno) as maxno from billmaster where finyear=" & finyear(billdate), cn, adOpenStatic, adLockReadOnly
 
 If (rec_billno.BOF And rec_billno.EOF) Or IsNull(rec_billno("maxno")) Then
    newbill = 1
 Else
    newbill = rec_billno("maxno") + 1
 End If
 
 On Error GoTo errorHandler
 If optro.Value = True Then
    ref = "r"
    refno = cmbrono
'''    refdate = Split(txtrodate, "-")(1) & "/" & Split(txtrodate, "-")(0) & "/" & Split(txtrodate, "-")(2)
 End If
'*************************************Anish*************************************************
 Dim recDmId As New ADODB.Recordset
 If optdm.Value = True Then
    ref = "d"
    'saving DM values in "DMTABLE"
    cn.Execute "insert into DmMaster (DmNo, DmDate) values ('" & txtdmno & "',#" & txtdmdate & "#)"
    refno = Trim(txtdmno)
'''    refdate = Split(txtdmdate, "-")(1) & "/" & Split(txtdmdate, "-")(0) & "/" & Split(txtdmdate, "-")(2)
 End If
 
'saving in master table
'on 17-05-03 we remove refdate from query
If optro.Value = True Then
    'inserting the values for RO in DmMaster table
    
    cn.Execute "insert into BillMaster (billno, finyear, billdate, ref, RoMaster_ID, client, cancelled) values ('" & newbill & "','" & finyear(billdate) & "','" & billdate & "','" & ref & "'," & arrRoId(cmbrono.ListIndex + 1) & ",'" & Trim(txtclient) & "',false)"

ElseIf optdm.Value = True Then
    'retrieving maximum DMID from DmMaster table
    recDmId.Open "select max(DmID) as maxDmId from DmMaster", cn, adOpenForwardOnly, adLockReadOnly
    
    'now inserting the values for Dm in billmater table
    cn.Execute "insert into BillMaster (billno, finyear, billdate, ref, RoMaster_ID, client, cancelled) values ('" & newbill & "','" & finyear(billdate) & "','" & billdate & "','" & ref & "'," & recDmId("MaxDmId") & ",'" & Trim(txtclient) & "',false)"

End If

'getting the Billmaster_Id for entering values into BillDetail table
Dim recBillID As New ADODB.Recordset
recBillID.Open "select BillMaster_Id from BillMaster where billno=" & newbill & " and finyear=" & finyear(billdate) & "", cn, adOpenForwardOnly, adLockReadOnly
 
'saving in detail table
 For i = 0 To 2
     If Trim(txtparticulars(i)) <> "" Then
        cn.Execute "insert into BillDetail (billmaster_Id, particulars, qty, unit, rate, amount) values (" & recBillID("BillMaster_ID") & ",'" & Trim(txtparticulars(i)) & "','" & Val(txtqty(i)) & "','" & Trim(txtunit(i)) & "','" & Val(txtrate(i)) & "','" & txtamount(i) & "')"
     End If
 Next
 
'******************************************************************************************

'''' cn.Execute "insert into billmaster values ('" & newbill & "','" & finyear(billdate) & "','" & billdate & "','" & ref & "','" & refno & "','" & CDate(refdate) & "','" & Trim(txtclient) & "',false)"
'''' 'saving in detail table
''''
'''' For i = 0 To 2
''''     If Trim(txtparticulars(i)) <> "" Then
''''        cn.Execute "insert into billdetail values ('" & newbill & "','" & finyear(billdate) & "','" & Trim(txtparticulars(i)) & "','" & Val(txtqty(i)) & "','" & Trim(txtunit(i)) & "','" & Val(txtrate(i)) & "','" & txtamount(i) & "')"
''''     End If
'''' Next
 
 MsgBox "Bill Saved into Database.", vbInformation
 
 addbill
 cmbbillno.ListIndex = cmbbillno.ListCount - 1
 cmbbillno.Enabled = True
 
 cmdprint.Enabled = True
 cmdcancel.Enabled = True
 cmdsave.Caption = "&New"
  
Exit Sub

errorHandler:

MsgBox "Error No. " & Err.Number & vbCrLf & Err.Description, vbCritical

'**************************************Anish************************************************
'deleting records in case of error
cn.Execute "delete * from billmaster where billno=" & newbill & " and finyear=" & finyear(billdate)
cn.Execute "delete * from billdetail where BillMaster_ID=" & recBillID("BillMaster_ID")

'*******************************************************************************************
''''cn.Execute "delete * from billmaster where billno=" & newbill
''''cn.Execute "delete * from billdetail where billno=" & newbill

Else

    cmbrono.ListIndex = -1
    cmbbillno.ListIndex = -1
    cmbbillno.Enabled = False
    ClearControls
    txtdate = Format(Date, "dd-mm-yyyy")
    cmdprint.Enabled = False
    cmdcancel.Enabled = False
    cmdsave.Caption = "&Save"


End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
lbldate = "(Ex. " & Format(Date, "dd-mm-yyyy") & ")"
addro 'Adds RO No in combo box
addbill
If cmbbillno.ListCount > 0 Then
    cmbbillno.ListIndex = 0
Else
    cmdsave.Value = True
End If
End Sub

Private Sub imgtip_Click()
Dim rec As New ADODB.Recordset, billedamt As Double, roamt As Double
Dim bills As String
Dim temprec As New ADODB.Recordset

bills = "Bill No." + Space(25) + "Bill Amount" + vbCrLf + String(40, "-")
On Error GoTo errorHandler
rec.Open "select billno,finyear from billmaster where ref='r' and cancelled=false and romaster_id=" & arrRoId(cmbrono.ListIndex + 1) & "", cn, adOpenStatic, adLockReadOnly
'rec.Open "select billno,finyear from billmaster where ref='r' and cancelled=false and refno='" & cmbrono & "'", cn, adOpenStatic, adLockReadOnly
i = 1
While rec.EOF = False
    temprec.Open "select sum(amount) as tempamt from billdetail where " _
    & "billmaster_id=(select billmaster_id from billmaster where " _
    & "billno=" & rec("billno") & " and finyear=" & rec("finyear") & ")", cn, adOpenStatic, adLockReadOnly
'    temprec.Open "select sum(amount) as tempamt from billdetail where billno=" & rec("billno") & " and finyear=" & rec("finyear"), cn, adOpenStatic, adLockReadOnly
    bills = bills & vbCrLf & rec("billno") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2) + Space(26) & temprec("tempamt")
    billedamt = billedamt + temprec("tempamt")
    i = i + 1
    rec.MoveNext
    temprec.Close
Wend

rec.Close
rec.Open "select rate,size1,size2 from rodetail where romaster_id=" & arrRoId(cmbrono.ListIndex + 1) & "", cn, adOpenStatic, adLockReadOnly
'rec.Open "select rate,size1,size2 from rodetail where rono=" & Split(cmbrono, "/")(0) & " and finyear=" & "20" + Split(cmbrono, "/")(1), cn, adOpenStatic, adLockReadOnly
While rec.EOF = False
    If rec("size2") = 0 Then
        roamt = roamt + rec("rate") * rec("size1")
    Else
        roamt = roamt + rec("rate") * rec("size1") * rec("size2")
    End If
    rec.MoveNext
Wend
rec.Close

bills = bills + vbCrLf + String(40, "-") + vbCrLf + "Total Billed Amount" + Space(7) + str(billedamt) + vbCrLf
bills = bills + "Ro Amount" + Space(20) + str(roamt) + vbCrLf
bills = bills + "Unbilled Amount" + Space(12) + str(roamt - billedamt) + vbCrLf + String(40, "-")
MsgBox bills, vbOKOnly, "RO No. " & cmbrono
Exit Sub

errorHandler:
MsgBox "Operation Failed!!"
End Sub

Private Sub optdm_Click()
cmbrono.Enabled = False
txtrodate.Enabled = False
imgtip.Visible = False
cmbrono.ListIndex = -1
txtrodate = ""
txtdmno.Enabled = True
txtdmdate.Enabled = True
End Sub

Private Sub optro_Click()
cmbrono.Enabled = True
txtrodate.Enabled = True
imgtip.Visible = True
txtdmno = ""
txtdmdate = ""
txtdmno.Enabled = False
txtdmdate.Enabled = False
End Sub
Sub addro()
Dim rec As New ADODB.Recordset
ReDim arrRoId(1)
cmbrono.Clear
'****************************************Anish****************************************
'fetching rono records from romaster
rec.Open "select romaster_id, rono,finyear from romaster where cancelled=false order by finyear,rono ", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If
'adding rono in combo
While rec.EOF = False
    arrRoId(UBound(arrRoId)) = rec("romaster_Id")
    temprono = rec("rono") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2)
    cmbrono.AddItem temprono
    temprono = ""
    ReDim Preserve arrRoId(UBound(arrRoId) + 1)
    rec.MoveNext
Wend
'*************************************************************************************
End Sub

Private Sub txtamount_Change(Index As Integer)
txttotal = Val(txtamount(0)) + Val(txtamount(1)) + Val(txtamount(2))
End Sub

Private Sub txtdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtdate) = 2 Or Len(txtdate) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtdate) = 2 Or Len(txtdate) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If


End Sub

Private Sub txtdmdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii = 45 And (Len(txtdmdate) = 2 Or Len(txtdmdate) = 5) Then Exit Sub ' allows '-'
If KeyAscii <> 45 And (Len(txtdmdate) = 2 Or Len(txtdmdate) = 5) Then
KeyAscii = 0 ' do not allow anything except - in 3rd and 6th place
Beep
Exit Sub
End If

If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtqty_Change(Index As Integer)
txtamount(Index) = CSng(Val(txtqty(Index))) * CSng(Val(txtrate(Index)))
End Sub
Function IsValid() As Boolean
 On Error GoTo errorHandler
 
 If optro.Value = True Then
    If cmbrono = "" Then
    IsValid = False
    MsgBox "Please select a RO No.", vbCritical, "Incomplete Entry"
    cmbrono.SetFocus
    Exit Function
    End If
 End If
 
If optdm.Value = True Then
    If Trim(txtdmno) = "" Then
    IsValid = False
    MsgBox "Please enter a DM No.", vbCritical, "Incomplete Entry"
    txtdmno.SetFocus
    Exit Function
    End If
    
    If IsDate(Split(txtdmdate, "-")(1) & "-" & Split(txtdmdate, "-")(0) & "-" & Split(txtdmdate, "-")(2)) = False Then
    IsValid = False
    MsgBox "Please enter a valid DM Date.", vbCritical, "Invalid Entry"
    txtdmdate.SetFocus
    Exit Function
    End If
End If
 
If IsDate(Split(txtdate, "-")(1) & "-" & Split(txtdate, "-")(0) & "-" & Split(txtdate, "-")(2)) = False Then
IsValid = False
MsgBox "Please enter a valid Date.", vbCritical, "Invalid Entry"
txtdate.SetFocus
Exit Function
End If

 
 If Trim(txtclient) = "" Then
 MsgBox "Please enter a Client Name.", vbCritical, "Incomplete Entry"
 txtclient.SetFocus
 IsValid = False
 Exit Function
 End If
  
 If Trim(txtparticulars(0)) = "" Then
 MsgBox "Please enter Particulars.", vbCritical, "Incomplete Entry"
 txtparticulars(0).SetFocus
 IsValid = False
 Exit Function
 End If

 For i = 0 To 2
    If Trim(txtparticulars(i)) <> "" Then
        If Trim(txtrate(i)) <> "" And (IsNumeric(txtrate(i)) = False Or Val(txtrate(i)) < 0 Or Val(txtrate(i)) > 10000) Then
            MsgBox "Please enter a valid Rate.", vbCritical, "Invalid Entry"
            txtrate(i).SetFocus
            IsValid = False
            Exit Function
        End If
        If Trim(txtqty(i)) <> "" And (IsNumeric(txtqty(i)) = False Or Val(txtqty(i)) < 0 Or Val(txtqty(i)) > 10000) Then
            MsgBox "Please enter a valid Qty/Size.", vbCritical, "Invalid Entry"
            txtqty(i).SetFocus
            IsValid = False
            Exit Function
        End If
        If IsNumeric(txtamount(i)) = False Or Val(txtamount(i)) < 0 Then
            MsgBox "Please enter a valid Amount.", vbCritical, "Invalid Entry"
            txtamount(i).SetFocus
            IsValid = False
            Exit Function
        End If
    End If
Next
   
IsValid = True

Exit Function

errorHandler:
IsValid = False
MsgBox "Invalid Entry!!", vbCritical

End Function

Private Sub txtqty_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub 'allows backspace
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtrate_Change(Index As Integer)
txtamount(Index) = CSng(Val(txtqty(Index))) * CSng(Val(txtrate(Index)))
End Sub

Private Sub txtrate_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Or KeyAscii = 46 Then Exit Sub 'allows backspace and .
If KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
Beep
End If

End Sub

Sub addbill()
Dim rec As New ADODB.Recordset

cmbbillno.Clear
'fetching bill records from billmaster
rec.Open "select billno,finyear from billmaster order by finyear,billno ", cn, adOpenForwardOnly, adLockReadOnly

If rec.BOF = True And rec.EOF = True Then
    Exit Sub
End If
'adding bill in combo

While rec.EOF = False
    tempbillno = rec("billno") & "/" & Right(rec("finyear"), 2) & "/" & Right(rec("finyear") + 1, 2)
    cmbbillno.AddItem tempbillno
    tempbillno = ""
    rec.MoveNext
Wend
End Sub

Sub ClearControls()
txtclient = ""
txtrodate = ""
txtdmdate = ""
txtdmno = ""
lblcancel.Visible = False
For i = 0 To 2
    txtparticulars(i) = ""
    txtqty(i) = ""
    txtunit(i) = ""
    txtrate(i) = ""
    txtamount(i) = ""
Next
End Sub
