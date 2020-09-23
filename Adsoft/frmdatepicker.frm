VERSION 5.00
Begin VB.Form frmdatepicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   3510
   Begin VB.TextBox txtdateto 
      Height          =   315
      Left            =   1605
      MaxLength       =   10
      TabIndex        =   1
      Top             =   1365
      Width           =   1380
   End
   Begin VB.TextBox txtdatefrom 
      Height          =   315
      Left            =   1605
      MaxLength       =   10
      TabIndex        =   0
      Top             =   660
      Width           =   1380
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
      Left            =   428
      TabIndex        =   2
      ToolTipText     =   "Print this RO"
      Top             =   2100
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
      Left            =   1868
      TabIndex        =   3
      ToolTipText     =   "Cancel Current RO"
      Top             =   2100
      Width           =   1215
   End
   Begin VB.Label lbldate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Ex. 26-03-1978)"
      Height          =   195
      Left            =   1590
      TabIndex        =   7
      Top             =   990
      Width           =   1170
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
      Left            =   525
      TabIndex        =   6
      Top             =   1365
      Width           =   735
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
      Left            =   525
      TabIndex        =   5
      Top             =   660
      Width           =   945
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
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   870
   End
End
Attribute VB_Name = "frmdatepicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public repname As String
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
On Error GoTo eror
Dim datefrom As Date, dateto As Date

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

ShowReport repname, datefrom, dateto
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
