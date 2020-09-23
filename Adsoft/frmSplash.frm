VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   3210
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   2865
      Left            =   120
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6885
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 3
End Sub
