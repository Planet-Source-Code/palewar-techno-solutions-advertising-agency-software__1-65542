VERSION 5.00
Begin VB.MDIForm frmmdi 
   BackColor       =   &H8000000C&
   Caption         =   "MAZDA CREATIVE INC."
   ClientHeight    =   6180
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6975
   Icon            =   "frmmdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewro 
         Caption         =   "&Release Order"
      End
      Begin VB.Menu mnunewbill 
         Caption         =   "&Bill"
      End
      Begin VB.Menu mnufilenote 
         Caption         =   "Debit/Credit &Note"
      End
      Begin VB.Menu mnufilesep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnufilepubbills 
         Caption         =   "&Publication Bills"
      End
      Begin VB.Menu mnufilesep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnudata 
      Caption         =   "&Data"
      Begin VB.Menu mnudataratecard 
         Caption         =   "Rate &Card"
      End
      Begin VB.Menu mnudatasep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnudatapublication 
         Caption         =   "&Publication"
      End
      Begin VB.Menu mnudataedition 
         Caption         =   "&Edition"
      End
      Begin VB.Menu mnudataposition 
         Caption         =   "&P&osition"
      End
      Begin VB.Menu mnudatapremium 
         Caption         =   "Pre&mium"
      End
   End
   Begin VB.Menu mnureports 
      Caption         =   "&Reports"
      Begin VB.Menu mnureportsales 
         Caption         =   "&Sales Report"
      End
      Begin VB.Menu mnureleaseorderregister 
         Caption         =   "&Release Order Register"
      End
      Begin VB.Menu mnupublicationregister 
         Caption         =   "&Publication Register"
      End
   End
   Begin VB.Menu mnutools 
      Caption         =   "&Tools"
      Begin VB.Menu mnubackup 
         Caption         =   "&Backup Database"
      End
      Begin VB.Menu mnurestore 
         Caption         =   "&Restore Database"
         Visible         =   0   'False
      End
      Begin VB.Menu mnutoolssep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuremovedb 
         Caption         =   "&Remove DataLocation"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
      Begin VB.Menu mnutilehori 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnutileverti 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnucascade 
         Caption         =   "&Cascade"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpabout 
         Caption         =   "&About AdSoft"
      End
   End
End
Attribute VB_Name = "frmmdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3BDF9EA90258"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"MDI Form"
Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then Me.PopupMenu mnufile

End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim yesno As Integer
'yesno = MsgBox("Click on YES to take BackUp of Database, Click NO to exit.", vbYesNo, "Back Up Database!!")
'If yesno = vbYes Then backuP
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
On Error Resume Next
cn.Close
Set cn = Nothing
End Sub

Private Sub mnubackup_Click()
backuP
End Sub

Private Sub mnucascade_Click()
Me.Arrange vbCascade
End Sub

Private Sub mnudataedition_Click()
frmedition.Show
frmedition.ZOrder 0
End Sub

Private Sub mnudataposition_Click()
frmposition.Show
frmposition.ZOrder 0
End Sub

Private Sub mnudatapremium_Click()
frmpremium.Show
frmpremium.ZOrder 0
End Sub

Private Sub mnudatapublication_Click()
frmpublication.Show
frmpublication.ZOrder 0
End Sub

Private Sub mnudataratecard_Click()
frmratecard.Show
frmratecard.ZOrder 0
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnufilenote_Click()
frmnote.Show
frmnote.ZOrder 0
End Sub

Private Sub mnufilepubbills_Click()
frmpubbill.Show
frmpubbill.ZOrder 0
End Sub

Private Sub mnuhelpabout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnunewbill_Click()
frmbill.Show
frmbill.ZOrder 0
End Sub

Private Sub mnunewro_Click()
frmro.Show
frmro.ZOrder 0
End Sub

Private Sub mnupublicationregister_Click()
frmdatepicker.repname = "pubregister"
frmdatepicker.Show
frmdatepicker.ZOrder 0
End Sub

Private Sub mnureleaseorderregister_Click()
frmdatepicker.repname = "roregister"
frmdatepicker.Show
frmdatepicker.ZOrder 0
End Sub

Private Sub mnuremovedb_Click()
Dim msg As Integer
msg = MsgBox("Current Data Location: '" & datalocation & "'" & Chr(13) & "Do you want to remove this Setting?", vbYesNo, "Confirm Remove Settings")
If msg = vbNo Then Exit Sub
On Error Resume Next
DeleteSetting "AdSoft", "startup", "datalocation"
MsgBox "Data Location Settings removed." & vbCrLf & "Click on OK to quit Software.", vbInformation, "Done!!"
Unload Me

End Sub

Private Sub mnureportsales_Click()
frmdatepicker.repname = "sales"
frmdatepicker.Show
frmdatepicker.ZOrder 0
End Sub

Private Sub mnurestore_Click()
Set cn = Nothing
restore
End Sub

Private Sub mnutilehori_Click()
Me.Arrange vbTileHorizontal
End Sub

Private Sub mnutileverti_Click()
Me.Arrange vbTileVertical
End Sub
