Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   Me.TxtTitre.Caption = "RECOUVREMENT DES CREANCES" & vbCrLf & "Version " & stVersion & " - " & stDateVersion

   Me.TxtTitreOmbre.Caption = "RECOUVREMENT DES CREANCES" & vbCrLf & "Version " & stVersion & " - " & stDateVersion

   DoEvents

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub
