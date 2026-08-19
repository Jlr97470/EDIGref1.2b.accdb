Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description
End Sub

Private Sub CmdTransmettre_Click()

   On Error GoTo Err_CmdTransmettre_Click

   GenererFichierExportReglement ERD_Code

   DoCmd.Close acForm, Me.Name

Exit_CmdTransmettre_Click:

   Exit Sub

Err_CmdTransmettre_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdTransmettre_Click"

   Resume Exit_CmdTransmettre_Click
End Sub
