Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ChkTIX_Flag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim StrSQL As String

   On Error GoTo Err_ChkTIX_Flag_MouseDown

   Select Case ChkTIX_Flag
      Case True

         StrSQL = "DELETE TMP_IMP_XLS.* FROM TMP_IMP_XLS WHERE [TIX_Ref Client]='" & [Ref Client] & "' AND [TIX_N°]='" & [N°] & "';"

      Case False

         StrSQL = "INSERT INTO TMP_IMP_XLS ([TIX_Ref Client], [TIX_N°]) VALUES('" & [Ref Client] & "','" & [N°] & "');"

   End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   ChkTIX_Flag.Requery

Exit_ChkTIX_Flag_MouseDown:

   Exit Sub

Err_ChkTIX_Flag_MouseDown:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkTIX_Flag_MouseDown"

   Resume Exit_ChkTIX_Flag_MouseDown
End Sub

Private Sub CmdOk_Click()

   On Error GoTo Err_CmdOk_Click

   Select Case ChkDbl
      Case True

         DoCmd.OpenForm "FrmFicheDebiteurListe", acNormal, , "Déb_Code_Client='" & Forms!FrmImportationXLS.Form.TxtIMPX_Cli_Code & "' AND Déb_Réf1_Client='" & [Ref Client] & "'", , acDialog

      Case False

   End Select

Err_CmdOk_Click:

   Exit Sub

Exit_CmdOk_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

   Resume Err_CmdOk_Click
End Sub
