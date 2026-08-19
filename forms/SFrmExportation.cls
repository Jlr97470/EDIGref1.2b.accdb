Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub ChkCli_Flag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim StrSQL As String

   On Error GoTo Err_ChkCli_Flag_MouseDown

   Select Case ChkCli_Flag
      Case True

         StrSQL = "DELETE * FROM TMP_CLIENT WHERE TMP_CLIENT.TCli_Code='" & Cli_Code & "';"

      Case False

         StrSQL = "INSERT INTO TMP_CLIENT ( TCli_Code ) VALUES('" & Cli_Code & "');"

   End Select

   CurrentDb.Execute StrSQL, dbFailOnError

   ChkCli_Flag.Requery

Exit_ChkCli_Flag_MouseDown:

   Exit Sub

Err_ChkCli_Flag_MouseDown:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkCli_Flag_MouseDown"

   Resume Exit_ChkCli_Flag_MouseDown
End Sub
