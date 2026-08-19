Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrForm As String

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   StrForm = Me.OpenArgs

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmbGroupe_AfterUpdate()

   On Error GoTo Err_CmbGroupe_AfterUpdate

   CmbCli_Groupe.Requery

   CmbCli_Groupe = Null

Exit_CmbGroupe_AfterUpdate:

   Exit Sub

Err_CmbGroupe_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbGroupe_AfterUpdate:"

   Resume Exit_CmbGroupe_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub

Private Sub CmdOk_Click()
   Dim BolWhere As Boolean
   Dim StrCli_Code As String
   Dim StrCli_Nom As String
   Dim LngGroupe As Long
   Dim LngCli_Groupe As Long

   On Error GoTo Err_CmdOk_Click

   StrCli_Code = Nz(TxtCli_Code)

   StrCli_Nom = Nz(TxtCli_Nom)

   LngCli_Groupe = Nz(CmbCli_Groupe)

   LngGroupe = Nz(CmbGroupe)

   Forms(StrForm).SQLWhere = ""

   BolWhere = False

   Select Case StrCli_Code
      Case ""

      Case Else

         Select Case InStr(StrCli_Code, "*")
            Case Is > 0

            Case Else

              StrCli_Code = StrCli_Code & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Code] LIKE """ & StrCli_Code & """"

         BolWhere = True

   End Select

   Select Case StrCli_Nom
      Case ""

      Case Else

         Select Case InStr(StrCli_Nom, "*")
           Case Is > 0

           Case Else

              StrCli_Nom = Nz(StrCli_Nom) & "*"

         End Select

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Nom] LIKE """ & StrCli_Nom & """"

         BolWhere = True

   End Select

   Select Case LngGroupe
      Case 0

      Case Else

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Groupe] In (SELECT CG_Id FROM CLIENT_GROUPE WHERE CG_Parent_Id=" & LngGroupe & ";)"

         BolWhere = True

   End Select

   Select Case LngCli_Groupe
      Case 0

      Case Else

         Forms(StrForm).SQLWhere = Forms(StrForm).SQLWhere & IIf(BolWhere, " AND ", "") & " [Cli_Groupe]=" & LngCli_Groupe

         BolWhere = True

   End Select

   Select Case CadClass
      Case 0

        Forms(StrForm).SQLOrder = "Cli_Code"

      Case 1

        Forms(StrForm).SQLOrder = "Cli_Nom"

      Case 2

        Forms(StrForm).SQLOrder = "(SELECT CG_Parent_Id FROM CLIENT_GROUPE WHERE CG_Id=Cli_Groupe;)"

      Case 3

        Forms(StrForm).SQLOrder = "Cli_Groupe"

    End Select

    DoCmd.Close acForm, Me.Name

Exit_CmdOk_Click:

    Exit Sub

Err_CmdOk_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOk_Click
End Sub
