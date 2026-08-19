Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub InitialiseListeChamp()
   Dim FldChamp As DAO.Field

   On Error Resume Next

   For Each FldChamp In CurrentDb.QueryDefs("SelDebiteurAdresse").Fields

      CmbTableChamp.AddItem FldChamp.Name & ";" & FldChamp.Properties("Description")

   Next

Exit_InitialiseListeChamp:

   Exit Sub

Err_InitialiseListeChamp:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : InitialiseListeChamp"

   Resume Exit_InitialiseListeChamp
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   InitialiseListeChamp

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmbFichierChamp_AfterUpdate()

   On Error GoTo Err_CmbFichierChamp_AfterUpdate

   CmdChampControleAjouter.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampControleSupprimer.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampAjouter.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampSupprimer.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

Exit_CmbFichierChamp_AfterUpdate:

   Exit Sub

Err_CmbFichierChamp_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbFichierChamp_AfterUpdate"

   Resume Exit_CmbFichierChamp_AfterUpdate
End Sub

Private Sub CmbTableChamp_AfterUpdate()

   On Error GoTo Err_CmbTableChamp_AfterUpdate

   CmdChampControleAjouter.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampControleSupprimer.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampAjouter.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

   CmdChampSupprimer.Enabled = (Not IsNull(CmbFichierChamp) And Not IsNull(CmbTableChamp))

Exit_CmbTableChamp_AfterUpdate:

   Exit Sub

Err_CmbTableChamp_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbTableChamp_AfterUpdate"

   Resume Exit_CmbTableChamp_AfterUpdate
End Sub

Private Sub CmdChampControleAjouter_Click()
   Dim IntNumero As Integer

   On Error GoTo Err_CmdChampControleAjouter_Click

   For IntNumero = 0 To LstChampControle.ListCount

      Select Case LstChampControle.Column(0, IntNumero)
         Case CmbFichierChamp

            Select Case LstChampControle.Column(2, IntNumero)
               Case CmbTableChamp

               Case Else

                  LstChampControle.RemoveItem IntNumero

                  LstChampControle.AddItem CmbFichierChamp & ";=>;" & CmbTableChamp

            End Select

            GoTo Exit_CmdChampControleAjouter_Click

         Case Else

      End Select

   Next

   LstChampControle.AddItem CmbFichierChamp & ";=>;" & CmbTableChamp

   CmdImporter.Enabled = (LstChampControle.ListCount > 0) And (LstChamp.ListCount > 0)

Exit_CmdChampControleAjouter_Click:

   Exit Sub

Err_CmdChampControleAjouter_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdChampControleAjouter_Click"

   Resume Exit_CmdChampControleAjouter_Click
End Sub

Private Sub CmdChampControleSupprimer_Click()
   Dim VarItem As Variant

   On Error GoTo Err_CmdChampControleSupprimer_Click

   For Each VarItem In LstChampControle.ItemsSelected

      LstChampControle.RemoveItem VarItem

   Next

   CmdImporter.Enabled = (LstChampControle.ListCount > 0) And (LstChamp.ListCount > 0)

Exit_CmdChampControleSupprimer_Click:

   Exit Sub

Err_CmdChampControleSupprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdChampControleSupprimer_Click"

   Resume Exit_CmdChampControleSupprimer_Click
End Sub

Private Sub CmdChampAjouter_Click()
   Dim IntNumero As Integer

   On Error GoTo Err_CmdChampAjouter_Click

   For IntNumero = 0 To LstChamp.ListCount

      Select Case LstChamp.Column(0, IntNumero)
         Case CmbFichierChamp

            Select Case LstChamp.Column(2, IntNumero)
               Case CmbTableChamp

               Case Else

                  LstChamp.RemoveItem IntNumero

                  LstChamp.AddItem CmbFichierChamp & ";=>;" & CmbTableChamp

            End Select

            GoTo Exit_CmdChampAjouter_Click

         Case Else

      End Select

   Next

   LstChamp.AddItem CmbFichierChamp & ";=>;" & CmbTableChamp

   CmdImporter.Enabled = (LstChampControle.ListCount > 0) And (LstChamp.ListCount > 0)

Exit_CmdChampAjouter_Click:

   Exit Sub

Err_CmdChampAjouter_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdChampAjouter_Click"

   Resume Exit_CmdChampAjouter_Click
End Sub

Private Sub CmdChampSupprimer_Click()
   Dim VarItem As Variant

   On Error GoTo Err_CmdChampSupprimer_Click

   For Each VarItem In LstChamp.ItemsSelected

      LstChamp.RemoveItem VarItem

   Next

   CmdImporter.Enabled = (LstChampControle.ListCount > 0) And (LstChamp.ListCount > 0)

Exit_CmdChampSupprimer_Click:

   Exit Sub

Err_CmdChampSupprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdChampSupprimer_Click"

   Resume Exit_CmdChampSupprimer_Click
End Sub

Private Sub CmdFermer_Click()
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_CmdFermer_Click

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "XLS_IMP"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description
End Sub

Private Sub CmdImprimer_Click()
   Dim StrSQL As String
   Dim StrSQLDebiteur As String
   Dim StrSQLImport As String
   Dim StrSQLJoin As String
   Dim IntRaportEtat As Integer
   Dim LngNumero As Long
   Dim LngChampNumero As Long

   On Error GoTo Err_CmdImprimer_Click

   StrSQL = "SELECT [Déb_Code], "

   For LngNumero = 0 To LstChampControle.ListCount - 1

      StrSQLImport = StrSQLImport & "[" & LstChampControle.Column(0, LngNumero) & "] As XLS_IMPChamp" & (LngNumero + 1) & ", "

      StrSQLDebiteur = StrSQLDebiteur & "[" & LstChampControle.Column(2, LngNumero) & "] As DebiteurAdresseChamp" & (LngNumero + 1) & ", "

      StrSQLJoin = StrSQLJoin & "[XLS_IMP].[" & LstChampControle.Column(0, LngNumero) & "]=[SelDebiteurAdresse].[" & LstChampControle.Column(2, LngNumero) & "] AND "

   Next

   LngChampNumero = LngNumero + 1

   For LngNumero = 0 To LstChamp.ListCount - 1

      StrSQLImport = StrSQLImport & "[" & LstChamp.Column(0, LngNumero) & "] As XLS_IMPChamp" & (LngChampNumero + LngNumero) & ", "

      StrSQLDebiteur = StrSQLDebiteur & "[" & LstChamp.Column(2, LngNumero) & "] As DebiteurAdresseChamp" & (LngChampNumero + LngNumero) & ", "

   Next

   StrSQL = StrSQL & Left(StrSQLDebiteur, Len(StrSQLDebiteur) - 2) & "," & Left(StrSQLImport, Len(StrSQLImport) - 2) & " FROM XLS_IMP INNER JOIN SelDebiteurAdresse ON "

   StrSQL = StrSQL & Left(StrSQLJoin, Len(StrSQLJoin) - 5) & ";"

   DoCmd.OpenReport "EtaCompteRenduImportDebiteur", acDesign

   Reports("EtaCompteRenduImportDebiteur").RecordSource = StrSQL

   DoCmd.Close acReport, "EtaCompteRenduImportDebiteur", acSaveYes

   On Error Resume Next

   DoCmd.OpenReport "EtaCompteRenduImportDebiteur", 2, , , acWindowNormal

   Select Case Err.Number
      Case 0

      Case 2501

         Err.Clear

      Case Else

         GoTo Err_CmdImprimer_Click

   End Select

   On Error GoTo Err_CmdImprimer_Click

   Do While DonneEtatRaport("EtaCompteRenduImportDebiteur", IntRaportEtat) = True

      DoEvents

   Loop

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdVisualiser_Click()
   Dim StrPath As String
   Dim FdgFichier As Office.FileDialog
   Dim VarPathFichier As Variant
   Dim StrFichier As String
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_CmdVisualiser_Click

    Set FdgFichier = Application.FileDialog(msoFileDialogOpen)

    FdgFichier.Filters.Clear

    FdgFichier.Filters.Add "Fichier CSV", "*.CSV"

    FdgFichier.AllowMultiSelect = False
    ' ouvre la boite de dialogue de choix de fichier avec le filtre des bases access
    Select Case FdgFichier.Show
        Case True

            VarPathFichier = FdgFichier.SelectedItems(1)
            ' recupere le nom de la base DATA
            StrFichier = Mid(VarPathFichier, InStrRev(VarPathFichier, "\") + 1)
            ' recupere le chemin de la base DATA
            StrPath = Left(VarPathFichier, InStrRev(VarPathFichier, "\"))

        Case False

         GoTo Exit_CmdVisualiser_Click

   End Select

   DoCmd.Hourglass True

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "XLS_IMP"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   DoCmd.TransferText acImportDelim, "FORMAT XLS DEBITEUR", "XLS_IMP", StrPath & StrFichier, True

   TxtFichier = StrFichier

   CmdImprimer.Enabled = True

   CmbFichierChamp.Visible = True

   CmbTableChamp.Visible = True

   LstChamp.Visible = True

   CmdChampAjouter.Visible = True

   CmdChampSupprimer.Visible = True

   LstChampControle.Visible = True

   CmdChampControleAjouter.Visible = True

   CmdChampControleSupprimer.Visible = True

Exit_CmdVisualiser_Click:

   DoCmd.Hourglass False

   Exit Sub

Err_CmdVisualiser_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdVisualiser_Click"

   Resume Exit_CmdVisualiser_Click
End Sub

Private Sub CmdImporter_Click()
   Dim DBBase As DAO.Database
   Dim WksEspace As DAO.Workspace
   Dim RstImport As DAO.Recordset
   Dim RstDebiteur As DAO.Recordset
   Dim StrSQL As String
   Dim StrSQLImport As String
   Dim StrSQLDebiteur As String
   Dim StrSQLJoin As String
   Dim StrNote As String
   Dim IntGrilleTauxNum As Integer
   Dim IntDelai As Integer
   Dim LngNumero As Long
   Dim BolDebiteurAdresse As Boolean
   Dim BolDebiteurTelephone As Boolean

   DoCmd.Hourglass True

   StrSQL = "SELECT [Déb_Code], "

   For LngNumero = 0 To LstChampControle.ListCount - 1

      StrSQLImport = StrSQLImport & "[" & LstChampControle.Column(0, LngNumero) & "], "

      StrSQLDebiteur = StrSQLDebiteur & "[" & LstChampControle.Column(2, LngNumero) & "], "

      StrSQLJoin = StrSQLJoin & "[XLS_IMP].[" & LstChampControle.Column(0, LngNumero) & "]=[SelDebiteurAdresse].[" & LstChampControle.Column(2, LngNumero) & "] AND "

   Next

   For LngNumero = 0 To LstChamp.ListCount - 1

      StrSQLImport = StrSQLImport & "[" & LstChamp.Column(0, LngNumero) & "], "

      StrSQLDebiteur = StrSQLDebiteur & "[" & LstChamp.Column(2, LngNumero) & "], "

   Next

   StrSQL = StrSQL & Left(StrSQLDebiteur, Len(StrSQLDebiteur) - 2) & "," & Left(StrSQLImport, Len(StrSQLImport) - 2) & " FROM XLS_IMP INNER JOIN SelDebiteurAdresse ON "

   StrSQL = StrSQL & Left(StrSQLJoin, Len(StrSQLJoin) - 5) & ";"

   Set WksEspace = DBEngine.Workspaces(0)

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   Set RstImport = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstDebiteur = DBBase.OpenRecordset("SELECT SelDebiteurAdresse.* FROM SelDebiteurAdresse;")

   Select Case RstImport.EOF
      Case True

      Case False

         RstImport.MoveLast

         RstImport.MoveFirst

         MajPatientez 0, 0, RstImport.RecordCount

         Do Until RstImport.EOF

            MajPatientez RstImport.AbsolutePosition

            BolDebiteurAdresse = False

            BolDebiteurTelephone = False

            StrNote = ""

            RstDebiteur.FindFirst "Déb_Code='" & RstImport("Déb_Code") & "'"

            RstDebiteur.Edit

            For LngNumero = 0 To LstChamp.ListCount - 1

               Select Case RstImport(LstChamp.Column(2, LngNumero))
                  Case RstImport(LstChamp.Column(0, LngNumero))

                  Case Else

                     Select Case Left(RstImport(LstChamp.Column(2, LngNumero)).Name, 6)
                        Case "AD_Tel"

                           BolDebiteurTelephone = True

                        Case Else

                           BolDebiteurAdresse = True

                     End Select

                     StrNote = StrNote & vbCrLf & RstImport(LstChamp.Column(2, LngNumero))

                     RstDebiteur(LstChamp.Column(2, LngNumero)) = RstImport(LstChamp.Column(0, LngNumero))

               End Select

            Next

            Select Case BolDebiteurAdresse
               Case True

                  RstDebiteur("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=1")

                  RstDebiteur("Déb_niv_relance") = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=1")

                  RstDebiteur("Déb_Date_proch_rel") = Null

               Case False

                  Select Case BolDebiteurTelephone
                     Case True

                        IntGrilleTauxNum = DLookup("Cli_GrilleTaux", "CLIENT", "Cli_Code = '" & RstDebiteur("Déb_Code_Client") & "'") '*** Numéro de la grille de taux du client

                        IntDelai = DLookup("Tx_Délai", "TAUX", "Tx_N° = 9 AND Tx_Grille = " & IntGrilleTauxNum & " AND Tx_Code_Action = 2")

                        RstDebiteur("Déb_Situation") = DLookup("CG_Code", "CODE_GESTION", "CG_Code_Action=1")

                        RstDebiteur("Déb_niv_relance") = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=1")

                        RstDebiteur("Déb_Date_proch_rel") = Date + IntDelai

                     Case False

                  End Select

            End Select

            Select Case BolDebiteurAdresse = True Or BolDebiteurTelephone = True
               Case True

                  RstDebiteur("AD_Note") = StrNote

               Case False

            End Select

            RstDebiteur.Update

            RstImport.MoveNext

         Loop

   End Select

   WksEspace.CommitTrans

Exit_CmdImporter_Click:

   DoCmd.Hourglass False

   DoCmd.Close acForm, "FrmPatientez"

   Exit Sub

Err_CmdImporter_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImporter_Click"

   Resume Exit_CmdImporter_Click
End Sub
