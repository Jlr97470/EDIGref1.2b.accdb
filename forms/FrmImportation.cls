Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Sub CmdFermer_Click()
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_CmdFermer_Click

   SFrmImportation.Form.Détail.Visible = False

   SFrmImportation.Form.RecordSource = ""

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "BLITZ_IMP", "BLITZ_DEB", "BLITZ_ADR", "BLITZ_INT"

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

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmImportationListe", acNormal, , , , acDialog

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdVisualiser_Click()
   Dim StrSQL As String
   Dim VarPathFichier As Variant
   Dim StrPath As String
   Dim StrFichier As String
   Dim TdfTable As DAO.TableDef
   Dim FdgFichier As Office.FileDialog
   Dim RstImportation As DAO.Recordset

   On Error GoTo Err_CmdVisualiser_Click

    Set FdgFichier = Application.FileDialog(msoFileDialogOpen)

    FdgFichier.Filters.Clear

    FdgFichier.Filters.Add "Fichier TXT", "*.TXT"

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

   SFrmImportation.Form.Détail.Visible = False

   SFrmImportation.Form.RecordSource = ""

   CmdImporter.Enabled = True

   TxtIMP_Lot_Date = Mid(StrFichier, 1, 8)

   TxtIMP_Lot_Num = Mid(StrFichier, 9, 2)

   TxtIMP_Lot_Site = Mid(StrFichier, 11, 5)

   TxtIMP_Lot_CodeGREF = Mid(StrFichier, 16, 5)

   TxtIMP_Lot_Contrat = Mid(StrFichier, 21, 5)

   TxtIMP_Date.Requery

   Select Case IsNull(TxtIMP_Date)
      Case True

      Case False

         CmdImporter.Enabled = False

         MsgBox "Ce fichier a déjà été importé le " & TxtIMP_Date & " !", vbExclamation, stSource

         GoTo Exit_CmdVisualiser_Click

   End Select

   DoCmd.Hourglass True

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "BLITZ_IMP", "BLITZ_DEB", "BLITZ_ADR", "BLITZ_INT"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   AttacheFichierBlitzCash StrPath, TxtIMP_Lot_Date & TxtIMP_Lot_Num & TxtIMP_Lot_Site & TxtIMP_Lot_CodeGREF & TxtIMP_Lot_Contrat

   StrSQL = "SELECT DLookUp(" & Chr(34) & "Cli_Code_Groupe" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_Code_Groupe, DLookUp(" & Chr(34) & "Cli_Code" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_Code, BLITZ_IMP.Imp_Code_Etablissement, BLITZ_IMP.Imp_Nom_Etablissement, Count(BLITZ_IMP.Imp_Cle_Impaye) AS Imp_NbCreance," & _
                  "DLookUp(" & Chr(34) & "Cli_MonnaieEnCours" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_MonnaieEnCours, Sum(CCur(Replace(BLITZ_IMP.Imp_Reste_Du,DonnerSeparateur(1),DonnerSeparateur(0)))) AS Imp_Montant " & _
                  "FROM BLITZ_IMP " & _
                  "GROUP BY DLookUp(" & Chr(34) & "Cli_Code_Groupe" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & "), " & _
                  "DLookUp(" & Chr(34) & "Cli_Code" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & "), BLITZ_IMP.Imp_Code_Etablissement, BLITZ_IMP.Imp_Nom_Etablissement, DLookUp(" & Chr(34) & "Cli_MonnaieEnCours" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ");"

   SFrmImportation.Form.RecordSource = StrSQL

   SFrmImportation.Form.Détail.Visible = True

   Set RstImportation = CurrentDb.OpenRecordset(StrSQL)

   TxtImp_Montant_m1 = 0

   TxtImp_Montant_m2 = 0

   Do Until RstImportation.EOF = True

      Select Case RstImportation("Cli_MonnaieEnCours")
         Case 1

            TxtImp_Montant_m1 = TxtImp_Montant_m1 + RstImportation("Imp_Montant")

         Case 2

            TxtImp_Montant_m2 = TxtImp_Montant_m2 + RstImportation("Imp_Montant")

      End Select

      RstImportation.MoveNext

   Loop

   TxtIMP_Date_Début.Requery

   TxtIMP_Date_Fin.Requery

   TxtIMP_Nb_Créances.Requery

   TxtIMP_Date_Début.Visible = True

   TxtIMP_Date_Fin.Visible = True

   TxtImp_Montant_m1.Visible = True

   TxtImp_Montant_m2.Visible = True

   TxtIMP_Nb_Créances.Visible = True

   DoCmd.Hourglass False

   Select Case TxtIMP_Nb_Créances
      Case 0

        CmdImporter.Enabled = False

        MsgBox "Ce lot ne contient aucune créance.", vbExclamation, stSource

      Case Else

   End Select

Exit_CmdVisualiser_Click:

   Exit Sub

Err_CmdVisualiser_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdVisualiser_Click"

   Resume Exit_CmdVisualiser_Click
End Sub

Private Sub CmdImporter_Click()
   Dim StrMsg As String
   Dim StrSQL As String
   Dim StrCodeGroupe As String
   Dim StrCodeClient As String
   Dim StrCodeEtab As String
   Dim StrCodeDebiteur As String
   Dim IntCode_Action As Integer
   Dim IntMonnaie As Integer
   Dim LngIMP_Id As Long
   Dim DBBase As DAO.Database
   Dim WksEspace As DAO.Workspace
   Dim RstDebiteur As DAO.Recordset
   Dim RstAdresseDebiteur As DAO.Recordset
   Dim RstClient As DAO.Recordset
   Dim RstImpayer As DAO.Recordset
   Dim RstImportation As DAO.Recordset
   Dim RstImportationDetail As DAO.Recordset

   On Error GoTo Err_CmdImporter_Click

   DoCmd.Hourglass True

   StrSQL = "SELECT DLookUp(" & Chr(34) & "Cli_Code_Groupe" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_Code_Groupe, DLookUp(" & Chr(34) & "Cli_Code" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_Code, BLITZ_IMP.Imp_Code_Etablissement, BLITZ_IMP.Imp_Nom_Etablissement, Count(BLITZ_IMP.Imp_Cle_Impaye) AS Imp_NbCreance," & _
                  "DLookUp(" & Chr(34) & "Cli_MonnaieEnCours" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ") AS Cli_MonnaieEnCours, Sum(CCur(Replace(BLITZ_IMP.Imp_Reste_Du,DonnerSeparateur(1),DonnerSeparateur(0)))) AS Imp_Montant " & _
                  "FROM BLITZ_IMP " & _
                  "GROUP BY DLookUp(" & Chr(34) & "Cli_Code_Groupe" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & "), " & _
                  "DLookUp(" & Chr(34) & "Cli_Code" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & "), BLITZ_IMP.Imp_Code_Etablissement, BLITZ_IMP.Imp_Nom_Etablissement, DLookUp(" & Chr(34) & "Cli_MonnaieEnCours" & Chr(34) & "," & Chr(34) & "CLIENT" & Chr(34) & "," & Chr(34) & "Cli_Code_Groupe='" & TxtIMP_Lot_Site & "' AND Cli_Code_Etab='" & Chr(34) & " & [Imp_Code_Etablissement] & " & Chr(34) & "'" & Chr(34) & ");"

   Set WksEspace = DBEngine.Workspaces(0)

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   Set RstImportationDetail = DBBase.OpenRecordset(StrSQL)

   RstImportationDetail.MoveFirst

   IntMonnaie = Nz(RstImportationDetail("Cli_MonnaieEnCours"))

   StrCodeGroupe = Nz(RstImportationDetail("Cli_Code_Groupe"))

   Do Until RstImportationDetail.EOF = True

      Select Case IsNull(RstImportationDetail("Cli_Code"))
         Case True

            MsgBox "IMPORTATION IMPOSSIBLE !" & vbCrLf & "Un ou plusieurs clients ne sont pas référencés dans la base GREF.", vbInformation, stSource

            RstImportationDetail.Close

            WksEspace.Rollback

            GoTo Exit_CmdImporter_Click

         Case False

      End Select

      Select Case RstImportationDetail("Cli_Code_Groupe")
         Case StrCodeGroupe

         Case Else

            MsgBox "IMPORTATION IMPOSSIBLE !" & vbCrLf & "Les clients n'appartiennent pas tous au même groupe.", vbInformation, stSource

            RstImportationDetail.Close

            WksEspace.Rollback

            GoTo Exit_CmdImporter_Click

      End Select

      Select Case RstImportationDetail("Cli_MonnaieEnCours")
         Case IntMonnaie

         Case Else

            MsgBox "IMPORTATION IMPOSSIBLE !" & vbCrLf & "Les clients de ce lot n'utilisent pas tous la même devise.", vbInformation, stSource

            RstImportationDetail.Close

            WksEspace.Rollback

            GoTo Exit_CmdImporter_Click

      End Select

      RstImportationDetail.MoveNext

   Loop

   Set RstImportation = DBBase.OpenRecordset("IMPORTATION")

   RstImportation.AddNew

   RstImportation("IMP_Date") = Date

   RstImportation("IMP_Lot_Date") = TxtIMP_Lot_Date

   RstImportation("IMP_Lot_Num") = TxtIMP_Lot_Num

   RstImportation("IMP_Lot_Site") = TxtIMP_Lot_Site

   RstImportation("IMP_Lot_CodeGREF") = TxtIMP_Lot_CodeGREF

   RstImportation("IMP_Lot_Contrat") = TxtIMP_Lot_Contrat

   RstImportation("IMP_Date_Début") = TxtIMP_Date_Début

   RstImportation("IMP_Date_Fin") = TxtIMP_Date_Fin

   RstImportation("IMP_Montant_m1") = TxtImp_Montant_m1

   RstImportation("IMP_Montant_m2") = TxtImp_Montant_m2

   LngIMP_Id = RstImportation("IMp_Id")

   RstImportation.Update

   RstImportationDetail.MoveFirst

   Do Until RstImportationDetail.EOF = True

      StrSQL = "INSERT INTO IMPORTATION_DETAIL (IMPD_Code, IMPD_CodeEtab, IMPD_NomEtab, IMPD_CodeGroupe, IMPD_CodeClient, IMPD_NbCréances, IMPD_Montant_m1, IMPD_Montant_m2, IMPD_Montant_mo) " & _
                  "VALUES(" & LngIMP_Id & "," & Chr(34) & RstImportationDetail("Imp_Code_Etablissement") & Chr(34) & "," & Chr(34) & RstImportationDetail("Imp_Nom_Etablissement") & Chr(34) & "," & Chr(34) & RstImportationDetail("Cli_Code_Groupe") & Chr(34) & "," & Chr(34) & RstImportationDetail("Cli_Code") & Chr(34) & "," & _
                  RstImportationDetail("Imp_NbCreance") & ",CCur(" & Chr(34) & IIf(RstImportationDetail("Cli_MonnaieEnCours") = 1, RstImportationDetail("Imp_Montant"), 0) & Chr(34) & "),CCur(" & Chr(34) & IIf(RstImportationDetail("Cli_MonnaieEnCours") = 2, RstImportationDetail("Imp_Montant"), 0) & Chr(34) & ")," & RstImportationDetail("Cli_MonnaieEnCours") & ");"

      DBBase.Execute StrSQL, dbFailOnError

      RstImportationDetail.MoveNext

   Loop

   RstImportationDetail.Close

   Set RstClient = DBBase.OpenRecordset("CLIENT")

   Set RstDebiteur = DBBase.OpenRecordset("DEBITEUR")

   Set RstAdresseDebiteur = DBBase.OpenRecordset("ADRESSE_DEBITEUR")

   StrSQL = "SELECT IMPORTATION_DETAIL.* FROM IMPORTATION_DETAIL WHERE IMPD_Code=" & LngIMP_Id & ";"

   Set RstImportationDetail = DBBase.OpenRecordset(StrSQL)

   StrSQL = "SELECT BLITZ_IMP.*, BLITZ_DEB.*, BLITZ_ADR.*, BLITZ_INT.* FROM ((BLITZ_IMP LEFT JOIN BLITZ_DEB ON BLITZ_IMP.Imp_Cle_Debiteur = BLITZ_DEB.Deb_Cle_Debiteur) LEFT JOIN BLITZ_ADR ON BLITZ_DEB.Deb_Cle_Debiteur = BLITZ_ADR.Adr_Cle_Debiteur) LEFT JOIN BLITZ_INT ON BLITZ_ADR.Adr_Cle_Etab_Debiteur = BLITZ_INT.Int_Cle_Etab_Debiteur ORDER BY Imp_Code_Etablissement, Imp_Cle_Impaye;"

   Set RstImpayer = DBBase.OpenRecordset(StrSQL, DB_OPEN_DYNASET)

   RstImpayer.MoveLast

   RstImpayer.MoveFirst

   MajPatientez 0, 0, RstImpayer.RecordCount

   Do Until RstImpayer.EOF = True

      MajPatientez RstImpayer.AbsolutePosition

      RstClient.FindFirst "Cli_Code_Groupe = '" & StrCodeGroupe & "' And Cli_Code_Etab='" & RstImpayer("Imp_Code_Etablissement") & "'"

      Select Case RstClient("Cli_Code")
         Case StrCodeClient

            StrCodeDebiteur = Format(Int(Right(StrCodeDebiteur, 5)) + 1, "00000")

            StrCodeDebiteur = StrCodeClient & StrCodeDebiteur

         Case Else

            StrCodeClient = RstClient("Cli_Code")

            StrCodeDebiteur = DonnerNouveauCodeDebiteur(StrCodeClient)

      End Select

      RstDebiteur.AddNew

      RstDebiteur("Déb_Date_Cré") = Date

      RstDebiteur("Déb_Code_Client") = StrCodeClient

      RstDebiteur("Déb_Code") = StrCodeDebiteur

      RstDebiteur("Déb_Nom") = RstImpayer("Deb_Nom") & " " & RstImpayer("Deb_Prenom")

      RstDebiteur("Déb_Typ_Cré") = RstImpayer("Imp_Type_Traitement") & " " & RstImpayer("Imp_CMC7_NumCheque") & " " & RstImpayer("Imp_Nom_Banque") & " " & RstImpayer("Imp_Date_Cheque")

      Select Case IntMonnaie
          Case cFRANC

              RstDebiteur("Déb_Mont_Prin_m1") = CCur(Replace(RstImpayer("Imp_Reste_Du"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

              RstDebiteur("Déb_Mont_Prin_m2") = CCur(Replace(RstImpayer("Imp_Reste_Du"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric))) * DonnerTauxConvertion(cFRANC)

          Case cEURO

              RstDebiteur("Déb_Mont_Prin_m2") = CCur(Replace(RstImpayer("Imp_Reste_Du"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

              RstDebiteur("Déb_Mont_Prin_m1") = CCur(Replace(RstImpayer("Imp_Reste_Du"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric))) * DonnerTauxConvertion(cEURO)

      End Select

      RstDebiteur("Déb_Mont_Prin_mo") = IntMonnaie

      RstDebiteur("Déb_MonnaieEnCours") = IntMonnaie

      RstDebiteur("Déb_Réf1_Client") = TxtIMP_Lot_Date & TxtIMP_Lot_Num & TxtIMP_Lot_Site & TxtIMP_Lot_CodeGREF & TxtIMP_Lot_Contrat

      RstDebiteur("Déb_Réf2_Client") = RstImpayer("Imp_Cle_Debiteur")

      RstDebiteur("Déb_Clé_Impayé") = RstImpayer("Imp_Cle_Impaye")

      RstDebiteur("Déb_Situation") = DonnerSituationGREF(StrCodeGroupe, Nz(RstImpayer("Imp_Motif_Impaye")))

      IntCode_Action = Nz(DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & RstDebiteur("Déb_Situation") & "'"), 0)

      RstDebiteur("Déb_niv_relance") = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

      RstDebiteur("Déb_Mont_Origine_Cli") = CCur(Replace(RstImpayer("Imp_Montant_TTC"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

      RstDebiteur("Déb_Mont_Récup_Cli") = CCur(Replace(RstImpayer("Imp_Regle"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

      RstDebiteur("Déb_RIB_Banque") = RstImpayer("Imp_Code_Banque")

      RstDebiteur("Déb_RIB_Guichet") = RstImpayer("Imp_Code_Guichet")

      RstDebiteur("Déb_RIB_Compte") = Left(RstImpayer("Imp_Num_Compte"), 11)

      RstDebiteur("Déb_RIB_Clé") = RstImpayer("Imp_Cle_RIB")

      RstDebiteur("Déb_RIB_Dom") = RstImpayer("Imp_Domiciliation")

      RstDebiteur("Déb_Date_Chèque") = RstImpayer("Imp_Date_Cheque")

      RstDebiteur("Déb_Num_Chèque") = RstImpayer("Imp_CMC7_NumCheque")

      RstDebiteur("Déb_Nom_Banque") = RstImpayer("Imp_Nom_Banque")

      RstImportationDetail.FindFirst "IMPD_CodeClient='" & StrCodeClient & "'"

      RstDebiteur("Déb_Lot_IMPD_Id") = RstImportationDetail("IMPD_Id")

      Select Case IntMonnaie
         Case cFRANC

            RstDebiteur("Déb_Frais_m1") = DonnerMontantFrais(StrCodeClient, RstDebiteur("Déb_Mont_Prin_m1"), IntMonnaie)

            RstDebiteur("Déb_Frais_m2") = RstDebiteur("Déb_Frais_m1") * DonnerTauxConvertion(cFRANC)

         Case cEURO

            RstDebiteur("Déb_Frais_m2") = DonnerMontantFrais(StrCodeClient, RstDebiteur("Déb_Mont_Prin_m2"), IntMonnaie)

            RstDebiteur("Déb_Frais_m1") = RstDebiteur("Déb_Frais_m2") * DonnerTauxConvertion(cEURO)

      End Select

      RstDebiteur.Update

      RstClient.Edit

      RstClient("Cli_Date_der_cré") = Date

      RstClient("Cli_Code_der_déb") = StrCodeDebiteur

      RstClient.Update

      RstAdresseDebiteur.AddNew

      RstAdresseDebiteur("AD_Code_Débiteur") = StrCodeDebiteur

      RstAdresseDebiteur("AD_TypeAdresse_ID") = Nz(DLookup("Adresse_Courrier", "PARAMETRE_GREF"))

      RstAdresseDebiteur("AD_Adr1") = Left(Trim(RstImpayer("Adr_Adr1")), 50)

      RstAdresseDebiteur("AD_Adr2") = Left(Trim(RstImpayer("Adr_Adr2")), 50)

      RstAdresseDebiteur("AD_Adr3") = Left(Trim(RstImpayer("Adr_Adr3")), 50)

      RstAdresseDebiteur("AD_Adr4") = Left(Trim(RstImpayer("Adr_Adr4")), 50)

      RstAdresseDebiteur("AD_CodePostal") = RstImpayer("Adr_CodePos")

      RstAdresseDebiteur("AD_Ville") = RstImpayer("Adr_Commune")

      RstAdresseDebiteur("AD_Pays") = RstImpayer("Adr_Pays")

      RstAdresseDebiteur("AD_Tel") = RstImpayer("Adr_Tel")

      RstAdresseDebiteur("AD_Fax") = RstImpayer("Adr_Fax")

      RstAdresseDebiteur("AD_email") = RstImpayer("Adr_Email")

      RstAdresseDebiteur("AD_NomComplémentaire") = RstImpayer("Int_Nom") & " " & RstImpayer("Int_Prenom")

      RstAdresseDebiteur.Update

      RstImpayer.MoveNext

   Loop

   RstClient.Close

   RstImpayer.Close

   RstDebiteur.Close

   RstAdresseDebiteur.Close

   WksEspace.CommitTrans

   DBBase.Close

   WksEspace.Close

   CmdVisualiser.SetFocus

   CmdImporter.Enabled = False

   Me.Requery

   DoCmd.Hourglass False

   MsgBox "IMPORTATION TERMINEE.", vbInformation, stSource

Exit_CmdImporter_Click:

   DoCmd.Close acForm, "FrmPatientez"

   Set RstClient = Nothing

   Set RstImpayer = Nothing

   Set RstDebiteur = Nothing

   Set RstImportation = Nothing

   Set RstImportationDetail = Nothing

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Sub

Err_CmdImporter_Click:

   Select Case TypeName(RstImpayer)
      Case "Nothing"

      Case Else

         Select Case RstImpayer.EOF
            Case True

            Case False

               StrMsg = "[Débiteur :" & RstImpayer("Deb_Nom") & " " & RstImpayer("Deb_Prenom") & "]"

         End Select

   End Select

   Select Case MsgBox("Erreur " & Err.Number & " : " & Err.description & vbCrLf & StrMsg & vbCrLf & vbCrLf & "Voulez vous annuler l'importation ?", vbYesNo + vbCritical, stSource)
      Case vbNo

         Select Case MsgBox("Si vous continuez, seules les informations correctes seront importées." & vbCrLf & "Etes-vous sûrs de vouloir continuer l'importation ?", vbYesNo + vbDefaultButton2 + vbExclamation, stSource)
            Case vbYes

               Resume Next

            Case Else

         End Select

      Case Else

   End Select

   WksEspace.Rollback

   Resume Exit_CmdImporter_Click
End Sub
