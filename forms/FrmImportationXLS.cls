Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private StrFichier As String
Private StrPath As String

Private Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   Select Case SFrmImportationXLS.Form.Détail.Visible
      Case True

         Select Case BolSelection
            Case True

               StrSQL = "INSERT INTO TMP_IMP_XLS ([TIX_Ref Client], [TIX_N°]) SELECT [Ref Client], [N°] FROM XLS_IMP LEFT JOIN TMP_IMP_XLS ON XLS_IMP.[Ref Client] = TMP_IMP_XLS.[TIX_Ref Client]  WHERE [TIX_Ref Client] Is Null AND Not IsNull(DLookUp(" & Chr(34) & "Déb_Code" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code_Client='" & TxtIMPX_Cli_Code & "' AND Déb_Réf1_Client='" & Chr(34) & " & [Ref Client] & " & Chr(34) & "' AND Déb_Num_Chèque='" & Chr(34) & " & [N°] & " & Chr(34) & "'" & Chr(34) & "));"

            Case False

               StrSQL = "DELETE TMP_IMP_XLS.* FROM TMP_IMP_XLS WHERE Not IsNull(DLookUp(" & Chr(34) & "Déb_Code" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code_Client='" & Forms!FrmImportationXLS.Form!TxtIMPX_Cli_Code & "' AND Déb_Réf1_Client='" & Chr(34) & " & [TIX_Ref Client] & " & Chr(34) & "' AND Déb_Num_Chèque='" & Chr(34) & " & [TIX_N°] & " & Chr(34) & "'" & Chr(34) & "));"

         End Select

         DoCmd.Hourglass True

         CurrentDb.Execute StrSQL, dbFailOnError

         SFrmImportationXLS.Requery

         DoCmd.Hourglass False

      Case False

   End Select

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
End Function

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   SFrmImportationXLS.Form.Détail.Visible = False

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub TxtIMPX_Cli_Code_BeforeUpdate(Cancel As Integer)

   On Error GoTo Err_TxtIMPX_Cli_Code_BeforeUpdate

   Select Case IsNull(DLookup("Cli_Code", "CLIENT", "Cli_Code = '" & Nz(TxtIMPX_Cli_Code) & "'")) And Not IsNull(TxtIMPX_Cli_Code)
      Case True

         Cancel = True

         MsgBox "Ce Client n'existe pas !" & vbCrLf & "Pour le créer, sélectionnez -Fiche Client- dans le menu principal.", vbInformation, stSource

         CmdVisualiser.Enabled = False

         CmdImporter.Enabled = False

      Case False

   End Select

Exit_TxtIMPX_Cli_Code_BeforeUpdate:

   Exit Sub

Err_TxtIMPX_Cli_Code_BeforeUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtIMPX_Cli_Code_BeforeUpdate"

   Resume Exit_TxtIMPX_Cli_Code_BeforeUpdate
End Sub

Private Sub TxtIMPX_Cli_Code_AfterUpdate()

   On Error GoTo Err_TxtIMPX_Cli_Code_AfterUpdate

   CmdVisualiser.Enabled = True

   Select Case SFrmImportationXLS.Form.Détail.Visible
      Case True

         SFrmImportationXLS.Form.Requery

      Case False

   End Select

Exit_TxtIMPX_Cli_Code_AfterUpdate:

   Exit Sub

Err_TxtIMPX_Cli_Code_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtIMPX_Cli_Code_AfterUpdate"

   Resume Exit_TxtIMPX_Cli_Code_AfterUpdate
End Sub

Private Sub CmdImprimer_Click()

   On Error GoTo Err_CmdImprimer_Click

   DoCmd.OpenForm "FrmImportationXLSListe", acNormal, , , , acDialog

Exit_CmdImprimer_Click:

   Exit Sub

Err_CmdImprimer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

   Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdFermer_Click()
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_CmdFermer_Click

   SFrmImportationXLS.Form.Détail.Visible = False

   SFrmImportationXLS.Form.RecordSource = ""

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "XLS_IMP"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   CurrentDb.Execute "DELETE TMP_IMP_XLS.* FROM TMP_IMP_XLS;", dbFailOnError

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description
End Sub

Private Sub CmdClientListe_Click()

   On Error GoTo Err_CmdClientListe_Click

   DoCmd.OpenForm "FrmFicheClientListe", acNormal, , , acFormReadOnly, acDialog

   Select Case DLookup("Cli_Bloquer", "CLIENT", "Cli_Code =" & "'" & StrFicheClientListeCli_Code & "'")
      Case True

            MsgBox "Client en situation de blocage !"

            GoTo Exit_CmdClientListe_Click

      Case Else

   End Select

   Select Case StrFicheClientListeCli_Code
      Case ""

      Case Else

         TxtIMPX_Cli_Code = StrFicheClientListeCli_Code

         CmdVisualiser.Enabled = True

         Select Case SFrmImportationXLS.Form.Détail.Visible
            Case True

               SFrmImportationXLS.Form.Requery

            Case False

         End Select

   End Select

Exit_CmdClientListe_Click:

   Exit Sub

Err_CmdClientListe_Click:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CmdClientListe_Click
End Sub

Private Sub CmdVisualiser_Click()
   Dim StrSQL As String
   Dim FdgFichier As Office.FileDialog
   Dim VarPathFichier As Variant
   Dim TdfTable As DAO.TableDef
   Dim RstImportation As DAO.Recordset

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

    Select Case Left(StrFichier, 5)
       Case TxtIMPX_Cli_Code

          TxtIMPX_Fichier = StrFichier

       Case Else

          MsgBox "Le fichier correspond pas au client !!", vbCritical

          GoTo Exit_CmdVisualiser_Click

    End Select

   SFrmImportationXLS.Form.Détail.Visible = False

   SFrmImportationXLS.Form.RecordSource = ""

   TxtIMPX_Date.Requery

   CmdImporter.Enabled = True

   Select Case IsNull(TxtIMPX_Date)
      Case True

      Case False

         CmdImporter.Enabled = False

         MsgBox "Ce fichier a déjà été importé le " & TxtIMPX_Date & " !", vbExclamation, stSource

         GoTo Exit_CmdVisualiser_Click

   End Select

   DoCmd.Hourglass True

   SFrmImportationXLS.Form.Détail.Visible = False

   SFrmImportationXLS.Form.RecordSource = ""

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "XLS_IMP"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   DoCmd.TransferText acImportDelim, "FORMAT XLS", "XLS_IMP", StrPath & StrFichier, True

   CurrentDb.Execute "DELETE TMP_IMP_XLS.* FROM TMP_IMP_XLS;", dbFailOnError

   CurrentDb.Execute "INSERT INTO TMP_IMP_XLS ([TIX_Ref Client],[TIX_N°]) SELECT [Ref Client],[N°] FROM XLS_IMP WHERE IsNull(DLookUp(" & Chr(34) & "Déb_Code" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code_Client='" & TxtIMPX_Cli_Code & "' AND Déb_Réf1_Client='" & Chr(34) & " & [Ref Client] & " & Chr(34) & "' AND Déb_Num_Chèque='" & Chr(34) & " & [N°] & " & Chr(34) & "'" & Chr(34) & "));", dbFailOnError

   StrSQL = "SELECT XLS_IMP.* FROM XLS_IMP WHERE Not IsNull(DLookUp(" & Chr(34) & "Déb_Code" & Chr(34) & "," & Chr(34) & "DEBITEUR" & Chr(34) & "," & Chr(34) & "Déb_Code_Client='" & TxtIMPX_Cli_Code & "' AND Déb_Réf1_Client='" & Chr(34) & " & [Ref Client] & " & Chr(34) & "' AND Déb_Num_Chèque='" & Chr(34) & " & [N°] & " & Chr(34) & "'" & Chr(34) & "));"

   SFrmImportationXLS.Form.RecordSource = StrSQL

   SFrmImportationXLS.Form.Détail.Visible = True

   TxtIMPX_Date_Début.Requery

   TxtIMPX_Date_Fin.Requery

   TxtIMPX_Montant_m1.Requery

   TxtIMPX_Montant_m2.Requery

   TxtIMPX_Nb_Créances.Requery

   TxtIMPX_Date_Début.Visible = True

   TxtIMPX_Date_Fin.Visible = True

   TxtIMPX_Montant_m1.Visible = True

   TxtIMPX_Montant_m2.Visible = True

   TxtIMPX_Nb_Créances.Visible = True

   Select Case TxtIMPX_Nb_Créances
      Case 0

        CmdImporter.Enabled = False

        MsgBox "Ce fichier ne contient aucune créance.", vbExclamation, stSource

      Case Else

   End Select

   Set RstImportation = CurrentDb.OpenRecordset(StrSQL)

   Select Case RstImportation.EOF
      Case True

      Case False

         RstImportation.MoveLast

         MsgBox RstImportation.RecordCount & " doublons detectés."

         RstImportation.Close

   End Select

Exit_CmdVisualiser_Click:

   DoCmd.Hourglass False

   Set RstImportation = Nothing

   Exit Sub

Err_CmdVisualiser_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdVisualiser_Click"

   Resume Exit_CmdVisualiser_Click
End Sub

Private Sub CmdImporter_Click()
   Dim StrSQL As String
   Dim StrCodeClient As String
   Dim StrCodeDebiteur As String
   Dim IntCode_Action As Integer
   Dim IntMonnaie As Integer
   Dim IntNbCreance As Integer
   Dim DatDateDebut As Date
   Dim DatDateFin As Date
   Dim CurMontant_m1 As Currency
   Dim CurMontant_m2 As Currency
   Dim DBBase As DAO.Database
   Dim WksEspace As DAO.Workspace
   Dim RstDebiteur As DAO.Recordset
   Dim RstAdresseDebiteur As DAO.Recordset
   Dim RstClient As DAO.Recordset
   Dim RstImpayer As DAO.Recordset

   On Error GoTo Err_CmdImporter_Click

   DoCmd.Hourglass True

   Set WksEspace = DBEngine.Workspaces(0)

   Set DBBase = WksEspace.Databases(0)

   WksEspace.BeginTrans

   Set RstClient = DBBase.OpenRecordset("CLIENT")

   Set RstDebiteur = DBBase.OpenRecordset("DEBITEUR")

   Set RstAdresseDebiteur = DBBase.OpenRecordset("ADRESSE_DEBITEUR")

   StrSQL = "SELECT XLS_IMP.* FROM XLS_IMP INNER JOIN TMP_IMP_XLS ON (XLS_IMP.N° = TMP_IMP_XLS.TIX_N°) AND (XLS_IMP.[Ref Client] = TMP_IMP_XLS.[TIX_Ref Client]);"

   Set RstImpayer = DBBase.OpenRecordset(StrSQL, DB_OPEN_DYNASET)

   RstImpayer.MoveLast

   RstImpayer.MoveFirst

   MajPatientez 0, 0, RstImpayer.RecordCount

   StrCodeClient = TxtIMPX_Cli_Code

   RstClient.FindFirst "Cli_Code='" & StrCodeClient & "'"

   IntMonnaie = RstClient("Cli_MonnaieEnCours")

   DatDateDebut = RstImpayer("Date")

   DatDateFin = RstImpayer("Date")

   Do Until RstImpayer.EOF = True

      MajPatientez RstImpayer.AbsolutePosition

      IntNbCreance = IntNbCreance + 1

      Select Case StrCodeDebiteur
         Case ""

            StrCodeDebiteur = DonnerNouveauCodeDebiteur(StrCodeClient)

         Case Else

            StrCodeDebiteur = Format(Int(Right(StrCodeDebiteur, 5)) + 1, "00000")

            StrCodeDebiteur = StrCodeClient & StrCodeDebiteur

      End Select

      RstDebiteur.AddNew

      RstDebiteur("Déb_Date_Cré") = Date

      RstDebiteur("Déb_Code_Client") = StrCodeClient

      RstDebiteur("Déb_Code") = StrCodeDebiteur

      RstDebiteur("Déb_Nom") = RstImpayer("Nom Prénom")

      RstDebiteur("Déb_Typ_Cré") = RstImpayer("Type") & " " & RstImpayer("Banque") & " N°" & RstImpayer("N°") & " DU " & RstImpayer("Date")

      Select Case IntMonnaie
          Case cFRANC

              RstDebiteur("Déb_Mont_Prin_m1") = CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

              RstDebiteur("Déb_Mont_Prin_m2") = CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric))) * DonnerTauxConvertion(cFRANC)

          Case cEURO

              RstDebiteur("Déb_Mont_Prin_m2") = CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

              RstDebiteur("Déb_Mont_Prin_m1") = CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric))) * DonnerTauxConvertion(cEURO)

      End Select

      RstDebiteur("Déb_Mont_Prin_mo") = IntMonnaie

      RstDebiteur("Déb_MonnaieEnCours") = IntMonnaie

      RstDebiteur("Déb_Réf1_Client") = RstImpayer("Ref Client")

      RstDebiteur("Déb_Réf2_Client") = StrFichier

      RstDebiteur("Déb_Situation") = DLookup("Situation_Déb_Défaut", "PARAMETRE_GREF")

      IntCode_Action = Nz(DLookup("CG_Code_Action", "CODE_GESTION", "CG_Code='" & RstDebiteur("Déb_Situation") & "'"), 0)

      RstDebiteur("Déb_niv_relance") = DLookup("Rel_N°", "RELANCE", "Rel_Code_Action=" & IntCode_Action)

      RstDebiteur("Déb_Mont_Origine_Cli") = CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

      RstDebiteur("Déb_RIB_Dom") = RstImpayer("Compte Bancaire")

      RstDebiteur("Déb_Date_Chèque") = RstImpayer("Date")

      RstDebiteur("Déb_Num_Chèque") = RstImpayer("N°")

      RstDebiteur("Déb_Nom_Banque") = RstImpayer("Banque")

      Select Case IntMonnaie
         Case cFRANC

            RstDebiteur("Déb_Frais_m1") = DonnerMontantFrais(StrCodeClient, RstDebiteur("Déb_Mont_Prin_m1"), IntMonnaie)

            RstDebiteur("Déb_Frais_m2") = RstDebiteur("Déb_Frais_m1") * DonnerTauxConvertion(cFRANC)

         Case cEURO

            RstDebiteur("Déb_Frais_m2") = DonnerMontantFrais(StrCodeClient, RstDebiteur("Déb_Mont_Prin_m2"), IntMonnaie)

            RstDebiteur("Déb_Frais_m1") = RstDebiteur("Déb_Frais_m2") * DonnerTauxConvertion(cEURO)

      End Select

      RstDebiteur.Update

      RstAdresseDebiteur.AddNew

      RstAdresseDebiteur("AD_Code_Débiteur") = StrCodeDebiteur

      RstAdresseDebiteur("AD_TypeAdresse_ID") = Nz(DLookup("Adresse_Courrier", "PARAMETRE_GREF"))

      RstAdresseDebiteur("AD_Adr1") = Left(Trim(RstImpayer("Adresse 1")), 50)

      RstAdresseDebiteur("AD_Adr2") = Left(Trim(RstImpayer("Adresse 2")), 50)

      RstAdresseDebiteur("AD_Adr3") = Left(Trim(RstImpayer("Adresse 3")), 50)

      RstAdresseDebiteur("AD_Adr4") = Left(Trim(RstImpayer("Adresse 4")), 50)

      RstAdresseDebiteur("AD_CodePostal") = RstImpayer("Code Postal")

      RstAdresseDebiteur("AD_Ville") = RstImpayer("Ville")

      RstAdresseDebiteur("AD_Pays") = RstImpayer("Pays")

      RstAdresseDebiteur("AD_Tel") = RstImpayer("Téléphone")

      RstAdresseDebiteur.Update

      Select Case RstImpayer("Date")
         Case Is < DatDateDebut

            DatDateDebut = RstImpayer("Date")

         Case Is > DatDateFin

            DatDateFin = RstImpayer("Date")

      End Select

      Select Case IntMonnaie
         Case 1

            CurMontant_m1 = CurMontant_m1 + CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

         Case 2

            CurMontant_m2 = CurMontant_m2 + CCur(Replace(RstImpayer("Montant"), DonnerSeparateur(cAlpha), DonnerSeparateur(cNumeric)))

      End Select

      RstImpayer.MoveNext

   Loop

   RstClient.FindFirst "Cli_Code='" & StrCodeClient & "'"

   RstClient.Edit

   RstClient("Cli_Date_der_cré") = Date

   RstClient("Cli_Code_der_déb") = StrCodeDebiteur

   RstClient.Update

   RstClient.Close

   RstImpayer.Close

   RstDebiteur.Close

   RstAdresseDebiteur.Close

   StrSQL = "DELETE TMP_IMP_XLS FROM TMP_IMP_XLS;"

   DBBase.Execute StrSQL, dbFailOnError

   StrSQL = "INSERT INTO IMPORTATIONXLS (IMPX_Date, IMPX_Cli_Code, IMPX_Fichier, IMPX_Date_Début, IMPX_Date_Fin, IMPX_Nb_Créances, IMPX_Montant_m1, IMPX_Montant_m2) VALUES (Date()," & Chr(34) & TxtIMPX_Cli_Code & Chr(34) & "," & Chr(34) & TxtIMPX_Fichier & Chr(34) & "," & ConvDate(DatDateDebut) & "," & ConvDate(DatDateFin) & "," & IntNbCreance & ",CCur(" & Chr(34) & CurMontant_m1 & Chr(34) & "),CCur(" & Chr(34) & CurMontant_m2 & Chr(34) & "));"

   DBBase.Execute StrSQL, dbFailOnError

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

   Set RstAdresseDebiteur = Nothing

   Set DBBase = Nothing

   Set WksEspace = Nothing

   Exit Sub

Err_CmdImporter_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImporter_Click"

   Resume Exit_CmdImporter_Click
End Sub
