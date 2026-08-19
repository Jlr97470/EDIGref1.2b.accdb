Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim StrEtat As String
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer
   Dim StrCG_Code As String
   Dim StrSQLWhere As String
   Dim BolWhere As Boolean
   Dim StrDateDebut As String
   Dim StrDateFin As String

   On Error GoTo Err_ImpVisu

   StrCG_Code = Nz(CmbGrp_Code)

   StrDateDebut = Nz(TxtDateDebut)

   StrDateFin = Nz(TxtDateFin)

   Select Case StrDateDebut
      Case ""

      Case Else

         StrDateDebut = ConvDate(StrDateDebut)

   End Select

   Select Case StrDateFin
      Case ""

      Case Else

         StrDateFin = ConvDate(StrDateFin)

   End Select

   Select Case StrCG_Code
      Case 0

      Case Else

         StrSQLWhere = StrSQLWhere + "Cli_Code_Groupe='" & StrCG_Code & "'"

         BolWhere = True

   End Select

   Select Case CadChoix
      Case 1

         StrEtat = "EtaCompteRenduExportReglementDate"

         Select Case StrDateDebut
            Case ""

            Case Else

               StrSQLWhere = StrSQLWhere + IIf(BolWhere, " AND ", "") & "EXPR_Date_Virement>=" & StrDateDebut

               BolWhere = True

         End Select

         Select Case StrDateFin
            Case ""

            Case Else

               StrSQLWhere = StrSQLWhere + IIf(BolWhere, " AND ", "") & "EXPR_Date_Virement<=" & StrDateFin

               BolWhere = True

         End Select

      Case 2

         StrEtat = "EtaCompteRenduExportIrrecouvrable"

         Select Case StrDateDebut
            Case ""

            Case Else

               StrSQLWhere = StrSQLWhere + IIf(BolWhere, " AND ", "") & "EXPI_Date_Virement>=" & StrDateDebut

               BolWhere = True

         End Select

         Select Case StrDateFin
            Case ""

            Case Else

               StrSQLWhere = StrSQLWhere + IIf(BolWhere, " AND ", "") & "EXPI_Date_Virement<=" & StrDateFin

               BolWhere = True

         End Select

   End Select

   DoCmd.OpenReport StrEtat, acDesign

   Reports(StrEtat).Filter = StrSQLWhere

   Reports(StrEtat).FilterOn = True

   DoCmd.Close acReport, StrEtat, acSaveYes

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport StrEtat, IntVisu, , IIf(CadChoix = 1, "ERD_Montant_mo=", "EID_Montant_mo=") & cFRANC & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cFRANC & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport(StrEtat, IntRaportEtat) = True

            DoEvents

         Loop

         On Error Resume Next

         DoCmd.OpenReport StrEtat, IntVisu, , IIf(CadChoix = 1, "ERD_Montant_mo=", "EID_Montant_mo=") & cEURO & IIf(StrSQLWhere = "", "", " AND ") & StrSQLWhere, acWindowNormal, cEURO & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCompteRenduExportIrrecouvrable", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduExportIrrecouvrable", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduExportIrrecouvrable", acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

    Exit Function

Err_ImpVisu:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

    Resume Exit_ImpVisu
End Function

Private Sub AfficherChoix()
   Dim StrSQL As String
   Dim StrSQLWhere As String
   Dim StrDateDebut As String
   Dim StrDateFin As String

   On Error GoTo Err_AfficherChoix

   StrDateDebut = Nz(TxtDateDebut)

   StrDateFin = Nz(TxtDateFin)

   Select Case StrDateDebut
      Case ""

      Case Else

         StrDateDebut = ConvDate(StrDateDebut)

   End Select

   Select Case StrDateFin
      Case ""

      Case Else

         StrDateFin = ConvDate(StrDateFin)

   End Select

   DoCmd.Hourglass True

   SFrmExportation.Visible = False

   Select Case CadArchive
      Case 1 'En cours : visualisation des clients Auto du groupe

         CmdRechercher.Enabled = True 'dévérouille le bouton Rechercher

         SFrmExportation.LinkChildFields = ""

         SFrmExportation.SourceObject = "SFrmExportation"

         SFrmExportation.LinkChildFields = "Cli_Code_Groupe"

         Select Case CadChoix
            Case 1

               StrSQL = "SELECT DISTINCT CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_Code_Groupe "

               StrSQLWhere = "FROM CLIENT INNER JOIN (REGLEMENT RIGHT JOIN (REGLEMENT_DETAIL RIGHT JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON REGLEMENT_DETAIL.RegD_Id = ENCAISSEMENT.Enc_RegD_Id) ON REGLEMENT.Reg_Id = REGLEMENT_DETAIL.RegD_Reg_Id) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client " & _
                             "WHERE ENCAISSEMENT.Enc_Situation='R' AND ENCAISSEMENT.Enc_Prin_m1<>0 AND ENCAISSEMENT.Enc_ERD_Id Is Null" & IIf(StrDateDebut = "", "", " AND REGLEMENT.Reg_Date>=" & StrDateDebut) & IIf(StrDateFin = "", "", " AND REGLEMENT.Reg_Date<=" & StrDateFin)

            Case 2

               StrSQL = "SELECT DISTINCT CLIENT.Cli_Code, CLIENT.Cli_Nom, CLIENT.Cli_Code_Groupe "

               StrSQLWhere = "FROM CLIENT INNER JOIN (CODE_GESTION INNER JOIN DEBITEUR ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client " & _
                             "WHERE CODE_GESTION.CG_Code_Action=6 AND DEBITEUR.Déb_Lot_EID_Id Is Null AND DEBITEUR.Déb_Date_der_rel Is Not Null" & IIf(StrDateDebut = "", "", " AND DEBITEUR.Déb_Date_der_rel>=" & StrDateDebut) & IIf(StrDateFin = "", "", " AND DEBITEUR.Déb_Date_der_rel<=" & StrDateFin)

         End Select

         CurrentDb.Execute "DELETE TMP_CLIENT.* FROM TMP_CLIENT;", dbFailOnError

         CurrentDb.Execute "INSERT INTO TMP_CLIENT (TCli_Code) SELECT Cli_Code " & StrSQLWhere & " AND Cli_Code_Groupe='" & CmbGrp_Code & "';"

       Case 2 'Archive : visualisation des lots déjà traités

           CmdRechercher.Enabled = False 'vérouille le bouton Rechercher

           Select Case CadChoix
               Case 1
                   'décoche tous les lots
                   StrSQL = "SELECT DISTINCTROW EXPORTATION_REG.* FROM EXPORTATION_REG " & _
                                 IIf(StrDateDebut + StrDateFin = "", "", " WHERE ") & _
                                 IIf(StrDateDebut = "", "", "EXPORTATION_REG.EXPR_Date>=" & StrDateDebut) & _
                                 IIf(StrDateDebut <> "" And StrDateFin <> "", " AND ", "") & _
                                 IIf(StrDateFin = "", "", "EXPORTATION_REG.EXPR_Date<=" & StrDateFin)

                   SFrmExportation.LinkChildFields = ""

                   SFrmExportation.SourceObject = "SFrmExportationArchiveREG"

                   SFrmExportation.LinkChildFields = "EXPR_Lot_Site"

               Case 2
                   'décoche tous les lots
                   StrSQL = "SELECT DISTINCTROW EXPORTATION_IRR.* FROM EXPORTATION_IRR " & _
                                 IIf(StrDateDebut + StrDateFin = "", "", " WHERE ") & _
                                 IIf(StrDateDebut = "", "", "EXPORTATION_IRR.EXPI_Date>=" & StrDateDebut) & _
                                 IIf(StrDateDebut <> "" And StrDateFin <> "", " AND ", "") & _
                                 IIf(StrDateFin = "", "", "EXPORTATION_IRR.EXPI_Date<=" & StrDateFin)

                   SFrmExportation.LinkChildFields = ""

                   SFrmExportation.SourceObject = "SFrmExportationArchiveIRR"

                   SFrmExportation.LinkChildFields = "EXPI_Lot_Site"

           End Select

   End Select

   Select Case SFrmExportation.Form.RecordSource
      Case StrSQL & StrSQLWhere & ";"

      Case Else

         SFrmExportation.Form.RecordSource = StrSQL & StrSQLWhere & ";"

   End Select

   SFrmExportation.Visible = True

   DoCmd.Hourglass False

Exit_AfficherChoix:

   Exit Sub

Err_AfficherChoix:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : AfficherChoix"

   Resume Exit_AfficherChoix
End Sub

Private Sub Form_Load()

   On Error GoTo Err_Form_Load

   AfficherChoix

Exit_Form_Load:

   Exit Sub

Err_Form_Load:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Load"

   Resume Exit_Form_Load
End Sub

Private Sub CmbGrp_Code_AfterUpdate()

   On Error GoTo Err_CmbGrp_Code_AfterUpdate

   AfficherChoix

Exit_CmbGrp_Code_AfterUpdate:

   Exit Sub

Err_CmbGrp_Code_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmbGrp_Code_AfterUpdate"

   Resume Exit_CmbGrp_Code_AfterUpdate
End Sub

Private Sub TxtDateDebut_AfterUpdate()

   On Error GoTo Err_TxtDateDebut_AfterUpdate

   AfficherChoix

Exit_TxtDateDebut_AfterUpdate:

   Exit Sub

Err_TxtDateDebut_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDateDebut_AfterUpdate"

   Resume Exit_TxtDateDebut_AfterUpdate
End Sub

Private Sub TxtDateFin_AfterUpdate()

   On Error GoTo Err_TxtDateFin_AfterUpdate

   AfficherChoix

Exit_TxtDateFin_AfterUpdate:

   Exit Sub

Err_TxtDateFin_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : TxtDateFin_AfterUpdate"

   Resume Exit_TxtDateFin_AfterUpdate
End Sub

Private Sub CadArchive_AfterUpdate()

   On Error GoTo Err_CadArchive_AfterUpdate

   AfficherChoix

Exit_CadArchive_AfterUpdate:

   Exit Sub

Err_CadArchive_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CadArchive_AfterUpdate"

   Resume Exit_CadArchive_AfterUpdate
End Sub

Private Sub CadChoix_AfterUpdate()

   On Error GoTo Err_CadChoix_AfterUpdate

   AfficherChoix

Exit_CadChoix_AfterUpdate:

   Exit Sub

Err_CadChoix_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CadChoix_AfterUpdate"

   Resume Exit_CadChoix_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   CurrentDb.Execute "DELETE TMP_CLIENT.* FROM TMP_CLIENT;", dbFailOnError

   CurrentDb.Execute "DELETE TMP_REGLEMENT.* FROM TMP_REGLEMENT;", dbFailOnError

   CurrentDb.Execute "DELETE TMP_IRRECOUVRABLE.* FROM TMP_IRRECOUVRABLE;", dbFailOnError

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description
End Sub

Private Sub CmdRechercher_Click()
   Dim RstClient As DAO.Recordset
   Dim StrSQL As String
   Dim StrSQLWhere As String
   Dim StrDateDebut As String
   Dim StrDateFin As String

   On Error GoTo Err_CmdRechercher_Click

   DoCmd.Hourglass True

   Select Case IsNull(CmbGrp_Code)
      Case True

        MsgBox "Groupe incorrect !", vbExclamation, stSource

        DoCmd.Hourglass False

        GoTo Exit_CmdRechercher_Click

      Case False

         Select Case CadArchive
            Case 1 'En cours : visualisation des clients Auto du groupe

               StrDateDebut = Nz(TxtDateDebut)

               StrDateFin = Nz(TxtDateFin)

               Select Case StrDateDebut
                  Case ""

                  Case Else

                     StrDateDebut = ConvDate(StrDateDebut)

               End Select

               Select Case StrDateFin
                  Case ""

                  Case Else

                     StrDateFin = ConvDate(StrDateFin)

               End Select

               StrSQL = "SELECT CLIENT.Cli_MonnaieEnCours FROM TMP_CLIENT INNER JOIN CLIENT ON TMP_CLIENT.TCli_Code = CLIENT.Cli_Code " & _
                             "GROUP BY CLIENT.Cli_MonnaieEnCours;"

               Set RstClient = CurrentDb.OpenRecordset(StrSQL)

               Select Case RstClient.EOF
                  Case True

                     MsgBox "Aucun client n'est sélectionné!", vbInformation, stSource

                     RstClient.Close

                     DoCmd.Hourglass False

                     GoTo Exit_CmdRechercher_Click

                  Case False

                     RstClient.MoveLast

                     Select Case RstClient.RecordCount
                        Case 1

                        Case Else

                           MsgBox "Exportation impossible. Les clients sélectionnés n'utilisent pas tous la même monnaie!", vbExclamation, stSource

                           RstClient.Close

                           DoCmd.Hourglass False

                           GoTo Exit_CmdRechercher_Click

                     End Select

               End Select

               RstClient.Close

               Select Case CadChoix
                   Case 1

                        Select Case StrDateDebut
                           Case ""

                           Case Else

                              StrSQLWhere = StrSQLWhere & " AND REGLEMENT.Reg_Date>=" & StrDateDebut

                        End Select

                        Select Case StrDateFin
                           Case ""

                           Case Else

                              StrSQLWhere = StrSQLWhere & " AND REGLEMENT.Reg_Date<=" & StrDateFin

                        End Select

                        CurrentDb.Execute "DELETE TMP_REGLEMENT.* FROM TMP_REGLEMENT;", dbFailOnError

                        StrSQL = "INSERT INTO TMP_REGLEMENT ( TReg_Enc_ID ) SELECT DISTINCT ENCAISSEMENT.Enc_ID " & _
                                      "FROM REGLEMENT RIGHT JOIN (REGLEMENT_DETAIL RIGHT JOIN ((CLIENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) INNER JOIN TMP_CLIENT ON CLIENT.Cli_Code = TMP_CLIENT.TCli_Code) ON REGLEMENT_DETAIL.RegD_Id = ENCAISSEMENT.Enc_RegD_Id) ON REGLEMENT.Reg_Id = REGLEMENT_DETAIL.RegD_Reg_Id " & _
                                      "WHERE ENCAISSEMENT.Enc_Situation='R' AND ENCAISSEMENT.Enc_Prin_m1<>0 AND ENCAISSEMENT.Enc_ERD_Id Is Null" & StrSQLWhere & ";"

                        CurrentDb.Execute StrSQL, dbFailOnError

                        DoCmd.Hourglass False

                        Select Case DCount("TReg_Enc_Id", "TMP_REGLEMENT")
                           Case Is > 0

                              DoCmd.OpenForm "FrmExportationReglement", acNormal, , , , acDialog, CmbGrp_Code & ";" & StrDateDebut & ";" & StrDateFin & ";"

                           Case Else

                              MsgBox "Aucun règlement à exporter pour les clients sélectionnés!", vbInformation, stSource

                        End Select

                     Case 2

                        Select Case StrDateDebut
                           Case ""

                           Case Else

                              StrSQLWhere = StrSQLWhere & " AND DEBITEUR.Déb_Date_der_rel>=" & StrDateDebut

                        End Select

                        Select Case StrDateFin
                           Case ""

                           Case Else

                              StrSQLWhere = StrSQLWhere & " AND DEBITEUR.Déb_Date_der_rel<=" & StrDateFin

                        End Select

                        CurrentDb.Execute "DELETE TMP_IRRECOUVRABLE.* FROM TMP_IRRECOUVRABLE;", dbFailOnError

                        StrSQL = "INSERT INTO TMP_IRRECOUVRABLE ( TIrr_Déb_Code ) SELECT DISTINCT DEBITEUR.Déb_Code " & _
                                      "FROM TMP_CLIENT INNER JOIN (CODE_GESTION INNER JOIN (CLIENT INNER JOIN DEBITEUR ON CLIENT.Cli_Code = DEBITEUR.Déb_Code_Client) ON CODE_GESTION.CG_Code = DEBITEUR.Déb_Situation) ON TMP_CLIENT.TCli_Code = CLIENT.Cli_Code " & _
                                      "WHERE CODE_GESTION.CG_Code_Action=6 AND DEBITEUR.Déb_Lot_EID_Id Is Null  AND DEBITEUR.Déb_Date_der_rel Is Not Null" & StrSQLWhere & ";"

                        CurrentDb.Execute StrSQL, dbFailOnError

                        DoCmd.Hourglass False

                        Select Case DCount("TIrr_Déb_Code", "TMP_IRRECOUVRABLE")
                           Case Is > 0

                              DoCmd.OpenForm "FrmExportationIrrecouvrable", acNormal, , , , acDialog, CmbGrp_Code & ";" & StrDateDebut & ";" & StrDateFin & ";"

                           Case Else

                              MsgBox "Aucun irrécouvrable à exporter pour les clients sélectionnés!", vbInformation, stSource

                        End Select

                 End Select

             Case 2 'Archive : visualisation des lots déjà traités

         End Select

   End Select

Exit_CmdRechercher_Click:

   Exit Sub

Err_CmdRechercher_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRechercher_Click"

   Resume Exit_CmdRechercher_Click
End Sub
