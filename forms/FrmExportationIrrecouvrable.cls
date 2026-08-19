Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private StrGrp_Code As String
Private StrDateDebut As String
Private StrDateFin As String

Public Function ToutRien(ByVal BolSelection As Boolean)
   Dim StrSQL As String

   On Error GoTo Err_ToutRien

   StrSQL = "UPDATE TMP_IRRECOUVRABLE SET TMP_IRRECOUVRABLE.TIrr_Flag = " & IIf(BolSelection, "True", "False") & ";"

   CurrentDb.Execute StrSQL, dbFailOnError

   Me.Requery

Exit_ToutRien:

   Exit Function

Err_ToutRien:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ToutRien"

   Resume Exit_ToutRien
End Function

Private Sub Form_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Form_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case Is < 1

         Cancel = True

         GoTo Exit_Form_Open

      Case Else

   End Select

   StrGrp_Code = StrParam(0)

   StrDateDebut = StrParam(1)

   StrDateFin = StrParam(2)

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub ChkTIrr_Flag_AfterUpdate()

   On Error GoTo Err_ChkTIrr_Flag_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTIrr_Flag_AfterUpdate:

   Exit Sub

Err_ChkTIrr_Flag_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkTIrr_Flag_AfterUpdate"

   Resume Exit_ChkTIrr_Flag_AfterUpdate
End Sub

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description
End Sub

Private Sub CmdTransmettre_Click()
   Dim WksEspace As DAO.Workspace
   Dim DBBase As DAO.Database
   Dim RstIrrecouvrable As DAO.Recordset
   Dim RstIrrecouvrableExport As DAO.Recordset
   Dim RstIrrecouvrableExportDetail As DAO.Recordset
   Dim LngEXPI_Id As Long
   Dim LngEID_Id As Long
   Dim StrSQL As String
   Dim StrLotNum As String
   Dim StrLotDate As String
   Dim StrLotSite As String
   Dim StrLotCodeGREF As String
   Dim StrLotContrat As String

   On Error GoTo Err_CmdTransmettre_Click

   Set WksEspace = DBEngine.Workspaces(0)    ' Cherche la zone de travail par défaut.

   Set DBBase = WksEspace.Databases(0) ' Cherche la base de données active.

   DoCmd.Hourglass True

   WksEspace.BeginTrans

   StrSQL = "SELECT DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Clé_Impayé, DEBITEUR.Déb_Situation, DEBITEUR.Déb_Date_der_rel, Round(DEBITEUR.Déb_Mont_Prin_m1,2) As Déb_Mont_Prin_m1, Round(DEBITEUR.Déb_Mont_Prin_m2,2) As Déb_Mont_Prin_m2, Round(Nz(Sum([Enc_Prin_m1]),0),2) AS Déb_Prin_Enc_m1, Round(Nz(Sum([Enc_Prin_m2]),0),2) AS Déb_Prin_Enc_m2 " & _
                  "FROM (TMP_IRRECOUVRABLE INNER JOIN DEBITEUR ON TMP_IRRECOUVRABLE.TIrr_Déb_Code = DEBITEUR.Déb_Code) LEFT JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur " & _
                  "WHERE (((TMP_IRRECOUVRABLE.TIrr_Flag) = True)) " & _
                  "GROUP BY DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Clé_Impayé, DEBITEUR.Déb_Situation, DEBITEUR.Déb_Date_der_rel, DEBITEUR.Déb_Mont_Prin_m1, DEBITEUR.Déb_Mont_Prin_m2;"

   Set RstIrrecouvrable = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstIrrecouvrableExport = DBBase.OpenRecordset("EXPORTATION_IRR")

   Set RstIrrecouvrableExportDetail = DBBase.OpenRecordset("EXPORTATION_IRR_DETAIL")

   Select Case RstIrrecouvrable.EOF
      Case True

         MsgBox "Il n'y a aucun irrecouvrable à transmettre", vbExclamation, stSource

         WksEspace.CommitTrans

      Case False

         StrLotDate = Year(Now) & IIf(Month(Now) < 10, "0", "") & Month(Now) & IIf(Day(Now) < 10, "0", "") & Day(Now)

         StrLotSite = StrGrp_Code

         StrLotCodeGREF = DLookup("Code_Sous_Traitant", "PARAMETRE_GREF")

         StrLotContrat = DLookup("Grp_Contrat", "GROUPE", "Grp_Code='" & StrGrp_Code & "'")

         StrLotNum = Format(Nz(DMax("EXPI_Lot_Num", "EXPORTATION_IRR", "EXPI_Lot_Date='" & StrLotDate & "' AND EXPI_Lot_Site='" & StrLotSite & "' AND EXPI_Lot_CodeGREF='" & StrLotCodeGREF & "' AND EXPI_Lot_Contrat='" & StrLotContrat & "'") + 1, "01"), "00")

         RstIrrecouvrableExport.AddNew

         RstIrrecouvrableExport("EXPI_Lot_Date") = StrLotDate

         RstIrrecouvrableExport("EXPI_Lot_Num") = StrLotNum

         RstIrrecouvrableExport("EXPI_Lot_Site") = StrLotSite

         RstIrrecouvrableExport("EXPI_Lot_CodeGREF") = StrLotCodeGREF

         RstIrrecouvrableExport("EXPI_Lot_Contrat") = StrLotContrat

         RstIrrecouvrableExport("EXPI_Date_Début") = IIf(IsDate(StrDateDebut), StrDateDebut, Null)

         RstIrrecouvrableExport("EXPI_Date_Fin") = IIf(IsDate(StrDateFin), StrDateFin, Null)

         RstIrrecouvrableExport("EXPI_Nb_Enreg") = 0

         RstIrrecouvrableExport("EXPI_Montant_m1") = 0

         RstIrrecouvrableExport("EXPI_Montant_m2") = 0

         LngEXPI_Id = RstIrrecouvrableExport("EXPI_Id")

         RstIrrecouvrableExport.Update

         RstIrrecouvrable.MoveLast

         RstIrrecouvrable.MoveFirst

         MajPatientez 0, 0, RstIrrecouvrable.RecordCount

         Do Until RstIrrecouvrable.EOF

            MajPatientez RstIrrecouvrable.AbsolutePosition

            RstIrrecouvrableExportDetail.AddNew

            RstIrrecouvrableExportDetail("EID_Code") = LngEXPI_Id

            RstIrrecouvrableExportDetail("EID_CodeDébiteur") = RstIrrecouvrable("Déb_Code")

            RstIrrecouvrableExportDetail("EID_NomDébiteur") = RstIrrecouvrable("Déb_Nom")

            RstIrrecouvrableExportDetail("EID_Clé_Impayé") = RstIrrecouvrable("Déb_Clé_Impayé")

            RstIrrecouvrableExportDetail("EID_Situation_Client") = DonnerSituationClient(StrGrp_Code, RstIrrecouvrable("Déb_Situation"))

            RstIrrecouvrableExportDetail("EID_Situation") = RstIrrecouvrable("Déb_Situation")

            RstIrrecouvrableExportDetail("EID_DateEditIrr") = RstIrrecouvrable("Déb_Date_der_rel")

            RstIrrecouvrableExportDetail("EID_Montant_m1") = RstIrrecouvrable("Déb_Mont_Prin_m1") - RstIrrecouvrable("Déb_Prin_Enc_m1")

            RstIrrecouvrableExportDetail("EID_Montant_m2") = RstIrrecouvrable("Déb_Mont_Prin_m2") - RstIrrecouvrable("Déb_Prin_Enc_m2")

            RstIrrecouvrableExportDetail("EID_Montant_mo") = Cli_MonnaieEnCours

            LngEID_Id = RstIrrecouvrableExportDetail("EID_Id")

            RstIrrecouvrableExport.FindFirst "EXPI_Id=" & LngEXPI_Id

            RstIrrecouvrableExport.Edit

            RstIrrecouvrableExport("EXPI_Nb_Enreg") = RstIrrecouvrableExport("EXPI_Nb_Enreg") + 1

            Select Case Cli_MonnaieEnCours
               Case 1

                  RstIrrecouvrableExport("EXPI_Montant_m1") = RstIrrecouvrableExport("EXPI_Montant_m1") + RstIrrecouvrableExportDetail("EID_Montant_m1")

               Case 2

                  RstIrrecouvrableExport("EXPI_Montant_m2") = RstIrrecouvrableExport("EXPI_Montant_m2") + RstIrrecouvrableExportDetail("EID_Montant_m2")

            End Select

            RstIrrecouvrableExport.Update

            RstIrrecouvrableExportDetail.Update

            DBBase.Execute "UPDATE DEBITEUR SET DEBITEUR.Déb_Lot_EID_Id = " & LngEID_Id & " WHERE Déb_Code='" & RstIrrecouvrable("Déb_Code") & "';", dbFailOnError

            RstIrrecouvrable.MoveNext

         Loop

         RstIrrecouvrableExport.FindFirst "EXPI_Id=" & LngEXPI_Id

         RstIrrecouvrableExport.Edit

         RstIrrecouvrableExport("EXPI_Date") = Date

         RstIrrecouvrableExport("EXPI_Date_Virement") = TxtEXPI_Date_Virement

         RstIrrecouvrableExport.Update

         RstIrrecouvrable.Close

         RstIrrecouvrableExport.Close

         RstIrrecouvrableExportDetail.Close

         DBBase.Execute "DELETE TMP_IRRECOUVRABLE.* FROM TMP_IRRECOUVRABLE;", dbFailOnError

         WksEspace.CommitTrans

         GenererFichierExportIrrecouvrable LngEXPI_Id

   End Select

   DoCmd.Close acForm, Me.Name

   DoCmd.Hourglass False

Exit_CmdTransmettre_Click:

   DoCmd.Close acForm, "FrmPatientez"

   Set RstIrrecouvrable = Nothing

   Set RstIrrecouvrableExport = Nothing

   Set RstIrrecouvrableExportDetail = Nothing

   Exit Sub

Err_CmdTransmettre_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdTransmettre_Click"

   Resume Exit_CmdTransmettre_Click
End Sub
