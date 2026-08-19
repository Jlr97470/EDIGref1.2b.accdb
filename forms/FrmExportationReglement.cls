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

   StrSQL = "UPDATE TMP_REGLEMENT SET TMP_REGLEMENT.TReg_Flag = " & IIf(BolSelection, "True", "False") & ";"

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

Private Sub ChkTReg_Flag_AfterUpdate()

   On Error GoTo Err_ChkTReg_Flag_AfterUpdate

   DoCmd.RunCommand acCmdSaveRecord

Exit_ChkTReg_Flag_AfterUpdate:

   Exit Sub

Err_ChkTReg_Flag_AfterUpdate:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ChkTReg_Flag_AfterUpdate"

   Resume Exit_ChkTReg_Flag_AfterUpdate
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
   Dim RstReglement As DAO.Recordset
   Dim RstReglementExport As DAO.Recordset
   Dim RstReglementExportDetail As DAO.Recordset
   Dim LngEXPR_Id As Long
   Dim LngERD_Id As Long
   Dim StrSQL As String
   Dim StrDéb_Code As String
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

   StrSQL = "SELECT ENCAISSEMENT.Enc_ID, DEBITEUR.Déb_Code, DEBITEUR.Déb_Nom, DEBITEUR.Déb_Clé_Impayé, ENCAISSEMENT.Enc_Prin_m1, ENCAISSEMENT.Enc_Prin_m2 " & _
                  "FROM TMP_REGLEMENT INNER JOIN (DEBITEUR INNER JOIN ENCAISSEMENT ON DEBITEUR.Déb_Code = ENCAISSEMENT.Enc_Code_Débiteur) ON TMP_REGLEMENT.TReg_Enc_ID = ENCAISSEMENT.Enc_ID " & _
                  "WHERE TReg_Flag=True ORDER BY Déb_Code;"

   Set RstReglement = DBBase.OpenRecordset(StrSQL, dbOpenDynaset)

   Set RstReglementExport = DBBase.OpenRecordset("EXPORTATION_REG")

   Set RstReglementExportDetail = DBBase.OpenRecordset("EXPORTATION_REG_DETAIL")

   Select Case RstReglement.EOF
      Case True

         MsgBox "Il n'y a aucun règlement à transmettre", vbExclamation, stSource

         WksEspace.CommitTrans

      Case False

         StrLotDate = Year(Now) & IIf(Month(Now) < 10, "0", "") & Month(Now) & IIf(Day(Now) < 10, "0", "") & Day(Now)

         StrLotSite = StrGrp_Code

         StrLotCodeGREF = DLookup("Code_Sous_Traitant", "PARAMETRE_GREF")

         StrLotContrat = DLookup("Grp_Contrat", "GROUPE", "Grp_Code='" & StrGrp_Code & "'")

         StrLotNum = Format(Nz(DMax("EXPR_Lot_Num", "EXPORTATION_REG", "EXPR_Lot_Date='" & StrLotDate & "' AND EXPR_Lot_Site='" & StrLotSite & "' AND EXPR_Lot_CodeGREF='" & StrLotCodeGREF & "' AND EXPR_Lot_Contrat='" & StrLotContrat & "'") + 1, "01"), "00")

         RstReglementExport.AddNew

         RstReglementExport("EXPR_Lot_Date") = StrLotDate

         RstReglementExport("EXPR_Lot_Num") = StrLotNum

         RstReglementExport("EXPR_Lot_Site") = StrLotSite

         RstReglementExport("EXPR_Lot_CodeGREF") = StrLotCodeGREF

         RstReglementExport("EXPR_Lot_Contrat") = StrLotContrat

         RstReglementExport("EXPR_Date_Début") = IIf(IsDate(StrDateDebut), StrDateDebut, Null)

         RstReglementExport("EXPR_Date_Fin") = IIf(IsDate(StrDateFin), StrDateFin, Null)

         RstReglementExport("EXPR_Nb_Enreg") = 0

         RstReglementExport("EXPR_Montant_m1") = 0

         RstReglementExport("EXPR_Montant_m2") = 0

         LngEXPR_Id = RstReglementExport("EXPR_Id")

         RstReglementExport.Update

         RstReglement.MoveLast

         RstReglement.MoveFirst

         MajPatientez 0, 0, RstReglement.RecordCount

         Do Until RstReglement.EOF

            MajPatientez RstReglement.AbsolutePosition

            Select Case StrDéb_Code <> RstReglement("Déb_Code")
               Case True

                  StrDéb_Code = RstReglement("Déb_Code")

                  RstReglementExportDetail.AddNew

                  RstReglementExportDetail("ERD_Code") = LngEXPR_Id

                  RstReglementExportDetail("ERD_CodeDébiteur") = StrDéb_Code

                  RstReglementExportDetail("ERD_NomDébiteur") = RstReglement("Déb_Nom")

                  RstReglementExportDetail("ERD_Clé_Impayé") = RstReglement("Déb_Clé_Impayé")

                  RstReglementExportDetail("ERD_Montant_m1") = RstReglement("Enc_Prin_m1")

                  RstReglementExportDetail("ERD_Montant_m2") = RstReglement("Enc_Prin_m2")

                  RstReglementExportDetail("ERD_Montant_mo") = Cli_MonnaieEnCours

                  LngERD_Id = RstReglementExportDetail("ERD_Id")

                  RstReglementExportDetail.Update

                  RstReglementExport.FindFirst "EXPR_Id=" & LngEXPR_Id

                  RstReglementExport.Edit

                  RstReglementExport("EXPR_Nb_Enreg") = RstReglementExport("EXPR_Nb_Enreg") + 1

                  RstReglementExport.Update

               Case False

                  RstReglementExportDetail.FindFirst "ERD_Id=" & LngERD_Id

                  RstReglementExportDetail.Edit

                  RstReglementExportDetail("ERD_Montant_m1") = RstReglementExportDetail("ERD_Montant_m1") + RstReglement("Enc_Prin_m1")

                  RstReglementExportDetail("ERD_Montant_m2") = RstReglementExportDetail("ERD_Montant_m2") + RstReglement("Enc_Prin_m2")

                  RstReglementExportDetail.Update

            End Select

            RstReglementExport.FindFirst "EXPR_Id=" & LngEXPR_Id

            RstReglementExport.Edit

            Select Case Cli_MonnaieEnCours
               Case 1

                  RstReglementExport("EXPR_Montant_m1") = RstReglementExport("EXPR_Montant_m1") + RstReglement("Enc_Prin_m1")

               Case 2

                  RstReglementExport("EXPR_Montant_m2") = RstReglementExport("EXPR_Montant_m2") + RstReglement("Enc_Prin_m2")

            End Select

            RstReglementExport.Update

            DBBase.Execute "UPDATE ENCAISSEMENT SET ENCAISSEMENT.Enc_ERD_Id = " & LngERD_Id & " WHERE Enc_ID=" & RstReglement("Enc_Id") & ";", dbFailOnError

            RstReglement.MoveNext

         Loop

         RstReglementExport.FindFirst "EXPR_Id=" & LngEXPR_Id

         RstReglementExport.Edit

         RstReglementExport("EXPR_Date") = Date

         RstReglementExport("EXPR_Date_Virement") = TxtEXPR_Date_Virement

         RstReglementExport.Update

         RstReglement.Close

         RstReglementExport.Close

         RstReglementExportDetail.Close

         DBBase.Execute "DELETE TMP_REGLEMENT.* FROM TMP_REGLEMENT;", dbFailOnError

         WksEspace.CommitTrans

         GenererFichierExportReglement LngEXPR_Id

   End Select

   DoCmd.Close acForm, Me.Name

   DoCmd.Hourglass False

Exit_CmdTransmettre_Click:

   DoCmd.Close acForm, "FrmPatientez"

   Set RstReglement = Nothing

   Set RstReglementExport = Nothing

   Set RstReglementExportDetail = Nothing

   Exit Sub

Err_CmdTransmettre_Click:

   WksEspace.Rollback

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdTransmettre_Click"

   Resume Exit_CmdTransmettre_Click
End Sub
