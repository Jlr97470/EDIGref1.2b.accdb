Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Private Function ImpVisu(ByVal IntVisu As Integer)
   Dim IntRepMsg As Integer
   Dim IntRaportEtat As Integer

   On Error GoTo Err_ImpVisu

   DoCmd.OpenReport "EtaCompteRenduExportIrrecouvrable", acDesign

   Reports("EtaCompteRenduExportIrrecouvrable").Filter = "EXPI_Id=" & Me.EXPI_Id

   Reports("EtaCompteRenduExportIrrecouvrable").FilterOn = True

   DoCmd.Close acReport, "EtaCompteRenduExportIrrecouvrable", acSaveYes

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduExportIrrecouvrable", IntVisu, , "EXPI_Id=" & EXPI_Id & " AND EID_Montant_mo=" & cFRANC, acWindowNormal, cFRANC & ";"

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

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduExportIrrecouvrable", IntVisu, , "EXPI_Id=" & EXPI_Id & " AND EID_Montant_mo=" & cEURO, acWindowNormal, cEURO & ";"

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

Private Sub CmdDetail_Click()

   On Error GoTo Err_CmdDetail_Click

   DoCmd.OpenForm "FrmExportationIrrecouvrableArchive", acNormal, , "EID_Code=" & EXPI_Id, , acDialog

Exit_CmdDetail_Click:

   Exit Sub

Err_CmdDetail_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdDetail_Click"

   Resume Exit_CmdDetail_Click
End Sub
