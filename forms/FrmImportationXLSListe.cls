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

   DoCmd.OpenReport "EtaCompteRenduImportXLS", acDesign

   Reports("EtaCompteRenduImportXLS").Filter = "IMPX_Id=" & IMPX_Id

   Reports("EtaCompteRenduImportXLS").FilterOn = True

   DoCmd.Close acReport, "EtaCompteRenduImportXLS", acSaveYes

   Select Case IntVisu
      Case 0, 2

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduImportXLS", IntVisu, , "IMPX_Id=" & IMPX_Id & " AND Déb_MonnaieEnCours=" & cFRANC, acWindowNormal, cFRANC & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCompteRenduImportXLS", IntRaportEtat) = True

            DoEvents

         Loop

         On Error Resume Next

         DoCmd.OpenReport "EtaCompteRenduImportXLS", IntVisu, , "IMPX_Id=" & IMPX_Id & " AND Déb_MonnaieEnCours=" & cEURO, acWindowNormal, cEURO & ";"

         Select Case Err.Number
            Case 0

            Case 2501

               Err.Clear

            Case Else

               GoTo Err_ImpVisu

         End Select

         On Error GoTo Err_ImpVisu

         Do While DonneEtatRaport("EtaCompteRenduImportXLS", IntRaportEtat) = True

            DoEvents

         Loop

      Case 5

         IntRepMsg = MsgBox("Souhaitez-vous exporter les données au format Excel ?" & vbCrLf & "Indiquez OUI pour extraire les données au format excel et NON pour le format texte.", vbYesNo + vbQuestion, "Format d'exportation")

         Select Case IntRepMsg
            Case vbYes

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduImportXLS", acFormatXLS, , False

            Case vbNo

               DoCmd.OutputTo acOutputReport, "EtaCompteRenduImportXLS", acFormatRTF, , False

         End Select

   End Select

Exit_ImpVisu:

    Exit Function

Err_ImpVisu:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ImpVisu"

    Resume Exit_ImpVisu
End Function

Private Sub CmdFermer_Click()

   On Error GoTo Err_CmdFermer_Click

   DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

   Exit Sub

Err_CmdFermer_Click:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

   Resume Exit_CmdFermer_Click
End Sub
