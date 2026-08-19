Attribute VB_Name = "ModFormulaire"
Option Compare Database
Option Explicit

Public Const VUECREATION = 0
Public Const VUEFORMULAIRE = 1
Public Const VUEFEUILLEDONNER = 2
Public Const VUEFERMER = 3

Public Function DonneEtatFormulaire(ByVal StrFormName As String, ByRef IntFormEtat As Integer) As Boolean
   Dim FrmDataBase As Access.Forms
   Dim FrmForm As Access.Form

   On Error GoTo Exit_DonneEtatFormulaire

   Set FrmDataBase = Forms

   DonneEtatFormulaire = False

   For Each FrmForm In FrmDataBase

      Select Case FrmForm.Name
         Case StrFormName

            IntFormEtat = FrmForm.CurrentView

            DonneEtatFormulaire = True

            GoTo Exit_DonneEtatFormulaire

         Case Else

      End Select
   Next

   IntFormEtat = VUEFERMER

   DonneEtatFormulaire = False

Exit_DonneEtatFormulaire:

   Exit Function

Err_DonneEtatFormulaire:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonneEtatFormulaire"

   Resume Exit_DonneEtatFormulaire
End Function

Public Function DonneEtatRaport(ByVal StrReportName As String, ByRef IntReportEtat As Integer) As Boolean
   Dim RptDataBase As Access.Reports
   Dim RptReport As Access.Report

   On Error GoTo Exit_DonneEtatRaport

   Set RptDataBase = Forms

   DonneEtatRaport = False

   For Each RptReport In RptDataBase

      Select Case RptReport.Name
         Case StrReportName

            IntReportEtat = RptReport.CurrentView

            DonneEtatRaport = True

            GoTo Exit_DonneEtatRaport

         Case Else

      End Select
   Next

   IntReportEtat = VUEFERMER

   DonneEtatRaport = False

Exit_DonneEtatRaport:

   Exit Function

Err_DonneEtatRaport:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : DonneEtatRaport"

   Resume Exit_DonneEtatRaport
End Function

Sub MajPatientez(Optional IntValue, Optional IntMin, Optional IntMax)

   On Error GoTo Err_MajPatientez

   DoCmd.OpenForm "FrmPatientez", acNormal, , , acFormEdit, acWindowNormal

   Forms("FrmPatientez").PgrPatientez.Visible = True

   Select Case IsMissing(IntMin)
      Case True

      Case False

         Forms("FrmPatientez").PgrPatientez.Min = IntMin

   End Select

   Select Case IsMissing(IntMax)
      Case True

      Case False

         Forms("FrmPatientez").PgrPatientez.Max = IntMax

   End Select

   Select Case IsMissing(IntValue)
      Case True

      Case False

         Forms("FrmPatientez").PgrPatientez.Value = IntValue

   End Select

   Forms("FrmPatientez").TxtPatientez = Forms("FrmPatientez").PgrPatientez.Value & "/" & Forms("FrmPatientez").PgrPatientez.Max

   DoEvents

Exit_MajPatientez:

   Exit Sub

Err_MajPatientez:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : MajPatientez"

   Resume Exit_MajPatientez
End Sub
