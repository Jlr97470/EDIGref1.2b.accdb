Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private iMonnaieVisu As Integer

Private Function Inverseur(ByVal IntMonnaie As Integer)

   On Error GoTo Err_Inverseur

   Select Case IntMonnaie
      Case cFRANC

         TxtERD_Montant_m.ControlSource = "=Round(ERD_Montant_m1,2)"

         TxtERD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtERD_Montant_mMonnaie.ControlSource = "=TxtERD_Montant_m*DonnerTauxConvertion(1)"

         TxtERD_Montant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalERD_Montant_m.ControlSource = "=Sum(Round(ERD_Montant_m1,2))"

         TxtTotalERD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalERD_Montant_mMonnaie.ControlSource = "=TxtTotalERD_Montant_m*DonnerTauxConvertion(1)"

         TxtTotalERD_Montant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

         TxtTotalGeneralERD_Montant_m.ControlSource = "=Sum(Round(Nz(ERD_Montant_m1,0),2))"

         TxtTotalGeneralERD_Montant_m.Format = "#,##0.00""" & " F" & """;-#,##0.00""" & " F" & """"

         TxtTotalGeneralERD_Montant_mMonnaie.ControlSource = "=TxtTotalGeneralERD_Montant_m*DonnerTauxConvertion(1)"

         TxtTotalGeneralERD_Montant_mMonnaie.Format = "(#,##0.00""" & " €" & """);(-#,##0.00""" & " €" & """" & ")"

      Case cEURO

         TxtERD_Montant_m.ControlSource = "=Round(ERD_Montant_m2,2)"

         TxtERD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtERD_Montant_mMonnaie.ControlSource = "=TxtERD_Montant_m*DonnerTauxConvertion(2)"

         TxtERD_Montant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalERD_Montant_m.ControlSource = "=Sum(Round(ERD_Montant_m2,2))"

         TxtTotalERD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalERD_Montant_mMonnaie.ControlSource = "=TxtTotalERD_Montant_m*DonnerTauxConvertion(2)"

         TxtTotalERD_Montant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtTotalGeneralERD_Montant_m.ControlSource = "=Sum(Round(Nz(ERD_Montant_m2,0),2))"

         TxtTotalGeneralERD_Montant_m.Format = "#,##0.00""" & " €" & """;-#,##0.00""" & " €" & """"

         TxtTotalGeneralERD_Montant_mMonnaie.ControlSource = "=TxtTotalGeneralERD_Montant_m*DonnerTauxConvertion(2)"

         TxtTotalGeneralERD_Montant_mMonnaie.Format = "(#,##0.00""" & " F" & """);(-#,##0.00""" & " F" & """" & ")"

         TxtERD_Montant_mMonnaie.Visible = False

         TxtTotalERD_Montant_mMonnaie.Visible = False

         TxtTotalGeneralERD_Montant_mMonnaie.Visible = False

         EtiERD_Montant_mMonnaie.Visible = False

         EtiTotalERD_Montant_mMonnaie.Visible = False

         EtiTotalGeneralERD_Montant_mMonnaie.Visible = False

   End Select

   iMonnaieVisu = IntMonnaie

Exit_Inverseur:

   Exit Function

Err_Inverseur:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Inverseur"

   Resume Exit_Inverseur
End Function

Private Sub Report_Open(Cancel As Integer)
   Dim StrParam() As String

   On Error GoTo Err_Report_Open

   StrParam = Split(Nz(Me.OpenArgs), ";")

   Select Case UBound(StrParam)
      Case -1

         Inverseur cEURO

      Case Else

         Inverseur Nz(StrParam(0), cEURO)

   End Select

Exit_Report_Open:

   Exit Sub

Err_Report_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_Open"

   Resume Exit_Report_Open
End Sub

Private Sub Report_NoData(Cancel As Integer)

   On Error GoTo Err_Report_NoData

   Cancel = True

'   MsgBox "Aucun lot selectionner en " & DLookup("Mon_Nom", "MONNAIE", "Mon_Id=" & iMonnaieVisu) & " !!", vbCritical

Exit_Report_NoData:

   Exit Sub

Err_Report_NoData:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_NoData"

   Resume Exit_Report_NoData
End Sub
