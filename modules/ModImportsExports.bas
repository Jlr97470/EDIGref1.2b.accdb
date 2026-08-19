Attribute VB_Name = "ModImportsExports"
Option Compare Database
Option Explicit

Public Function ExporteExcelRequete(ByVal StrSQL As String) As Boolean
   Dim AppExcel As Excel.Application
   Dim WbkClasseur As Excel.Workbook
   Dim WstFeuille As Excel.Worksheet
   Dim RsRequete As DAO.Recordset

   On Error GoTo Err_ExporteExcelRequete

   ExporteExcelRequete = True

   DoCmd.Hourglass True

   DoEvents

   Set RsRequete = CurrentDb.OpenRecordset(StrSQL)

   Select Case RsRequete.RecordCount
      Case vbEmpty

      Case Else

         On Error Resume Next

         On Error GoTo Err_ExporteExcelRequete

         DoEvents

         On Error Resume Next

         Set AppExcel = GetObject(, "Excel.Application")

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               Set AppExcel = CreateObject("Excel.Application")

         End Select

         On Error GoTo Err_ExporteExcelRequete

         Set WbkClasseur = AppExcel.Workbooks.Add

         Set WstFeuille = WbkClasseur.Worksheets(1)

         On Error Resume Next

         WstFeuille.Name = Left(RsRequete.Properties("Description").Value, 31)

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               WstFeuille.Name = "Exportation Excel"

         End Select

         WstFeuille.Rows.RowHeight = 14

         ExporteExcelRequeteDonner WstFeuille, RsRequete, 1, 1

         AppExcel.Visible = True

   End Select

   RsRequete.Close

Exit_ExporteExcelRequete:

   DoCmd.Close acForm, "FrmGestionTraitements"

   DoCmd.Hourglass False

   Set RsRequete = Nothing

   Set WstFeuille = Nothing

   Set WbkClasseur = Nothing

   Set AppExcel = Nothing

   Exit Function

Err_ExporteExcelRequete:

   ExporteExcelRequete = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ExporteExcelRequete"

   Resume Exit_ExporteExcelRequete
End Function

Public Function ExporteExcelRequeteDonner(ByRef WstFeuille As Excel.Worksheet, ByRef RsRequete As DAO.Recordset, ByVal LngPosCol As Long, ByVal LngPosLig As Long) As Boolean
   Dim FldChamp As DAO.Field
   Dim HypChamp As Excel.Hyperlinks
   Dim LngPosColChamp As Long
   Dim LngPosLigChamp As Long

   On Error GoTo Err_ExporteExcelRequeteDonner

   ExporteExcelRequeteDonner = True

   DoCmd.Hourglass True

   DoEvents

   With WstFeuille

         On Error Resume Next

         .Cells(LngPosLig, LngPosCol).Value = RsRequete.Properties("Description").Value

         Select Case Err.Number
            Case vbEmpty

            Case Else

               Err.Clear

               .Cells(LngPosLig, LngPosCol).Value = "Exportation Excel"

         End Select

         On Error GoTo Err_ExporteExcelRequeteDonner

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).MergeCells = True

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).HorizontalAlignment = -4108

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig, LngPosCol), .Cells(LngPosLig, LngPosCol + RsRequete.Fields.Count - 1)).Interior.Color = RGB(100, 100, 100)

         LngPosColChamp = vbEmpty

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Properties("Description").Value

            Select Case Err.Number
               Case vbEmpty

               Case Else

                  Err.Clear

                  .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Value = FldChamp.Name

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.LineStyle = 1

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).Borders.Weight = 3

            .Cells(LngPosLig + 1, LngPosCol + LngPosColChamp).HorizontalAlignment = -4108

            LngPosColChamp = LngPosColChamp + 1

            DoEvents

         Next

         DoEvents

         .Cells(LngPosLig + 2, LngPosCol).CopyFromRecordset RsRequete

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders.LineStyle = 1

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(8).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(9).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(7).Weight = 3

         .Range(.Cells(LngPosLig + 2, LngPosCol), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + RsRequete.Fields.Count - 1)).Borders(10).Weight = 3

         LngPosColChamp = vbEmpty

         DoEvents

         For Each FldChamp In RsRequete.Fields

            On Error Resume Next

            Select Case FldChamp.Properties("Format")
               Case vbNullString

               Case "Percent"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "0.00%"

               Case "Short Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY"

               Case "General Date"

                  .Range(.Cells(LngPosLig + 2, LngPosCol + LngPosColChamp), .Cells(LngPosLig + 2 + RsRequete.RecordCount - 1, LngPosCol + LngPosColChamp)).NumberFormat = "DD/MM/YYYY HH:MM"

               Case Else

            End Select

            On Error GoTo Err_ExporteExcelRequeteDonner

            Select Case InStr(1, FldChamp.Name, "EMail")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "mailto:" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            Select Case InStr(1, FldChamp.Name, "Web")
               Case 0

               Case Else

                  For LngPosLigChamp = 0 To RsRequete.RecordCount - 1

                     Select Case InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#")
                        Case 0

                        Case Else

                           .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value = Left(.Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, InStr(1, .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value, "#") - 1)

                     End Select

                     Set HypChamp = .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Hyperlinks

                     HypChamp.Add .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp), "http://" & .Cells(LngPosLig + 2 + LngPosLigChamp, LngPosCol + LngPosColChamp).Value

                  Next

            End Select

            .Columns(LngPosCol + LngPosColChamp).AutoFit

            LngPosColChamp = LngPosColChamp + 1

            DoEvents

         Next

         .PageSetup.Orientation = 2

         .PageSetup.Zoom = False

         .PageSetup.FitToPagesWide = 1

         .PageSetup.FitToPagesTall = 10

   End With

Exit_ExporteExcelRequeteDonner:

   DoCmd.Close acForm, "FrmGestionTraitements"

   DoCmd.Hourglass False

   Set HypChamp = Nothing

   Exit Function

Err_ExporteExcelRequeteDonner:

   ExporteExcelRequeteDonner = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ExporteExcelRequeteDonner"

   Resume Exit_ExporteExcelRequeteDonner
End Function
