Attribute VB_Name = "ModBase"
Option Compare Database
Option Explicit

Public Function VerifierBase() As Boolean
   Dim FsoSystemFichier As New Scripting.FileSystemObject
   Dim StrPathEDIMDB As String
   Dim StrFichEDIMDB As String
   Dim StrPathMDB As String
   Dim StrFichMDB As String
   Dim StrFichArchMDB As String
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_VerifierBase

   VerifierBase = True

   CurrentDb.Properties("AppTitle") = Nz(DLookup("Bas_Nom", "BASE"))

   StrPathMDB = Nz(DLookup("Bas_Path", "BASE"))

   StrFichMDB = Nz(DLookup("Bas_MDB", "BASE"))

   StrFichArchMDB = DLookup("Bas_Archive_MDB", "BASE")

   StrPathEDIMDB = Nz(DLookup("Bas_Path_EDI", "BASE"))

   StrFichEDIMDB = Nz(DLookup("Bas_MDB_EDI", "BASE"))

   Select Case FsoSystemFichier.FileExists(StrPathMDB & StrFichMDB)
      Case True

      Case False

         VerifierBase = False

         MsgBox "Base " & StrPathMDB & StrFichMDB & " Inexistante !!"

         GoTo Exit_VerifierBase

   End Select

   Select Case FsoSystemFichier.FileExists(StrPathMDB & StrFichArchMDB)
      Case True

      Case False

         VerifierBase = False

         MsgBox "Base " & StrPathMDB & StrFichArchMDB & " Inexistante !!"

         GoTo Exit_VerifierBase

   End Select

   Select Case FsoSystemFichier.FileExists(StrPathEDIMDB & StrFichEDIMDB)
      Case True

      Case False

         VerifierBase = False

         MsgBox "Base " & StrPathEDIMDB & StrFichEDIMDB & " Inexistante !!"

         GoTo Exit_VerifierBase

   End Select

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Connect
         Case ""

         Case Else

            Select Case DLookup("TblNumero", "TABLE", "TblNom='" & TdfTable.Name & "'")
               Case 1

                  SauveTableInformationConnexion TdfTable, "", "", StrPathMDB, StrFichMDB

               Case 2

                  SauveTableInformationConnexion TdfTable, "", "", StrPathMDB, StrFichArchMDB

               Case 3

                  SauveTableInformationConnexion TdfTable, "", "", StrPathEDIMDB, StrFichEDIMDB

            End Select

      End Select

   Next

   CurrentDb.TableDefs.Refresh

   'MajBDD12

Exit_VerifierBase:

   Exit Function

Err_VerifierBase:

   VerifierBase = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : VerifierBase"

   Resume Exit_VerifierBase
End Function

Public Function LectureTableInformationConnexion(ByVal TdfTable As TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim StrConnexion As String
   Dim IntConnexionLongueur As Integer
   Dim StrCheminFichier As String
   Dim IntCheminFichierLongueur As Integer

   On Error GoTo Err_LectureTableInformationConnexion

   StrConnexion = TdfTable.Connect

   IntConnexionLongueur = Len(StrConnexion)

   Select Case IntConnexionLongueur
      Case vbEmpty

         StrType = vbNullString

         StrParametre = vbNullString

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

         GoTo Exit_LectureTableInformationConnexion

      Case Else

   End Select

   StrType = Mid(StrConnexion, 1, InStr(1, StrConnexion, ";") - 1)

   StrCheminFichier = Mid(StrConnexion, InStr(1, StrConnexion, "DATABASE=") + 9)

   StrParametre = Replace(Replace(StrConnexion, StrType & ";", vbNullString), "DATABASE=" & StrCheminFichier, vbNullString)

   IntCheminFichierLongueur = Len(StrCheminFichier)

   Select Case IntCheminFichierLongueur
      Case vbEmpty

         StrChemin = vbNullString

         StrFichier = vbNullString

         LectureTableInformationConnexion = False

      Case Else

         Select Case StrType
            Case vbNullString, "Excel 5.0"

               StrChemin = Mid(StrCheminFichier, 1, InStrRev(StrCheminFichier, "\") - 1)

               StrFichier = Mid(StrCheminFichier, InStrRev(StrCheminFichier, "\") + 1)

            Case "Text"

               StrChemin = StrCheminFichier

               StrFichier = TdfTable.SourceTableName

            Case Else

         End Select

         LectureTableInformationConnexion = True

   End Select

Exit_LectureTableInformationConnexion:

   Exit Function

Err_LectureTableInformationConnexion:

   LectureTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.description, , stSource & " : LectureTableInformationConnexion"

   Resume Exit_LectureTableInformationConnexion
End Function

Public Function SauveTableInformationConnexion(ByRef TdfTable As DAO.TableDef, ByRef StrType As String, ByRef StrParametre, ByRef StrChemin As String, ByRef StrFichier As String) As Boolean
   Dim strConnect As String

   On Error GoTo Err_SauveTableInformationConnexion

   SauveTableInformationConnexion = True

   Select Case StrType
      Case vbNullString

         strConnect = ";DATABASE=" & StrChemin & StrFichier

      Case "Excel 5.0"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin & StrFichier

      Case "Text"

         strConnect = StrType & ";" & StrParametre & ";DATABASE=" & StrChemin

      Case Else

   End Select

   Select Case TdfTable.Connect
      Case strConnect

      Case Else

         TdfTable.Connect = strConnect

         TdfTable.RefreshLink

   End Select

Exit_SauveTableInformationConnexion:

   Exit Function

Err_SauveTableInformationConnexion:

   SauveTableInformationConnexion = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : SauveTableInformationConnexion"

   Resume Exit_SauveTableInformationConnexion
End Function

Public Function AttacheFichierBlitzCash(ByVal StrPatch As String, ByVal StrLot As String) As Boolean
   Dim TdfTable As DAO.TableDef

   On Error GoTo Err_AttacheFichierBlitzCash

   For Each TdfTable In CurrentDb.TableDefs

      Select Case TdfTable.Name
         Case "BLITZ_IMP", "BLITZ_DEB", "BLITZ_ADR", "BLITZ_INT"

            CurrentDb.TableDefs.Delete TdfTable.Name

         Case Else

      End Select

   Next

   DoCmd.TransferText acImportDelim, "FORMAT BLITZ IMP", "BLITZ_IMP", StrPatch & StrLot & "IMP.TXT", False

   DoCmd.TransferText acImportDelim, "FORMAT BLITZ DEB", "BLITZ_DEB", StrPatch & StrLot & "DEB.TXT", False

   DoCmd.TransferText acImportDelim, "FORMAT BLITZ ADR", "BLITZ_ADR", StrPatch & StrLot & "ADR.TXT", False

   DoCmd.TransferText acImportDelim, "FORMAT BLITZ INT", "BLITZ_INT", StrPatch & StrLot & "INT.TXT", False

Exit_AttacheFichierBlitzCash:

   Exit Function

Err_AttacheFichierBlitzCash:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : AttacheFichierBlitzCash"

   Resume Exit_AttacheFichierBlitzCash
End Function

Public Function MajBDD12() As Boolean
   Dim FsoFichierSystem As New Scripting.FileSystemObject
   Dim FilFichier As Scripting.File
   Dim DBBaseSource As DAO.Database
   Dim DBBaseSourceMDB As DAO.Database
   Dim DBBaseDestination As DAO.Database
   Dim DBBaseDestinationMDB As DAO.Database
   Dim TdfTableSource As DAO.TableDef
   Dim TdfTableDestination As DAO.TableDef
   Dim RstTableSource As DAO.Recordset
   Dim RstTableDestination As DAO.Recordset
   Dim RstTable As DAO.Recordset
   Dim FldChamp As DAO.Field
   Dim StrSQL As String
   Dim StrPath As String
   Dim StrPathMDB As String
   Dim StrFichierSource As String
   Dim StrFichierSourceMDB As String
   Dim StrFichierDestination As String
   Dim StrFichierDestinationMDB As String
   Dim LngIMPD_Id As Long
   Dim LngERD_Id As Long
   Dim LngEID_Id As Long
   Dim LngERD_Code As Long
   Dim StrERD_CodeDébiteur As String
   Dim IntFichierNumero As Integer

   On Error GoTo Err_MajBDD12

   MajBDD12 = True

   StrPath = DLookup("Bas_Path_EDI", "BASE")

   StrPathMDB = DLookup("Bas_Path", "BASE")

   StrFichierSource = "TEDIGREF2003OLD.MDB"

   StrFichierSourceMDB = "TGREF2003OLD.MDB"

   StrFichierDestination = DLookup("Bas_MDB_EDI", "BASE")

   StrFichierDestinationMDB = DLookup("Bas_MDB", "BASE")

   Select Case FsoFichierSystem.FileExists(StrPath & StrFichierSource)
      Case True

      Case False

         GoTo Exit_MajBDD12

   End Select

   Select Case FsoFichierSystem.FileExists(StrPathMDB & StrFichierSourceMDB)
      Case True

      Case False

         GoTo Exit_MajBDD12

   End Select

   IntFichierNumero = FreeFile

   Open StrPath & "MajBDD.LOG" For Output As #IntFichierNumero

   Print #IntFichierNumero, "Maj 1.2 : " & Now()

   Set DBBaseSource = DBEngine.OpenDatabase(StrPath & StrFichierSource)

   Set DBBaseSourceMDB = DBEngine.OpenDatabase(StrPathMDB & StrFichierSourceMDB)

   Set DBBaseDestination = DBEngine.OpenDatabase(StrPath & StrFichierDestination)

   Set DBBaseDestinationMDB = DBEngine.OpenDatabase(StrPathMDB & StrFichierDestinationMDB)

   Set TdfTableSource = DBBaseSource.TableDefs("IMPORTATION")

   Set TdfTableDestination = DBBaseDestination.TableDefs("IMPORTATION")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "IMP_Montant_mo"

            Case "IMP_Montant_m1"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(IMPORTATION_DETAIL.IMPD_Montant),0)  AS IMPD_Montant FROM IMPORTATION_DETAIL WHERE IMPD_Code=" & RstTableSource("IMP_id") & " AND IMPD_Monnaie=1;")

               RstTableDestination("IMP_Montant_m1") = RstTable("IMPD_Montant")

               RstTable.Close

            Case "IMP_Montant_m2"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(IMPORTATION_DETAIL.IMPD_Montant),0)  AS IMPD_Montant FROM IMPORTATION_DETAIL WHERE IMPD_Code=" & RstTableSource("IMP_id") & " AND IMPD_Monnaie=2;")

               RstTableDestination("IMP_Montant_m2") = RstTable("IMPD_Montant")

               RstTable.Close

            Case "IMP_Date"

               RstTableDestination(FldChamp.Name) = IIf(IsNull(FldChamp.Value), Null, Format(FldChamp.Value, "Short Date"))

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " IMPORTATION: " & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("IMPORTATION_DETAIL")

   Set TdfTableDestination = DBBaseDestination.TableDefs("IMPORTATION_DETAIL")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      RstTableDestination("IMPD_Code") = RstTableSource("IMPD_Code")

      RstTableDestination("IMPD_CodeClient") = RstTableSource("IMPD_CodeClient")

      RstTableDestination("IMPD_CodeEtab") = RstTableSource("IMPD_CodeEtab")

      RstTableDestination("IMPD_NomEtab") = RstTableSource("IMPD_NomEtab")

      RstTableDestination("IMPD_CodeGroupe") = RstTableSource("IMPD_CodeGroupe")

      RstTableDestination("IMPD_NbCréances") = RstTableSource("IMPD_NbCréances")

      RstTableDestination("IMPD_Montant_m" & RstTableSource("IMPD_Monnaie")) = RstTableSource("IMPD_Montant")

      RstTableDestination("IMPD_Montant_mo") = RstTableSource("IMPD_Monnaie")

      Select Case RstTableSource("IMPD_Monnaie")
         Case cFRANC

            RstTableDestination("IMPD_Montant_m2") = RstTableDestination("IMPD_Montant_m1") * DonnerTauxConvertion(cFRANC)

         Case cEURO

            RstTableDestination("IMPD_Montant_m1") = RstTableDestination("IMPD_Montant_m2") * DonnerTauxConvertion(cEURO)

      End Select

      LngIMPD_Id = RstTableDestination("IMPD_Id")

      RstTableDestination.Update

      Set RstTable = DBBaseSourceMDB.OpenRecordset("SELECT Déb_Code FROM DEBITEUR WHERE Déb_Lot_Imp_Id=" & RstTableSource("IMPD_Code") & " AND Déb_Code_Client='" & RstTableSource("IMPD_CodeClient") & "';")

      Do While Not RstTable.EOF

         DBBaseDestinationMDB.Execute "UPDATE DEBITEUR SET DEBITEUR.Déb_Lot_IMPD_Id = " & LngIMPD_Id & " WHERE DEBITEUR.Déb_Code='" & RstTable("Déb_Code") & "';", dbFailOnError

         RstTable.MoveNext

      Loop

      RstTable.Close

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " IMPORTATION_DETAIL:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("EXPORTATION_REG")

   Set TdfTableDestination = DBBaseDestination.TableDefs("EXPORTATION_REG")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "EXPR_Montant_mo", "EXPR_Flag1"

            Case "EXPR_Montant_m1"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(EXPORTATION_REG_DETAIL.ERD_Montant),0)  AS ERD_Montant FROM EXPORTATION_REG_DETAIL WHERE ERD_Code=" & RstTableSource("EXPR_Id") & " AND ERD_Monnaie=1;")

               RstTableDestination("EXPR_Montant_m1") = RstTable("ERD_Montant")

               RstTable.Close

            Case "EXPR_Montant_m2"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(EXPORTATION_REG_DETAIL.ERD_Montant),0)  AS ERD_Montant FROM EXPORTATION_REG_DETAIL WHERE ERD_Code=" & RstTableSource("EXPR_Id") & " AND ERD_Monnaie=2;")

               RstTableDestination("EXPR_Montant_m2") = RstTable("ERD_Montant")

               RstTable.Close

            Case "EXPR_Date"

               RstTableDestination(FldChamp.Name) = IIf(IsNull(FldChamp.Value), Null, Format(FldChamp.Value, "Short Date"))

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " EXPORTATION_REG:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrSQL = "SELECT * FROM EXPORTATION_REG_DETAIL ORDER BY ERD_Code,ERD_CodeDébiteur;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("EXPORTATION_REG_DETAIL")

   Set RstTableDestination = TdfTableDestination.OpenRecordset(dbOpenDynaset)

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      Select Case RstTableSource("ERD_Code") <> LngERD_Code Or RstTableSource("ERD_CodeDébiteur") <> StrERD_CodeDébiteur
         Case True

            LngERD_Code = RstTableSource("ERD_Code")

            StrERD_CodeDébiteur = RstTableSource("ERD_CodeDébiteur")

            RstTableDestination.AddNew

            RstTableDestination("ERD_Code") = LngERD_Code

            RstTableDestination("ERD_CodeDébiteur") = StrERD_CodeDébiteur

            RstTableDestination("ERD_NomDébiteur") = RstTableSource("ERD_NomDébiteur")

            RstTableDestination("ERD_Clé_Impayé") = RstTableSource("ERD_Clé_Impayé")

            RstTableDestination("ERD_Montant_mo") = RstTableSource("ERD_Monnaie")

            RstTableDestination("ERD_Montant_m" & RstTableSource("ERD_Monnaie")) = RstTableSource("ERD_Montant")

            Select Case RstTableSource("ERD_Monnaie")
               Case cFRANC

                  RstTableDestination("ERD_Montant_m2") = RstTableDestination("ERD_Montant_m1") * DonnerTauxConvertion(cFRANC)

               Case cEURO

                  RstTableDestination("ERD_Montant_m1") = RstTableDestination("ERD_Montant_m2") * DonnerTauxConvertion(cEURO)

            End Select

            LngERD_Id = RstTableDestination("ERD_Id")

            RstTableDestination.Update

         Case False

            RstTableDestination.FindFirst "ERD_Id=" & LngERD_Id

            RstTableDestination.Edit

            RstTableDestination("ERD_Montant_m" & RstTableSource("ERD_Monnaie")) = RstTableDestination("ERD_Montant_m" & RstTableSource("ERD_Monnaie")) + RstTableSource("ERD_Montant")

            Select Case RstTableSource("ERD_Monnaie")
               Case cFRANC

                  RstTableDestination("ERD_Montant_m2") = RstTableDestination("ERD_Montant_m1") * DonnerTauxConvertion(cFRANC)

               Case cEURO

                  RstTableDestination("ERD_Montant_m1") = RstTableDestination("ERD_Montant_m2") * DonnerTauxConvertion(cEURO)

            End Select

            RstTableDestination.Update

      End Select

      DBBaseDestinationMDB.Execute "UPDATE ENCAISSEMENT SET ENCAISSEMENT.Enc_ERD_Id = " & LngERD_Id & " WHERE Enc_Id=" & RstTableSource("ERD_Enc_Id") & ";", dbFailOnError

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " EXPORTATION_REG_DETAIL:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   Set TdfTableSource = DBBaseSource.TableDefs("EXPORTATION_IRR")

   Set TdfTableDestination = DBBaseDestination.TableDefs("EXPORTATION_IRR")

   Set RstTableSource = TdfTableSource.OpenRecordset(dbOpenDynaset)

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      For Each FldChamp In RstTableSource.Fields

         Select Case FldChamp.Name
            Case "EXPI_Montant_mo", "EXPI_Flag1"

            Case "EXPI_Montant_m1"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(EXPORTATION_IRR_DETAIL.EID_Montant),0)  AS EID_Montant FROM EXPORTATION_IRR_DETAIL WHERE EID_Code=" & RstTableSource("EXPI_Id") & " AND EID_Monnaie=1;")

               RstTableDestination("EXPI_Montant_m1") = RstTable("EID_Montant")

               RstTable.Close

            Case "EXPI_Montant_m2"

               Set RstTable = DBBaseSource.OpenRecordset("SELECT nz(Sum(EXPORTATION_IRR_DETAIL.EID_Montant),0)  AS EID_Montant FROM EXPORTATION_IRR_DETAIL WHERE EID_Code=" & RstTableSource("EXPI_Id") & " AND EID_Monnaie=2;")

               RstTableDestination("EXPI_Montant_m2") = RstTable("EID_Montant")

               RstTable.Close

            Case "EXPI_Date"

               RstTableDestination(FldChamp.Name) = IIf(IsNull(FldChamp.Value), Null, Format(FldChamp.Value, "Short Date"))

            Case Else

               RstTableDestination(FldChamp.Name) = FldChamp.Value

         End Select

      Next

      RstTableDestination.Update

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " EXPORTATION_IRR:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close

   StrSQL = "SELECT * FROM EXPORTATION_IRR_DETAIL ORDER BY EID_Code, EID_CodeDébiteur;"

   Set RstTableSource = DBBaseSource.OpenRecordset(StrSQL, dbOpenDynaset)

   Set TdfTableDestination = DBBaseDestination.TableDefs("EXPORTATION_IRR_DETAIL")

   Set RstTableDestination = TdfTableDestination.OpenRecordset()

   RstTableSource.MoveLast

   RstTableSource.MoveFirst

   MajPatientez 0, 0, RstTableSource.RecordCount

   Do While Not RstTableSource.EOF

      MajPatientez RstTableSource.AbsolutePosition

      RstTableDestination.AddNew

      RstTableDestination("EID_Code") = RstTableSource("EID_Code")

      RstTableDestination("EID_CodeDébiteur") = RstTableSource("EID_CodeDébiteur")

      RstTableDestination("EID_NomDébiteur") = RstTableSource("EID_NomDébiteur")

      RstTableDestination("EID_Clé_Impayé") = RstTableSource("EID_Clé_Impayé")

      RstTableDestination("EID_Situation") = RstTableSource("EID_Situation")

      RstTableDestination("EID_Situation_Client") = RstTableSource("EID_Situation_Client")

      RstTableDestination("EID_DateEditIrr") = RstTableSource("EID_DateEditIrr")

      RstTableDestination("EID_Montant_mo") = RstTableSource("EID_Monnaie")

      RstTableDestination("EID_Montant_m" & RstTableSource("EID_Monnaie")) = RstTableSource("EID_Montant")

      Select Case RstTableSource("EID_Monnaie")
         Case cFRANC

            RstTableDestination("EID_Montant_m2") = RstTableDestination("EID_Montant_m1") * DonnerTauxConvertion(cFRANC)

         Case cEURO

            RstTableDestination("EID_Montant_m1") = RstTableDestination("EID_Montant_m2") * DonnerTauxConvertion(cEURO)

      End Select

      LngEID_Id = RstTableDestination("EID_Id")

      RstTableDestination.Update

      DBBaseDestinationMDB.Execute "UPDATE DEBITEUR SET DEBITEUR.Déb_Lot_EID_Id = " & LngEID_Id & " WHERE Déb_Code='" & RstTableSource("EID_CodeDébiteur") & "';", dbFailOnError

      RstTableSource.MoveNext

   Loop

   Print #IntFichierNumero, Format(Time(), "Long Time") & " EXPORTATION_IRR_DETAIL:" & RstTableSource.RecordCount & "=> " & RstTableDestination.RecordCount

   RstTableSource.Close

   RstTableDestination.Close





   DBBaseSource.Close

   DBBaseDestination.Close

Exit_MajBDD12:

   DoCmd.Close acForm, "FrmPatientez"

   Close #IntFichierNumero

   Set RstTable = Nothing

   Set RstTableSource = Nothing

   Set RstTableDestination = Nothing

   Set TdfTableSource = Nothing

   Set TdfTableDestination = Nothing

   Set DBBaseSource = Nothing

   Set DBBaseDestination = Nothing

   Set FilFichier = Nothing

   Set FsoFichierSystem = Nothing

   Exit Function

Err_MajBDD12:

   MajBDD12 = False

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : MajBDD12"

   Resume Exit_MajBDD12
End Function
