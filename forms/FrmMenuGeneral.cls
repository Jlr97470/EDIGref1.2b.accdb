Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CreerOption()
    Const conNombreBoutons = 10
    Dim DBBase As DAO.Database
    Dim RstMenu As DAO.Recordset
    Dim StrSQL As String
    Dim intOption As Integer

    On Error GoTo Err_CreerOption

    Me![CmdOption1].SetFocus

    For intOption = 2 To conNombreBoutons

        Me("CmdOption" & intOption).Visible = False

        Me("EtiOption" & intOption).Visible = False

    Next intOption

    Set DBBase = CurrentDb()

    StrSQL = "SELECT * FROM MENU WHERE Men_Numero> 0 AND Men_Page=" & Me![Men_Page] & " ORDER BY [Men_Numero];"

    Set RstMenu = DBBase.OpenRecordset(StrSQL)

    Select Case RstMenu.EOF
      Case True

        Me![EtiOption1].Caption = "Il n'y a pas d'éléments pour cette page de Menu Général"

      Case False

        While (Not (RstMenu.EOF))

            Me("CmdOption" & RstMenu![Men_Numero]).Visible = True

            Me("EtiOption" & RstMenu![Men_Numero]).Visible = True

            Me("EtiOption" & RstMenu![Men_Numero]).Caption = RstMenu![Men_Libelle]

            RstMenu.MoveNext

        Wend

    End Select

    RstMenu.Close

    DBBase.Close

Exit_CreerOption:

   Set RstMenu = Nothing

   Set DBBase = Nothing

   Exit Sub

Err_CreerOption:

   MsgBox Err.Number & " " & Err.description

   Resume Exit_CreerOption
End Sub

Private Function ButtonClick(intBtn As Integer)
    Const conCmdAllerMenuGénéral = 1
    Const conCmdOuvrirFormulaireAjouter = 2
    Const conCmdOuvrirFormulaireParcourir = 3
    Const conCmdOuvrirEtat = 4
    Const conCmdPersonnaliserMenuGénéral = 5
    Const conCmdQuitterApplication = 6
    Const conCmdExécuterMacro = 7
    Const conCmdExécuterCode = 8
    Const conErrDoCmdAnnulé = 2501    ' Cas particulier d'erreur.
    Dim DBBase As DAO.Database
    Dim RstMenu As DAO.Recordset

   On Error GoTo Err_ButtonClick

   Set DBBase = CurrentDb()

   Set RstMenu = DBBase.OpenRecordset("MENU", dbOpenDynaset)

   RstMenu.FindFirst "[Men_Page]=" & Me![Men_Page] & " AND [Men_Numero]=" & intBtn

   Select Case RstMenu.NoMatch
      Case True

        MsgBox "Erreur lors de la lecture de la table Élements du Menu Général."

        RstMenu.Close

        DBBase.Close

        GoTo Exit_ButtonClick

      Case False

    End Select

    Select Case RstMenu![Men_Command]
        Case conCmdAllerMenuGénéral        ' Sélectionne un autre Menu Général.

            Me.Filter = "[Men_Numero] = 0 AND [Men_Page]=" & RstMenu![Men_Argument]

        Case conCmdOuvrirFormulaireAjouter        ' Ouvre un formulaire en mode ajout.

            DoCmd.OpenForm RstMenu![Men_Argument], , , , acAdd

        Case conCmdOuvrirFormulaireParcourir        ' Ouvre un formulaire.

            Select Case IsNullVide(RstMenu![Men_Argument1])
               Case True

                  DoCmd.OpenForm RstMenu![Men_Argument]

               Case False  'Filtre supplémentaire : Condition Where

                DoCmd.OpenForm RstMenu![Men_Argument], , , RstMenu![Men_Argument1]

            End Select

        Case conCmdOuvrirEtat        ' Ouvre un état.

            DoCmd.OpenReport RstMenu![Men_Argument], acPreview

        Case conCmdPersonnaliserMenuGénéral
            ' Prend en charge le cas où le Menu Général
            ' n'est pas installé (p.ex. Installation minimale).
        Case conCmdQuitterApplication        ' Quitte l'application.
            'CloseCurrentDatabase
            Application.Quit
        ' Exécute une macro.
        Case conCmdExécuterMacro

            DoCmd.RunMacro RstMenu![Men_Argument]

        Case conCmdExécuterCode        ' Exécute le code.

            Application.Run RstMenu![Men_Argument]

        Case Else        ' Toute autre commande est inconnue.

            MsgBox "Option inconnue."

    End Select
    ' Ferme le recordset et la base de données.
    RstMenu.Close

    DBBase.Close

Exit_ButtonClick:

   Set RstMenu = Nothing

   Set DBBase = Nothing

   Exit Function

Err_ButtonClick:
    ' Si l'action a été annulée par l'utilisateur ne pas afficher de message d'erreur, mais reprendre à la ligne suivante.
    Select Case Err
      Case conErrDoCmdAnnulé

        Resume Next

      Case Else

        MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ButtonClick"

        Resume Exit_ButtonClick

    End Select

End Function

Private Sub Form_Open(Annuler As Integer)

   On Error GoTo Err_Form_Open

   Me.EtiVersion.Caption = "Version " & stVersion & " - " & stDateVersion

   Me.Filter = "[Men_Numero] = 0 AND [Men_Argument] = 'Défaut'"

   Me.FilterOn = True

Exit_Form_Open:

    Exit Sub

Err_Form_Open:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : ButtonClick"

    Resume Exit_Form_Open
End Sub

Private Sub Form_Current()

   On Error GoTo Err_Form_Current

   Me.Caption = Nz(Me![Men_Libelle], "")

   CreerOption

Exit_Form_Current:

   Exit Sub

Err_Form_Current:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Current"

   Resume Exit_Form_Current
End Sub
