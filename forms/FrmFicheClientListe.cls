Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Utilise l'ordre de la base pour les comparaisons de chaînes
Option Explicit

Public SQLBase As String
Public SQLWhere As String
Public SQLWhereGlobal As String
Public SQLOrder As String

Private Sub Form_Open(Cancel As Integer)

   On Error GoTo Err_Form_Open

   SQLBase = "SELECT * FROM CLIENT"

   SQLWhere = ""

   SQLWhereGlobal = Nz(Me.OpenArgs)

   SQLOrder = "Cli_Code"

   Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_Form_Open:

   Exit Sub

Err_Form_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Form_Open"

   Resume Exit_Form_Open
End Sub

Private Sub CmdImprimer_Click()

    On Error GoTo Err_CmdImprimer_Click

    DoCmd.OpenForm "FrmRapportListe", acNormal, , , , acWindowNormal

Exit_CmdImprimer_Click:

    Exit Sub

Err_CmdImprimer_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdImprimer_Click"

    Resume Exit_CmdImprimer_Click
End Sub

Private Sub CmdOk_Click()

   On Error GoTo Err_CmdOk_Click

    StrFicheClientListeCli_Code = Cli_Code

    DoCmd.Close acForm, Me.Name

Exit_CmdOk_Click:

    Exit Sub

Err_CmdOk_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdOk_Click"

    Resume Exit_CmdOk_Click
End Sub

Private Sub CmdRechercher_Click()

    On Error GoTo Err_CmdRechercher_Click

    DoCmd.OpenForm "FrmFicheClientRechercher", acNormal, , , , acDialog, "FrmFicheClientListe"

    Me.RecordSource = SQLBase & IIf((SQLWhereGlobal & SQLWhere) = "", " ", " WHERE ") & SQLWhereGlobal & IIf(SQLWhereGlobal <> "" And SQLWhere <> "", " AND " & SQLWhere, SQLWhere) & IIf(SQLOrder = "", "", " ORDER BY " & SQLOrder) & ";"

Exit_CmdRechercher_Click:

    Exit Sub

Err_CmdRechercher_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdRechercher_Click"

    Resume Exit_CmdRechercher_Click
End Sub

Private Sub CmdFermer_Click()

    On Error GoTo Err_CmdFermer_Click

    StrFicheClientListeCli_Code = ""

    DoCmd.Close acForm, Me.Name

Exit_CmdFermer_Click:

    Exit Sub

Err_CmdFermer_Click:

    MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : CmdFermer_Click"

    Resume Exit_CmdFermer_Click
End Sub
