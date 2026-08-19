Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
   Dim IntNumero As Integer

   On Error GoTo Err_Report_Open

   For IntNumero = 1 To 16

      Select Case InStr(Me.RecordSource, Me.Controls("TxtXLS_IMPChamp" & IntNumero).ControlSource)
         Case Is > 0

         Case Else

            Me.Controls("TxtXLS_IMPChamp" & IntNumero).ControlSource = ""

            Me.Controls("TxtXLS_IMPChamp" & IntNumero).Visible = False

            Me.Controls("TxtDebiteurAdresseChamp" & IntNumero).ControlSource = ""

            Me.Controls("TxtDebiteurAdresseChamp" & IntNumero).Visible = False

            Me.Controls("TxtChamp" & IntNumero).ControlSource = ""

            Me.Controls("TxtChamp" & IntNumero).Visible = False

      End Select

   Next

Exit_Report_Open:

   Exit Sub

Err_Report_Open:

   MsgBox Err.Number & " " & Err.description, vbCritical, stSource & " : Report_Open"

   Resume Exit_Report_Open
End Sub
