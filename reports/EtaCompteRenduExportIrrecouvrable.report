Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9075
    ItemSuffix =77
    Left =1935
    Top =1635
    OnNoData ="[Event Procedure]"
    Filter ="Cli_Code_Groupe='CHAMP'"
    RecSrcDt = Begin
        0x40cac4dcb5eae240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.*, EXPORTATION_IRR.*, EXPORTATION_IRR_DETAIL.*, DEBITE"
        "UR.*, CODE_GESTION.CG_Libellé FROM (CLIENT INNER JOIN (EXPORTATION_IRR INNER JOI"
        "N (DEBITEUR INNER JOIN EXPORTATION_IRR_DETAIL ON DEBITEUR.Déb_Lot_EID_Id=EXPORTA"
        "TION_IRR_DETAIL.EID_Id) ON EXPORTATION_IRR.EXPI_Id=EXPORTATION_IRR_DETAIL.EID_Co"
        "de) ON CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client) INNER JOIN CODE_GESTION ON EXPO"
        "RTATION_IRR_DETAIL.EID_Situation=CODE_GESTION.CG_Code; "
    OnOpen ="[Event Procedure]"
    Begin
        Begin Label
            OldBorderStyle =1
            FontWeight =700
            FontName ="Arial"
        End
        Begin Rectangle
            BackStyle =0
            Width =850
            Height =280
        End
        Begin Line
            Width =1701
        End
        Begin BoundObjectFrame
            AutoLabel = NotDefault
            Width =1998
            Height =280
            LabelX =-1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            Width =1701
            Height =280
            LabelX =-1701
            FontName ="Arial"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2325
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =56
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiCompteRenduExportIrrecouvrable"
                    Caption ="Compte rendu d'exportation :"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5610
                    Top =105
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Line
                    BorderWidth =2
                    Width =9033
                    Name ="TxtCompteRenduExportIrrecouvrable"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =1170
                    Width =1875
                    Height =300
                    FontSize =11
                    FontWeight =400
                    Name ="EtiCli_Code"
                    Caption ="Référence Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =2880
                    Top =1170
                    Width =5430
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =1905
                    Top =1170
                    Width =945
                    Height =315
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5129
                    Top =1590
                    Width =855
                    Height =585
                    FontSize =10
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date\015\012Créance"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7140
                    Top =1740
                    Width =1305
                    Height =285
                    FontSize =10
                    Name ="EtiMontant_m"
                    Caption ="Principal"
                End
                Begin Rectangle
                    Top =2205
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportIrrecouvrable2"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =2963
                    Top =570
                    Width =3075
                    Height =435
                    FontSize =16
                    Name ="EtiCompteRenduExportIrrecouvrable1"
                    Caption ="IRRECOUVRABLES"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1095
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Top =480
                    Width =8940
                    Height =210
                    TabIndex =1
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Width =3915
                    Height =255
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Width =1155
                    Height =255
                    FontWeight =700
                    TabIndex =3
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5100
                    Top =195
                    Width =975
                    Height =285
                    TabIndex =4
                    Name ="TxtDéb_Date_Cré"
                    ControlSource ="Déb_Date_Cré"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6075
                    Top =195
                    Width =1365
                    Height =285
                    TabIndex =5
                    Name ="TxtEID_Montant_m"
                    ControlSource ="=Round([EID_Montant_m1],2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =495
                    Top =255
                    Width =2490
                    Height =225
                    TabIndex =6
                    Name ="TxtDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =255
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf1_Client"
                    Caption ="Réf1 :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =3480
                    Top =255
                    Width =1620
                    Height =225
                    TabIndex =7
                    Name ="TxtDéb_Réf2_Client"
                    ControlSource ="Déb_Réf2_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =2985
                    Top =255
                    Width =495
                    Height =225
                    Name ="EtiDéb_Réf2_Client"
                    Caption ="Réf2 :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7920
                    Top =195
                    Width =1140
                    Height =285
                    TabIndex =8
                    Name ="TxtEID_Montant_mMonnaie"
                    ControlSource ="=[TxtEID_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7485
                    Top =195
                    Width =435
                    Height =285
                    FontWeight =400
                    Name ="EtiEID_Montant_mMonnaie"
                    Caption ="soit"
                End
                Begin Rectangle
                    BorderLineStyle =3
                    Top =1005
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportIrrecouvrable"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1350
                    Top =690
                    Width =2625
                    Height =225
                    Name ="TxtCG_Libellé"
                    ControlSource ="CG_Libellé"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =690
                    Width =1335
                    Height =225
                    Name ="EtiCG_Libellé"
                    Caption ="Situation GREF :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =5385
                    Top =690
                    Width =2625
                    Height =225
                    TabIndex =9
                    Name ="TxtEID_Situation_Client"
                    ControlSource ="EID_Situation_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =3990
                    Top =690
                    Width =1365
                    Height =225
                    Name ="EtiEID_Situation_Client"
                    Caption ="Situation Client :"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =500
            Name ="PiedGroupe1"
            Begin
                Begin Rectangle
                    Top =135
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportIrrecouvrable1"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =185
                    Width =1530
                    Height =315
                    FontSize =12
                    Name ="EtiTotalEID_Montant_m"
                    Caption ="Total Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =1530
                    Top =215
                    Width =1920
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalEID_Montant_m"
                    ControlSource ="=Sum(Round([EID_Montant_m1],2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =3510
                    Top =210
                    Width =510
                    Height =270
                    FontSize =10
                    FontWeight =400
                    Name ="EtiTotalEID_Montant_mMonnaie"
                    Caption ="soit"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =4068
                    Top =215
                    Width =1830
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalEID_Montant_mMonnaie"
                    ControlSource ="=[TxtTotalEID_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduExportIrrecouvrable.cls"
