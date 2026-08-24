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
    ItemSuffix =73
    Left =3090
    Top =1935
    OnNoData ="[Event Procedure]"
    Filter ="EXPR_Id=1602"
    RecSrcDt = Begin
        0xe891f163b0eae240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.*, EXPORTATION_REG.*, EXPORTATION_REG_DETAIL.*, DEBITE"
        "UR.* FROM CLIENT INNER JOIN (EXPORTATION_REG INNER JOIN (EXPORTATION_REG_DETAIL "
        "INNER JOIN DEBITEUR ON EXPORTATION_REG_DETAIL.ERD_CodeDébiteur=DEBITEUR.Déb_Code"
        ") ON EXPORTATION_REG.EXPR_Id=EXPORTATION_REG_DETAIL.ERD_Code) ON CLIENT.Cli_Code"
        "=DEBITEUR.Déb_Code_Client; "
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
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="EXPR_Date_Virement"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            RepeatSection = NotDefault
            Name ="EntêteGroupe0"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =56
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiCompteRenduExportReglement"
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
                    Name ="TrtCompteRenduExportReglement"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =1125
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
                    Top =1125
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
                    Top =1125
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
                    Top =540
                    Width =9033
                    Height =480
                    FontSize =16
                    Name ="EtiCompteRenduExportReglement1"
                    Caption ="REGLEMENT"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =735
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5129
                    Top =105
                    Width =855
                    Height =585
                    FontSize =10
                    Name ="EtiDéb_Date_Cré"
                    Caption ="Date\015\012Créance"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7245
                    Top =255
                    Width =1095
                    Height =285
                    FontSize =10
                    Name ="EtiERD_Montant_m"
                    Caption ="Règlement"
                End
                Begin Rectangle
                    Top =720
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportReglement"
                End
                Begin Label
                    OldBorderStyle =0
                    Width =1320
                    Height =300
                    FontSize =11
                    FontWeight =400
                    Name ="EtiEXPR_Date_Virement"
                    Caption ="Virement le :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =1350
                    Width =1365
                    Height =300
                    FontSize =11
                    Name ="TxtEXPR_Date_Virement"
                    ControlSource ="EXPR_Date_Virement"
                    Format ="Short Date"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =900
            Name ="Détail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Top =480
                    Width =8940
                    Height =210
                    Name ="TxtDéb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Width =3705
                    Height =255
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Width =1155
                    Height =255
                    FontWeight =700
                    TabIndex =2
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
                    TabIndex =3
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
                    TabIndex =4
                    Name ="TxtERD_Montant_m"
                    ControlSource ="=Round([ERD_Montant_m1],2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =495
                    Top =255
                    Width =2490
                    Height =225
                    TabIndex =5
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
                    TabIndex =6
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
                Begin Rectangle
                    BorderLineStyle =3
                    Top =810
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportReglement2"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7920
                    Top =195
                    Width =1140
                    Height =285
                    TabIndex =7
                    Name ="TxtERD_Montant_mMonnaie"
                    ControlSource ="=[TxtERD_Montant_m]*DonnerTauxConvertion(1)"
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
                    Name ="EtiERD_Montant_mMonnaie"
                    Caption ="soit"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =405
            BreakLevel =1
            Name ="PiedGroupe0"
            Begin
                Begin Rectangle
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduExportReglement1"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Top =50
                    Width =1530
                    Height =315
                    FontSize =12
                    Name ="EtiTotalERD_Montant_m"
                    Caption ="Total Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =1530
                    Top =80
                    Width =1920
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalERD_Montant_m"
                    ControlSource ="=Sum(Round([ERD_Montant_m1],2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =3510
                    Top =75
                    Width =510
                    Height =270
                    FontSize =10
                    FontWeight =400
                    Name ="EtiTotalERD_Montant_mMonnaie"
                    Caption ="soit"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =4068
                    Top =80
                    Width =1830
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalERD_Montant_mMonnaie"
                    ControlSource ="=[TxtTotalERD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduExportReglement.cls"
