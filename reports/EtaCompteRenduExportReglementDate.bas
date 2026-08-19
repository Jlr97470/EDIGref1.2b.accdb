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
    ItemSuffix =81
    Left =2445
    Top =1935
    OnNoData ="[Event Procedure]"
    Filter ="Cli_Code_Groupe='CHAMP'"
    RecSrcDt = Begin
        0xfb0aab96afeae240
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
            GroupFooter = NotDefault
            ControlSource ="Cli_Code"
        End
        Begin BreakLevel
            ControlSource ="Déb_Code"
        End
        Begin FormHeader
            Height =1560
            Name ="EntêteÉtat"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Width =9033
                    Height =480
                    FontSize =14
                    Name ="EtiCompteRenduExportReglementDate"
                    Caption ="Liste des règlements :"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5610
                    Top =49
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =855
                    Width =2325
                    Height =285
                    FontSize =11
                    Name ="EtiDateDebut"
                    Caption ="Période de virement :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =2355
                    Top =855
                    Width =1365
                    Height =270
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtDateDébut"
                    Format ="Short Date"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =3795
                    Top =855
                    Width =150
                    Height =270
                    FontSize =11
                    Name ="EtiDateFin"
                    Caption ="-"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    Left =4020
                    Top =855
                    Width =1365
                    Height =270
                    FontSize =11
                    FontWeight =700
                    TabIndex =2
                    Name ="TxtDateFin"
                    Format ="Short Date"
                End
            End
        End
        Begin PageHeader
            Height =0
            Name ="EntêtePage"
        End
        Begin BreakHeader
            Height =932
            Name ="EntêteGroupe0"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =45
                    Width =9033
                    Name ="Ligne26"
                End
                Begin Label
                    OldBorderStyle =0
                    Left =15
                    Top =180
                    Width =1860
                    Height =300
                    FontSize =11
                    FontWeight =400
                    Name ="EtiCli_Code"
                    Caption ="Référence Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Top =495
                    Width =5100
                    Height =345
                    FontSize =12
                    FontWeight =700
                    Name ="TxtCli_Nom"
                    ControlSource ="Cli_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =1920
                    Top =150
                    Width =945
                    Height =345
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =5279
                    Top =375
                    Width =585
                    Height =525
                    FontSize =10
                    Name ="EtiEXPR_Date_Virement"
                    Caption ="Date\015\012Virmt"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7260
                    Top =525
                    Width =1095
                    Height =285
                    FontSize =10
                    Name ="EtiERD_Montant_m"
                    Caption ="Règlement"
                End
                Begin Rectangle
                    Left =15
                    Top =930
                    Width =9016
                    Height =2
                    Name ="Boîte46"
                End
            End
        End
        Begin Section
            Height =600
            Name ="Détail1"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    Left =1170
                    Width =3705
                    Height =255
                    FontWeight =700
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Width =1155
                    Height =255
                    FontWeight =700
                    TabIndex =1
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
                    TabIndex =2
                    Name ="TxtEXPR_Date_Virement"
                    ControlSource ="EXPR_Date_Virement"
                    Format ="Short Date"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =6075
                    Top =195
                    Width =1365
                    Height =285
                    TabIndex =3
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
                    TabIndex =4
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
                    TabIndex =5
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
                    Top =540
                    Width =9016
                    Height =2
                    Name ="Boîte61"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7965
                    Top =195
                    Width =1095
                    Height =285
                    TabIndex =6
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
                    FontSize =10
                    FontWeight =400
                    Name ="EtiERD_Montant_mMonnaie"
                    Caption ="soit"
                End
            End
        End
        Begin BreakFooter
            ForceNewPage =2
            Height =510
            Name ="PiedGroupe1"
            Begin
                Begin Rectangle
                    Top =60
                    Width =9016
                    Height =2
                    Name ="Boîte52"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Left =3165
                    Top =105
                    Width =1530
                    Height =315
                    FontSize =12
                    Name ="EtiTotalERD_Montant_m"
                    Caption ="Total Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =4695
                    Top =135
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
                    Left =6675
                    Top =130
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
                    Left =7233
                    Top =135
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
        Begin PageFooter
            Height =0
            Name ="PiedPage"
        End
        Begin FormFooter
            Height =630
            Name ="PiedÉtat"
            Begin
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Left =3795
                    Top =240
                    Width =870
                    Height =345
                    FontSize =12
                    Name ="EtiTotalGeneralERD_Montant_m"
                    Caption ="Total  :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =4695
                    Top =270
                    Width =1920
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalGeneralERD_Montant_m"
                    ControlSource ="=Sum(Round(Nz([ERD_Montant_m1],0),2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =6675
                    Top =265
                    Width =510
                    Height =270
                    FontSize =10
                    FontWeight =400
                    Name ="EtiTotalGeneralERD_Montant_mMonnaie"
                    Caption ="soit"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =7233
                    Top =270
                    Width =1830
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalGeneralERD_Montant_mMonnaie"
                    ControlSource ="=[TotalGeneralERD_Montant_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduExportReglementDate.cls"
