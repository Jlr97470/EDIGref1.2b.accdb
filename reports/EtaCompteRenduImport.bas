Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    FilterOn = NotDefault
    DateGrouping =1
    GridY =10
    Width =9078
    ItemSuffix =70
    Left =795
    Top =1875
    OnNoData ="[Event Procedure]"
    Filter ="Imp_Id=694"
    RecSrcDt = Begin
        0x49ad3135afeae240
    End
    RecordSource ="SELECT DISTINCTROW CLIENT.*, IMPORTATION.*, IMPORTATION_DETAIL.*, DEBITEUR.* FRO"
        "M (IMPORTATION INNER JOIN (DEBITEUR INNER JOIN IMPORTATION_DETAIL ON DEBITEUR.Dé"
        "b_Lot_IMPD_Id=IMPORTATION_DETAIL.IMPD_Id) ON IMPORTATION.IMP_Id=IMPORTATION_DETA"
        "IL.IMPD_Code) INNER JOIN CLIENT ON IMPORTATION_DETAIL.IMPD_CodeClient=CLIENT.Cli"
        "_Code; "
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
            GroupHeader = NotDefault
            ControlSource ="IMP_Id"
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
                    Name ="EtiCompteRenduImport"
                    Caption ="Compte rendu d'importation : "
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
                    Name ="TrtCompteRenduImport"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =960
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
                    Top =960
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
                    Top =960
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
                    Left =6090
                    Top =1740
                    Width =1305
                    Height =285
                    FontSize =10
                    Name ="EtiDéb_Mont_Prin_m1"
                    Caption ="Principal"
                End
                Begin Rectangle
                    Top =2205
                    Width =9016
                    Height =2
                    Name ="RecCompteRenduImport"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="EntêteGroupe1"
            Begin
                Begin Label
                    OldBorderStyle =0
                    Top =75
                    Width =465
                    Height =225
                    Name ="EtiLot"
                    Caption ="Lot  :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =510
                    Top =60
                    Width =4620
                    Height =255
                    FontWeight =700
                    Name ="TxtLot"
                    ControlSource ="=[IMP_Lot_Date] & [IMP_Lot_Num] & [IMP_Lot_Site] & [IMP_Lot_CodeGREF] & [IMP_Lot"
                        "_Contrat]"
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
                    Width =9060
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
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="=Round([Déb_Mont_Prin_m1],2)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    OldBorderStyle =0
                    Left =495
                    Top =255
                    Width =2490
                    Height =225
                    TabIndex =5
                    Name ="EtiDéb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                End
                Begin Label
                    OldBorderStyle =0
                    Top =255
                    Width =495
                    Height =225
                    Name ="Texte58"
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
                    Name ="RecCompteRenduImport2"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =7920
                    Top =195
                    Width =1140
                    Height =285
                    TabIndex =7
                    Name ="TxtDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtDéb_Mont_Prin_m]*DonnerTauxConvetion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7485
                    Top =195
                    Width =405
                    Height =285
                    FontSize =10
                    FontWeight =400
                    Name ="EtiSoit"
                    Caption ="soit"
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
                    Name ="RecCompteRenduImport1"
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =1
                    Left =3990
                    Top =180
                    Width =1530
                    Height =315
                    FontSize =12
                    Name ="EtiTotalDéb_Mont_Prin_m"
                    Caption ="Total Client :"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    Left =5520
                    Top =180
                    Width =1905
                    Height =285
                    FontSize =10
                    FontWeight =700
                    Name ="TxtTotalDéb_Mont_Prin_m"
                    ControlSource ="=Sum(Round([Déb_Mont_Prin_m1],2))"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin Label
                    OldBorderStyle =0
                    TextAlign =2
                    Left =7470
                    Top =180
                    Width =420
                    Height =270
                    FontSize =10
                    FontWeight =400
                    Name ="EtiTotalSoit"
                    Caption ="soit"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    Left =7938
                    Top =195
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="TxtTotalDéb_Mont_Prin_mMonnaie"
                    ControlSource ="=[TxtTotalDéb_Mont_Prin_m]*DonnerTauxConvertion(1)"
                    Format ="(#,##0.00 €);(-#,##0.00 €)"
                End
            End
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduImport.cls"
