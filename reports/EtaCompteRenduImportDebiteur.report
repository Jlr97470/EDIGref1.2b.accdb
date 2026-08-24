Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9033
    DatasheetFontHeight =10
    ItemSuffix =66
    Left =2040
    Top =3060
    DatasheetGridlinesColor =12632256
    RecordSource ="SELECT [Déb_Code], [Déb_Code] AS DebiteurAdresseChamp1, [Déb_Nom] AS DebiteurAdr"
        "esseChamp2, [AD_Adr1] AS DebiteurAdresseChamp3, [AD_Adr2] AS DebiteurAdresseCham"
        "p4, [AD_Adr3] AS DebiteurAdresseChamp5, [AD_Adr4] AS DebiteurAdresseChamp6, [AD_"
        "CodePostal] AS DebiteurAdresseChamp7, [AD_Ville] AS DebiteurAdresseChamp8, [AD_T"
        "el] AS DebiteurAdresseChamp9, [AD_Tel2] AS DebiteurAdresseChamp10, [AD_Fax] AS D"
        "ebiteurAdresseChamp11, [AD_email] AS DebiteurAdresseChamp12, [Ref Gref] AS XLS_I"
        "MPChamp1, [Nom Prénom] AS XLS_IMPChamp2, [Adresse 1] AS XLS_IMPChamp3, [Adresse "
        "2] AS XLS_IMPChamp4, [Adresse 3] AS XLS_IMPChamp5, [Adresse 4] AS XLS_IMPChamp6,"
        " [Code Postal] AS XLS_IMPChamp7, [Ville] AS XLS_IMPChamp8, [Téléphone 1] AS XLS_"
        "IMPChamp9, [Téléphone 2] AS XLS_IMPChamp10, [Fax] AS XLS_IMPChamp11, [Email] AS "
        "XLS_IMPChamp12 FROM XLS_IMP INNER JOIN SelDebiteurAdresse ON (XLS_IMP.[Nom Préno"
        "m]=SelDebiteurAdresse.Déb_Nom) AND (XLS_IMP.[Ref Gref]=SelDebiteurAdresse.Déb_Co"
        "de); "
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1134
            Name ="EntêteÉtat"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =1
                    TextFontFamily =34
                    Width =9033
                    Height =480
                    FontSize =14
                    FontWeight =700
                    Name ="EtiCompteRenduImport"
                    Caption ="Compte rendu d'importation : "
                End
                Begin TextBox
                    FontItalic = NotDefault
                    FELineBreak = NotDefault
                    TextAlign =3
                    TextFontFamily =34
                    Left =5610
                    Top =49
                    Width =3405
                    Height =285
                    FontSize =10
                    Name ="TxtDate"
                    ControlSource ="=Date()"
                    Format ="Long Date"
                    AsianLineBreak =0
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanShrink = NotDefault
            Height =4095
            Name ="Détail"
            Begin
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =240
                    Width =4176
                    TabIndex =7
                    Name ="TxtDebiteurAdresseChamp1"
                    ControlSource ="DebiteurAdresseChamp1"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =480
                    Width =4176
                    TabIndex =17
                    Name ="TxtDebiteurAdresseChamp2"
                    ControlSource ="DebiteurAdresseChamp2"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =720
                    Width =4176
                    TabIndex =18
                    Name ="TxtDebiteurAdresseChamp3"
                    ControlSource ="DebiteurAdresseChamp3"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =960
                    Width =4176
                    TabIndex =19
                    Name ="TxtDebiteurAdresseChamp4"
                    ControlSource ="DebiteurAdresseChamp4"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =1200
                    Width =4176
                    TabIndex =20
                    Name ="TxtDebiteurAdresseChamp5"
                    ControlSource ="DebiteurAdresseChamp5"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =1440
                    Width =4176
                    TabIndex =21
                    Name ="TxtDebiteurAdresseChamp6"
                    ControlSource ="DebiteurAdresseChamp6"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =1680
                    Width =4176
                    TabIndex =22
                    Name ="TxtDebiteurAdresseChamp7"
                    ControlSource ="DebiteurAdresseChamp7"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =1920
                    Width =4176
                    TabIndex =23
                    Name ="TxtDebiteurAdresseChamp8"
                    ControlSource ="DebiteurAdresseChamp8"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =240
                    Width =4281
                    TabIndex =24
                    Name ="TxtXLS_IMPChamp1"
                    ControlSource ="XLS_IMPChamp1"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =480
                    Width =4281
                    TabIndex =25
                    Name ="TxtXLS_IMPChamp2"
                    ControlSource ="XLS_IMPChamp2"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =720
                    Width =4281
                    TabIndex =26
                    Name ="TxtXLS_IMPChamp3"
                    ControlSource ="XLS_IMPChamp3"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =960
                    Width =4281
                    TabIndex =27
                    Name ="TxtXLS_IMPChamp4"
                    ControlSource ="XLS_IMPChamp4"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =1200
                    Width =4281
                    TabIndex =28
                    Name ="TxtXLS_IMPChamp5"
                    ControlSource ="XLS_IMPChamp5"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =1440
                    Width =4281
                    TabIndex =29
                    Name ="TxtXLS_IMPChamp6"
                    ControlSource ="XLS_IMPChamp6"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =1680
                    Width =4281
                    TabIndex =30
                    Name ="TxtXLS_IMPChamp7"
                    ControlSource ="XLS_IMPChamp7"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =1920
                    Width =4281
                    TabIndex =31
                    Name ="TxtXLS_IMPChamp8"
                    ControlSource ="XLS_IMPChamp8"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =2160
                    Width =4176
                    TabIndex =32
                    Name ="TxtDebiteurAdresseChamp9"
                    ControlSource ="DebiteurAdresseChamp9"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =2400
                    Width =4176
                    TabIndex =33
                    Name ="TxtDebiteurAdresseChamp10"
                    ControlSource ="DebiteurAdresseChamp10"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =2640
                    Width =4176
                    TabIndex =34
                    Name ="TxtDebiteurAdresseChamp11"
                    ControlSource ="DebiteurAdresseChamp11"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =2880
                    Width =4176
                    TabIndex =35
                    Name ="TxtDebiteurAdresseChamp12"
                    ControlSource ="DebiteurAdresseChamp12"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =3120
                    Width =4176
                    TabIndex =36
                    Name ="TxtDebiteurAdresseChamp13"
                    ControlSource ="DebiteurAdresseChamp13"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =3360
                    Width =4176
                    TabIndex =37
                    Name ="TxtDebiteurAdresseChamp14"
                    ControlSource ="DebiteurAdresseChamp14"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =3600
                    Width =4176
                    TabIndex =38
                    Name ="TxtDebiteurAdresseChamp15"
                    ControlSource ="DebiteurAdresseChamp15"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4845
                    Top =3840
                    Width =4176
                    TabIndex =39
                    Name ="TxtDebiteurAdresseChamp16"
                    ControlSource ="DebiteurAdresseChamp16"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2160
                    Width =4281
                    TabIndex =40
                    Name ="TxtXLS_IMPChamp9"
                    ControlSource ="XLS_IMPChamp9"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2400
                    Width =4281
                    TabIndex =41
                    Name ="TxtXLS_IMPChamp10"
                    ControlSource ="XLS_IMPChamp10"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2640
                    Width =4281
                    TabIndex =42
                    Name ="TxtXLS_IMPChamp11"
                    ControlSource ="XLS_IMPChamp11"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =2880
                    Width =4281
                    TabIndex =43
                    Name ="TxtXLS_IMPChamp12"
                    ControlSource ="XLS_IMPChamp12"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3120
                    Width =4281
                    TabIndex =44
                    Name ="TxtXLS_IMPChamp13"
                    ControlSource ="XLS_IMPChamp13"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3360
                    Width =4281
                    TabIndex =45
                    Name ="TxtXLS_IMPChamp14"
                    ControlSource ="XLS_IMPChamp14"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3600
                    Width =4281
                    TabIndex =46
                    Name ="TxtXLS_IMPChamp15"
                    ControlSource ="XLS_IMPChamp15"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Top =3840
                    Width =4281
                    TabIndex =47
                    Name ="TxtXLS_IMPChamp16"
                    ControlSource ="XLS_IMPChamp16"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =4845
                    ColumnWidth =1230
                    FontWeight =700
                    TabIndex =48
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                    StatusBarText ="Debiteur Code"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =240
                    Width =570
                    TabIndex =8
                    Name ="TxtChamp1"
                    ControlSource ="=IIf([XLS_IMPChamp1]<>[DebiteurAdresseChamp1] Or (IsNull([XLS_IMPChamp1]) Xor Is"
                        "Null([DebiteurAdresseChamp1])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =480
                    Width =570
                    TabIndex =10
                    Name ="TxtChamp2"
                    ControlSource ="=IIf([XLS_IMPChamp2]<>[DebiteurAdresseChamp2] Or (IsNull([XLS_IMPChamp2]) Xor Is"
                        "Null([DebiteurAdresseChamp2])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =720
                    Width =570
                    Name ="TxtChamp3"
                    ControlSource ="=IIf([XLS_IMPChamp3]<>[DebiteurAdresseChamp3] Or (IsNull([XLS_IMPChamp3]) Xor Is"
                        "Null([DebiteurAdresseChamp3])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =960
                    Width =570
                    TabIndex =6
                    Name ="TxtChamp4"
                    ControlSource ="=IIf([XLS_IMPChamp4]<>[DebiteurAdresseChamp4] Or (IsNull([XLS_IMPChamp4]) Xor Is"
                        "Null([DebiteurAdresseChamp4])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =1200
                    Width =570
                    TabIndex =2
                    Name ="TxtChamp5"
                    ControlSource ="=IIf([XLS_IMPChamp5]<>[DebiteurAdresseChamp5] Or (IsNull([XLS_IMPChamp5]) Xor Is"
                        "Null([DebiteurAdresseChamp5])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =1440
                    Width =570
                    TabIndex =1
                    Name ="TxtChamp6"
                    ControlSource ="=IIf([XLS_IMPChamp6]<>[DebiteurAdresseChamp6] Or (IsNull([XLS_IMPChamp6]) Xor Is"
                        "Null([DebiteurAdresseChamp6])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =1680
                    Width =570
                    TabIndex =12
                    Name ="TxtChamp7"
                    ControlSource ="=IIf([XLS_IMPChamp7]<>[DebiteurAdresseChamp7] Or (IsNull([XLS_IMPChamp7]) Xor Is"
                        "Null([DebiteurAdresseChamp7])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =1920
                    Width =570
                    TabIndex =15
                    Name ="TxtChamp8"
                    ControlSource ="=IIf([XLS_IMPChamp8]<>[DebiteurAdresseChamp8] Or (IsNull([XLS_IMPChamp8]) Xor Is"
                        "Null([DebiteurAdresseChamp8])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =2160
                    Width =570
                    TabIndex =3
                    Name ="TxtChamp9"
                    ControlSource ="=IIf([XLS_IMPChamp9]<>[DebiteurAdresseChamp9] Or (IsNull([XLS_IMPChamp9]) Xor Is"
                        "Null([DebiteurAdresseChamp9])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =2400
                    Width =570
                    TabIndex =4
                    Name ="TxtChamp10"
                    ControlSource ="=IIf([XLS_IMPChamp10]<>[DebiteurAdresseChamp10] Or (IsNull([XLS_IMPChamp10]) Xor"
                        " IsNull([DebiteurAdresseChamp10])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =2640
                    Width =570
                    TabIndex =14
                    Name ="TxtChamp11"
                    ControlSource ="=IIf([XLS_IMPChamp8]<>[DebiteurAdresseChamp11] Or (IsNull([XLS_IMPChamp11]) Xor "
                        "IsNull([DebiteurAdresseChamp11])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =2880
                    Width =570
                    TabIndex =11
                    Name ="TxtChamp12"
                    ControlSource ="=IIf([XLS_IMPChamp12]<>[DebiteurAdresseChamp12] Or (IsNull([XLS_IMPChamp12]) Xor"
                        " IsNull([DebiteurAdresseChamp12])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =3360
                    Width =570
                    TabIndex =13
                    Name ="TxtChamp14"
                    ControlSource ="=IIf([XLS_IMPChamp14]<>[DebiteurAdresseChamp14] Or (IsNull([XLS_IMPChamp14]) Xor"
                        " IsNull([DebiteurAdresseChamp14])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =3120
                    Width =570
                    TabIndex =9
                    Name ="TxtChamp13"
                    ControlSource ="=IIf([XLS_IMPChamp13]<>[DebiteurAdresseChamp13] Or (IsNull([XLS_IMPChamp13]) Xor"
                        " IsNull([DebiteurAdresseChamp13])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =3600
                    Width =570
                    TabIndex =5
                    Name ="TxtChamp15"
                    ControlSource ="=IIf([XLS_IMPChamp15]<>[DebiteurAdresseChamp15] Or (IsNull([XLS_IMPChamp15]) Xor"
                        " IsNull([DebiteurAdresseChamp15])),\"=>\",Null)"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4275
                    Top =3840
                    Width =570
                    TabIndex =16
                    Name ="TxtChamp16"
                    ControlSource ="=IIf([XLS_IMPChamp16]<>[DebiteurAdresseChamp16] Or (IsNull([XLS_IMPChamp16]) Xor"
                        " IsNull([DebiteurAdresseChamp16])),\"=>\",Null)"
                End
                Begin Line
                    Top =4095
                    Width =8910
                    Name ="TrtCompte"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PiedÉtat"
        End
    End
End
CodeBehindForm
' See "EtaCompteRenduImportDebiteur.cls"
