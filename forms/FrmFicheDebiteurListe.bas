Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =7596
    ItemSuffix =77
    Left =3660
    Top =1560
    Right =11595
    Bottom =7845
    Filter ="Déb_Code_Client='90170' AND Déb_Réf1_Client='457193068'"
    RecSrcDt = Begin
        0x3774c71a16eae240
    End
    RecordSource ="SELECT DEBITEUR.*, ADRESSE_DEBITEUR.AD_CodePostal, ADRESSE_DEBITEUR.AD_Ville FRO"
        "M DEBITEUR INNER JOIN ADRESSE_DEBITEUR ON DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_"
        "Code_Débiteur; "
    Caption ="Liste des Débiteurs"
    AllowDatasheetView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
        End
        Begin Rectangle
            SpecialEffect =2
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
        End
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
        End
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =1701
            Height =255
            LabelX =-1701
        End
        Begin FormHeader
            Height =931
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Top =690
                    Width =1080
                    Height =240
                    Name ="EtiDéb_Code"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1110
                    Top =690
                    Width =3585
                    Height =240
                    Name ="EtiDéb_Nom"
                    Caption ="Nom"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =471
                    Height =456
                    Name ="CmdFermer"
                    Caption ="Fermer"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000019000000180000000100040000000000800100000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888888888888888888888888880888888 ,
                        0x8888888888888888888888880088888888888888888888888888888030888888 ,
                        0x8888888888888888888888033088888888888888888888888888803330888888 ,
                        0x8888888888888888888803333088888888888888888888888000333330000000 ,
                        0x00000888888888888880333330ffff0888888888888888888880333330ffff08 ,
                        0x88888888888888888880333300ffff0888888888888888888880333330ffff08 ,
                        0x88888488888888888880333330ffff0888884488888888888880333330ffff08 ,
                        0x88844488888888888880333330ffff0888444444488888888880333330ffff08 ,
                        0x84444444488888888880333330ffff0884444444488888888880333330ffff08 ,
                        0x8844444448888888888033330fffff08888444888888800088803330ffffff08 ,
                        0x88884488888880ff8880330fffffff0888888488888880ff888030ffffffff08 ,
                        0x888888888888800088800fffffffff0888888888888888888880000000000008 ,
                        0x8888888888888888000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    Left =4695
                    Top =690
                    Width =1140
                    Height =240
                    Name ="EtiDéb_Mont_Prin_m"
                    Caption ="Créance"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =5895
                    Top =691
                    Width =420
                    Height =240
                    Name ="EtiDéb_Code_Enc"
                    Caption ="Enc."
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =6330
                    Top =691
                    Width =420
                    Height =240
                    Name ="EtiDéb_Situation"
                    Caption ="Sit."
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =6765
                    Top =691
                    Width =420
                    Height =240
                    Name ="EtiDéb_niv_relance"
                    Caption ="Rel."
                End
            End
        End
        Begin Section
            Height =855
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Top =15
                    Width =1101
                    Height =240
                    Name ="TxtDéb_Code"
                    ControlSource ="Déb_Code"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =4755
                    Top =15
                    Width =1131
                    Height =240
                    TabIndex =2
                    Name ="TxtDéb_Mont_Prin_m"
                    ControlSource ="Déb_Mont_Prin_m2"
                    Format ="Standard"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1155
                    Top =15
                    Width =3576
                    Height =240
                    TabIndex =1
                    Name ="TxtDéb_Nom"
                    ControlSource ="Déb_Nom"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =5915
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =3
                    Name ="TxtDéb_Code_Enc"
                    ControlSource ="Déb_Code_Enc"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =6350
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =4
                    Name ="TxtDéb_Situation"
                    ControlSource ="Déb_Situation"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1155
                    Top =285
                    Width =1011
                    Height =240
                    TabIndex =6
                    Name ="TxtAD_CodePostal"
                    ControlSource ="AD_CodePostal"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =2190
                    Top =285
                    Width =3696
                    Height =240
                    TabIndex =7
                    Name ="TxtAD_Ville"
                    ControlSource ="AD_Ville"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =6785
                    Top =15
                    Width =397
                    Height =240
                    TabIndex =5
                    Name ="TxtDéb_niv_relance"
                    ControlSource ="Déb_niv_relance"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1155
                    Top =570
                    TabIndex =8
                    Name ="Déb_Réf1_Client"
                    ControlSource ="Déb_Réf1_Client"
                    StatusBarText ="Debiteur Reference Client 1"
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =570
                    Width =4326
                    TabIndex =9
                    Name ="Déb_Typ_Cré"
                    ControlSource ="Déb_Typ_Cré"
                    StatusBarText ="Debiteur Type De Creance"
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "FrmFicheDebiteurListe.cls"
