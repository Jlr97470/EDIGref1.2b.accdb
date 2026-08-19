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
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =6051
    ItemSuffix =49
    Left =3810
    Top =75
    Right =10215
    Bottom =5055
    Filter ="ERD_Code=147"
    RecSrcDt = Begin
        0x6720f8ac4dede240
    End
    RecordSource ="SELECT DISTINCTROW EXPORTATION_REG_DETAIL.*, EXPORTATION_REG_DETAIL.ERD_CodeDébi"
        "teur FROM EXPORTATION_REG_DETAIL ORDER BY EXPORTATION_REG_DETAIL.ERD_CodeDébiteu"
        "r; "
    Caption ="Règlements définitifs à transmettre"
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
            Height =1080
            BackColor =12632256
            Name ="EntêteFormulaire1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =45
                    Top =840
                    Width =675
                    Height =240
                    Name ="EtiERD_CodeDébiteur"
                    Caption ="Code"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =1230
                    Top =840
                    Width =3585
                    Height =240
                    Name ="EtiERD_NomDébiteur"
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =4890
                    Top =840
                    Width =1080
                    Height =240
                    Name ="EtiERD_Montant_m"
                    Caption ="Montant"
                End
            End
        End
        Begin Section
            Height =270
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =30
                    Top =15
                    Width =1101
                    Height =240
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtERD_CodeDébiteur"
                    ControlSource ="EXPORTATION_REG_DETAIL.ERD_CodeDébiteur"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1230
                    Top =15
                    Width =3576
                    Height =240
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtERD_NomDébiteur"
                    ControlSource ="ERD_NomDébiteur"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    Left =4875
                    Top =15
                    Width =1101
                    Height =240
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtERD_Montant_m"
                    ControlSource ="=IIf([ERD_Montant_mo]=1,[ERD_Montant_m1],[ERD_Montant_m2])"
                    Format ="Standard"
                End
            End
        End
        Begin FormFooter
            Height =675
            BackColor =12632256
            Name ="PiedFormulaire2"
            Begin
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =4125
                    Top =105
                    Width =1635
                    Height =450
                    FontWeight =300
                    Name ="CmdTransmettre"
                    Caption ="Transmettre"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "FrmExportationReglementArchive.cls"
