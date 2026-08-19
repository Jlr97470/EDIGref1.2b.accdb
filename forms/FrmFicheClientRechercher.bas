Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =5100
    ItemSuffix =88
    Left =4335
    Top =60
    Right =9990
    Bottom =3105
    RecSrcDt = Begin
        0xd8bdb33134e3e240
    End
    Caption ="Rechercher un client"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin CommandButton
            TextFontFamily =2
            Width =1701
            Height =283
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin TextBox
            Width =1701
            LabelX =-1701
        End
        Begin ComboBox
            Width =1701
            LabelX =-1701
        End
        Begin FormHeader
            Height =2145
            BackColor =12632256
            Name ="EntêteFormulaire0"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =683
                    Top =113
                    Width =4281
                    Height =1925
                    BackColor =8421504
                    Name ="RecFicheClienrRechercher"
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =3244
                    Top =415
                    Width =1131
                    Name ="TxtCli_Code"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =1877
                            Top =415
                            Width =1230
                            Height =240
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiCli_Code"
                            Caption ="Code  :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    Left =2003
                    Top =750
                    Width =2376
                    TabIndex =1
                    Name ="TxtCli_Nom"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =1425
                            Top =750
                            Width =555
                            Height =240
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiCli_Nom"
                            Caption ="Nom :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    DisplayWhen =2
                    OverlapFlags =85
                    Left =113
                    Top =113
                    Width =471
                    Height =471
                    TabIndex =3
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
                        0x0000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =120
                    Top =660
                    Width =472
                    Height =472
                    TabIndex =2
                    Name ="CmdOK"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =223
                    Left =4440
                    Top =135
                    Width =525
                    Height =195
                    BackColor =8421504
                    ForeColor =65535
                    Name ="EtiClass"
                    Caption ="Class."
                    FontName ="MS Sans Serif"
                End
                Begin OptionGroup
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =215
                    Left =4500
                    Top =330
                    Width =392
                    Height =1715
                    TabIndex =4
                    Name ="CadClass"
                    DefaultValue ="0"
                    Begin
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4611
                            Top =463
                            OptionValue =0
                            Name ="OptCli_Code"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4611
                            Top =793
                            OptionValue =1
                            Name ="OptCli_Nom"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4605
                            Top =1140
                            OptionValue =2
                            Name ="OptCG_Parent_Id"
                        End
                        Begin OptionButton
                            OverlapFlags =215
                            Left =4602
                            Top =1485
                            OptionValue =3
                            Name ="OptCli_Groupe"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2198
                    Top =1440
                    Width =2181
                    Height =255
                    TabIndex =5
                    Name ="CmbCli_Groupe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE INNER JOIN C"
                        "LIENT_GROUPE AS CLIENT_GROUPE_1 ON CLIENT_GROUPE.CG_Parent_Id=CLIENT_GROUPE_1.CG"
                        "_Id WHERE (((CLIENT_GROUPE.CG_Parent_Id)=Formulaires!FrmFicheClientRechercher!Cm"
                        "bGroupe)); "
                    ColumnWidths ="0;1701"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =915
                            Top =1440
                            Width =1260
                            Height =240
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiCli_Groupe"
                            Caption ="Sous Groupe :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2198
                    Top =1080
                    Width =2181
                    Height =255
                    TabIndex =6
                    Name ="CmbGroupe"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CLIENT_GROUPE.CG_Id, CLIENT_GROUPE.CG_Nom FROM CLIENT_GROUPE WHERE (((CLI"
                        "ENT_GROUPE.CG_Parent_Id) Is Null)); "
                    ColumnWidths ="0;1701"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            TextAlign =3
                            Left =915
                            Top =1080
                            Width =1260
                            Height =240
                            BackColor =8421504
                            ForeColor =16777215
                            Name ="EtiCG_Parent_Id"
                            Caption ="Groupe :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin Section
            Height =20
            BackColor =12632256
            Name ="Détail0"
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire1"
        End
    End
End
CodeBehindForm
' See "FrmFicheClientRechercher.cls"
