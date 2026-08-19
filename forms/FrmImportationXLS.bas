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
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =1
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11745
    DatasheetFontHeight =10
    ItemSuffix =15
    Left =1725
    Top =255
    Right =13815
    Bottom =7350
    RecSrcDt = Begin
        0x4f5b376570ebe240
    End
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    AllowDatasheetView =0
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
            Width =850
            Height =850
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin CustomControl
            SpecialEffect =2
            Width =4536
            Height =2835
        End
        Begin FormHeader
            Height =1425
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin CommandButton
                    DisplayWhen =2
                    OverlapFlags =85
                    TextFontFamily =2
                    Left =165
                    Top =113
                    Width =471
                    Height =471
                    FontSize =10
                    FontWeight =700
                    ForeColor =0
                    Name ="CmdFermer"
                    Caption ="FermerSimple"
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
                        0x8888888888888888
                    End
                    FontName ="System"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2842
                    Top =195
                    Width =4365
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8355711
                    Name ="EtiImportationTitreOmbre"
                    Caption ="Importation des créances"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =2812
                    Top =150
                    Width =4365
                    Height =405
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiImportationTitre"
                    Caption ="Importation des créances"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =165
                    Top =735
                    Width =1260
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiIMPX_Cli_Code"
                    Caption ="Client         :"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2385
                    Top =735
                    Width =3314
                    TabIndex =3
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="TxtIMPX_Cli_Nom"
                    ControlSource ="=DLookUp(\"Cli_Nom\",\"CLIENT\",\"Cli_Code='\" & [TxtIMPX_Cli_Code] & \"'\")"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin TextBox
                    FELineBreak = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1425
                    Top =735
                    Width =615
                    Height =255
                    TabIndex =1
                    Name ="TxtIMPX_Cli_Code"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2100
                    Top =735
                    Width =260
                    Height =275
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =0
                    Name ="CmdClientListe"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x280000000a000000040000000100040000000000200000000000000000000000 ,
                        0x1000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888800000080080080080000008008008008000000 ,
                        0x8888888888000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =10500
                    Top =225
                    Width =720
                    Height =270
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiRien"
                    Caption ="Rien"
                    FontName ="MS Sans Serif"
                    OnClick ="=ToutRien(False)"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =9750
                    Top =225
                    Width =720
                    Height =270
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiTout"
                    Caption ="Tout"
                    FontName ="MS Sans Serif"
                    OnClick ="=ToutRien(True)"
                End
                Begin Rectangle
                    SpecialEffect =1
                    OverlapFlags =215
                    Left =9645
                    Top =120
                    Width =1695
                    Height =450
                    Name ="RecToutRien"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =8220
                    Top =750
                    Width =3141
                    TabIndex =4
                    Name ="TxtIMPX_Fichier"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =1
                            Left =6810
                            Top =750
                            Width =1410
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            ForeColor =16777215
                            Name ="EtiIMPX_Fichier"
                            Caption ="Fichier :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    Left =780
                    Top =120
                    Width =472
                    Height =472
                    FontWeight =700
                    TabIndex =5
                    ForeColor =0
                    Name ="CmdImprimer"
                    Caption ="&i"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x1000000010000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00888888888888888888000000000008888088888888808088 ,
                        0x00000000000008080888888bbb88000808888887778808080000000000000880 ,
                        0x08888888888080808000000000080800880ffffffff080808880f00000f00008 ,
                        0x8880ffffffff088888880f00000f088888880ffffffff0888888800000000088 ,
                        0x8888888888888888
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    TextAlign =1
                    Left =8220
                    Top =1110
                    Width =3135
                    TabIndex =6
                    BackColor =12632256
                    Name ="TxtIMPX_Date"
                    ControlSource ="=dlookup(\"IMPX_Date\",\"IMPORTATIONXLS\",\"IMPX_Cli_Code='\" & [TxtIMPX_Cli_Cod"
                        "e] & \"' AND IMPX_Fichier='\" & [TxtIMPX_Fichier] & \"'\")"
                    Format ="Long Date"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =93
                            TextAlign =3
                            Left =6810
                            Top =1110
                            Width =1410
                            Height =240
                            FontWeight =700
                            BackColor =12632256
                            Name ="EtiIMPX_Date"
                            Caption ="Importé le :"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3750
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin Subform
                    OverlapFlags =85
                    Left =1860
                    Width =9885
                    Height =3735
                    Name ="SFrmImportationXLS"
                    SourceObject ="Form.SFrmImportationXLS"

                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Left =684
                    Top =540
                    Width =993
                    Height =255
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtIMPX_Date_Début"
                    ControlSource ="=DMin(\"Date\",\"XLS_IMP\")"
                    Format ="Short Date"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Left =684
                    Top =870
                    Width =993
                    Height =255
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtIMPX_Date_Fin"
                    ControlSource ="=DMax(\"Date\",\"XLS_IMP\")"
                    Format ="Short Date"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =165
                    Top =105
                    Width =1530
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiDate"
                    Caption ="Créances datant"
                    FontName ="MS Sans Serif"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =405
                    Top =1905
                    Width =1019
                    Height =285
                    FontWeight =700
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtIMPX_Nb_Créances"
                    ControlSource ="=DCount(\"[Ref Client]\",\"XLS_IMP\")"
                    Format ="Standard"
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =255
                    Left =45
                    Top =1350
                    Width =1725
                    Height =960
                    Name ="RecIMP_Nb_Créances"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =93
                    TextAlign =3
                    Left =150
                    Top =3000
                    Width =1529
                    Height =285
                    FontWeight =700
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtIMPX_Montant_m1"
                    ControlSource ="=IIf(DLookUp(\"Cli_MonnaieEnCours\",\"CLIENT\",\"Cli_Code='\" & [TxtIMPX_Cli_Cod"
                        "e] & \"'\")=1,DSum(\"Replace(Montant,DonnerSeparateur(1),DonnerSeparateur(0))\","
                        "\"XLS_IMP\"),0)"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =255
                    Left =45
                    Top =2445
                    Width =1725
                    Height =1275
                    Name ="RecIMP_Montant_m"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    FELineBreak = NotDefault
                    OverlapFlags =247
                    TextAlign =3
                    Left =150
                    Top =3330
                    Width =1529
                    Height =285
                    FontWeight =700
                    TabIndex =5
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtIMPX_Montant_m2"
                    ControlSource ="=IIf(DLookUp(\"Cli_MonnaieEnCours\",\"CLIENT\",\"Cli_Code='\" & [TxtIMPX_Cli_Cod"
                        "e] & \"'\")=2,DSum(\"Replace(Montant,DonnerSeparateur(1),DonnerSeparateur(0))\","
                        "\"XLS_IMP\"),0)"
                    Format ="#,##0.00 €;-#,##0.00 €"
                    DefaultValue ="\"0.00 E\""
                    FontName ="MS Sans Serif"
                    AsianLineBreak =0

                End
                Begin Rectangle
                    SpecialEffect =2
                    OverlapFlags =255
                    Left =45
                    Top =105
                    Width =1725
                    Height =1125
                    Name ="RecDate"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    Left =165
                    Top =540
                    Width =480
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiIMPX_Date_Début"
                    Caption ="du  :"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =1
                    Left =165
                    Top =855
                    Width =480
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiIMPX_Date_Fin"
                    Caption ="au  :"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =120
                    Top =1410
                    Width =1560
                    Height =405
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiIMPX_Nb_Créances"
                    Caption ="Nombre de créances"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =135
                    Top =2520
                    Width =1560
                    Height =405
                    FontWeight =700
                    BackColor =12632256
                    Name ="EtiIMPX_Montant_m"
                    Caption ="Somme des créances"
                    FontName ="MS Sans Serif"
                End
            End
        End
        Begin FormFooter
            Height =765
            BackColor =12632256
            Name ="PiedFormulaire"
            Begin
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =73
                    Left =6030
                    Top =180
                    Width =2190
                    Height =450
                    FontSize =10
                    FontWeight =300
                    ForeColor =0
                    Name ="CmdImporter"
                    Caption ="&Importer un lot"
                    OnClick ="[Event Procedure]"
                    FontName ="System"

                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =86
                    Left =3690
                    Top =180
                    Width =2190
                    Height =450
                    FontSize =10
                    FontWeight =300
                    TabIndex =1
                    ForeColor =0
                    Name ="CmdVisualiser"
                    Caption ="&Visualiser un lot"
                    OnClick ="[Event Procedure]"
                    FontName ="System"

                End
            End
        End
    End
End
CodeBehindForm
' See "FrmImportationXLS.cls"
