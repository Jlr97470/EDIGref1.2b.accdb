Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =0
    GridX =5
    GridY =5
    Width =6936
    ItemSuffix =55
    Left =5160
    Top =2625
    Right =12120
    Bottom =6270
    RecSrcDt = Begin
        0xeca9fd9dd3ebe240
    End
    AllowDatasheetView =0
    Begin
        Begin Label
            TextAlign =3
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
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
        End
        Begin CheckBox
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-236
        End
        Begin TextBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin ListBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =1417
            LabelX =-236
        End
        Begin ComboBox
            SpecialEffect =2
            LabelAlign =3
            Width =1701
            Height =255
            LabelX =-236
        End
        Begin FormHeader
            Height =232
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1590
                    Width =1170
                    Height =210
                    FontSize =7
                    Name ="EtiImp_Code_Etablissement"
                    Caption ="Etablissement"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5025
                    Width =630
                    Height =225
                    FontSize =7
                    Name ="EtiImp_NbCreance"
                    Caption ="Nombre"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5670
                    Width =1260
                    Height =225
                    FontSize =7
                    Name ="EtiImp_Montant"
                    Caption ="Montant"
                    FontName ="Small Fonts"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =645
                    Width =540
                    Height =210
                    FontSize =7
                    Name ="EtiCli_Code"
                    Caption ="Client"
                    FontName ="Small Fonts"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =270
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =49
                    Left =1585
                    Top =30
                    Width =876
                    Height =225
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtImp_Code_Etablissement"
                    ControlSource ="Imp_Code_Etablissement"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    TextFontFamily =49
                    Left =2515
                    Top =30
                    Width =2526
                    Height =225
                    TabIndex =1
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtImp_Nom_Etablissement"
                    ControlSource ="Imp_Nom_Etablissement"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =49
                    Left =5085
                    Top =30
                    Width =561
                    Height =225
                    TabIndex =2
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtImp_NbCreance"
                    ControlSource ="Imp_NbCreance"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =49
                    Left =5685
                    Top =30
                    Width =1251
                    Height =225
                    TabIndex =3
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtImp_Montant"
                    ControlSource ="=Switch([Cli_MonnaieEnCours]=1,Format([Imp_Montant],\"#,##0.00\"\" F\"\"\"),[Cli"
                        "_MonnaieEnCours]=2,Format([Imp_Montant],\"#,##0.00 €\"),True,Null)"
                    FontName ="Courier New"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =49
                    Left =685
                    Top =30
                    Width =606
                    Height =225
                    TabIndex =4
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCli_Code"
                    ControlSource ="Cli_Code"
                    FontName ="Courier New"
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =1
                    TextFontFamily =2
                    Left =1300
                    Top =30
                    Width =240
                    Height =240
                    FontSize =11
                    ForeColor =255
                    Name ="EtiCorrespondance"
                    Caption ="ð"
                    FontName ="Wingdings"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =49
                    Left =30
                    Top =30
                    Width =636
                    Height =225
                    TabIndex =5
                    BackColor =8388608
                    ForeColor =16777215
                    Name ="TxtCli_Code_Groupe"
                    ControlSource ="Cli_Code_Groupe"
                    FontName ="Courier New"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire"
        End
    End
End
CodeBehindForm
' See "SFrmImportation.cls"
