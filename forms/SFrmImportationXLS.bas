Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    TabularFamily =1
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =9645
    DatasheetFontHeight =10
    ItemSuffix =7
    Left =3630
    Top =2505
    Right =13230
    Bottom =5940
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x110187a070ebe240
    End
    DatasheetFontName ="Arial"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin FormHeader
            Height =240
            BackColor =12632256
            Name ="EntêteFormulaire"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =30
                    Width =1710
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiRef Client"
                    Caption ="Ref Client"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="EtiRef_Client"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =1785
                    Width =1680
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiNom Prénom"
                    Caption ="Nom Prénom"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="EtiNom_Prénom"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =3495
                    Width =1455
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiCode Postal"
                    Caption ="Code Postal"
                    FontName ="MS Sans Serif"
                    EventProcPrefix ="EtiCode_Postal"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =4995
                    Width =1545
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiVille"
                    Caption ="Ville"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =6570
                    Width =1125
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiDate"
                    Caption ="Date"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    BackStyle =1
                    OverlapFlags =85
                    TextAlign =1
                    Left =7725
                    Width =1095
                    Height =240
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =16777215
                    Name ="EtiMontant"
                    Caption ="Montant"
                    FontName ="MS Sans Serif"
                End
            End
        End
        Begin Section
            Visible = NotDefault
            Height =345
            BackColor =12632256
            Name ="Détail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =30
                    Top =30
                    Name ="TxtRef Client"
                    ControlSource ="Ref Client"
                    Format ="@"
                    EventProcPrefix ="TxtRef_Client"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1770
                    Top =31
                    TabIndex =1
                    Name ="TxtNom Prénom"
                    ControlSource ="Nom Prénom"
                    EventProcPrefix ="TxtNom_Prénom"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3495
                    Top =31
                    Width =1461
                    TabIndex =2
                    Name ="TxtCode Postal"
                    ControlSource ="Code Postal"
                    EventProcPrefix ="TxtCode_Postal"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4995
                    Top =31
                    Width =1536
                    TabIndex =3
                    Name ="TxtVille"
                    ControlSource ="Ville"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7725
                    Top =31
                    Width =1101
                    TabIndex =4
                    Name ="TxtMontant"
                    ControlSource ="=Replace([Montant],DonnerSeparateur(1),DonnerSeparateur(0))"
                    Format ="General Number"
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6570
                    Top =31
                    Width =1131
                    TabIndex =5
                    Name ="TxtDate"
                    ControlSource ="Date"
                    Format ="Short Date"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =9315
                    Top =45
                    Width =320
                    Height =245
                    FontSize =6
                    TabIndex =6
                    ForeColor =0
                    Name ="CmdOk"
                    Caption ="OK"
                    OnClick ="[Event Procedure]"
                    FontName ="Small Fonts"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8895
                    Top =75
                    Width =187
                    Height =187
                    TabIndex =7
                    Name ="ChkTIX_Flag"
                    ControlSource ="=Not IsNull(DLookUp(\"[TIX_Ref Client]\",\"TMP_IMP_XLS\",\"[TIX_Ref Client]='\" "
                        "& [Ref Client] & \"' AND [TIX_N°]='\" & [N°] & \"'\"))"
                    DefaultValue ="Yes"
                    OnMouseDown ="[Event Procedure]"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =85
                    BorderWidth =3
                    Left =9120
                    Top =75
                    Width =187
                    Height =187
                    TabIndex =8
                    Name ="ChkDbl"
                    ControlSource ="=Not IsNull(DLookUp(\"Déb_Code\",\"DEBITEUR\",\"Déb_Code_Client='\" & Forms!FrmI"
                        "mportationXLS!TxtIMPX_Cli_Code & \"' AND Déb_Réf1_Client='\" & [Ref Client] & \""
                        "'\"))"
                    StatusBarText ="Relance Deja Editer"
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
' See "SFrmImportationXLS.cls"
