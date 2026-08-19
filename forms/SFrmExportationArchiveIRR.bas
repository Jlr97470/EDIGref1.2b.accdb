Version =20
VersionRequired =20
Begin Form
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
    Width =3855
    ItemSuffix =53
    Left =4395
    Top =45
    Right =8595
    Bottom =4155
    RecSrcDt = Begin
        0x6b40700d55ede240
    End
    RecordSource ="SELECT DISTINCTROW EXPORTATION_IRR.*, EXPORTATION_IRR.EXPI_Date FROM EXPORTATION"
        "_IRR ORDER BY EXPORTATION_IRR.EXPI_Date DESC; "
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
            Height =240
            BackColor =12632256
            Name ="EntêteFormulaire1"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Width =855
                    Height =240
                    FontWeight =400
                    Name ="EtiEXPI_Lot_Date"
                    Caption ="Date"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =855
                    Width =300
                    Height =240
                    FontWeight =400
                    Name ="EtiEXPI_Lot_Num"
                    Caption ="N°"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2475
                    Width =615
                    Height =240
                    FontWeight =400
                    Name ="EtiEXPI_Lot_Contrat"
                    Caption ="Contrat"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1170
                    Width =645
                    Height =240
                    FontWeight =400
                    Name ="EtiEXPI_Lot_Site"
                    Caption ="Site"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1815
                    Width =675
                    Height =240
                    FontWeight =400
                    Name ="EtiEXPI_Lot_CodeGREF"
                    Caption ="Gref"
                End
            End
        End
        Begin Section
            Height =300
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =30
                    Top =30
                    Width =870
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtEXPI_Lot_Date"
                    ControlSource ="EXPI_Lot_Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    Left =915
                    Top =30
                    Width =270
                    TabIndex =1
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtEXPI_Lot_Num"
                    ControlSource ="EXPI_Lot_Num"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    Left =1185
                    Top =30
                    Width =630
                    TabIndex =2
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtEXPI_Lot_Site"
                    ControlSource ="EXPI_Lot_Site"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    TextAlign =2
                    Left =1815
                    Top =30
                    Width =615
                    TabIndex =3
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtEXPI_Lot_CodeGREF"
                    ControlSource ="EXPI_Lot_CodeGREF"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =87
                    TextAlign =2
                    Left =2430
                    Top =30
                    Width =615
                    TabIndex =4
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtEXPI_Lot_Contrat"
                    ControlSource ="EXPI_Lot_Contrat"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3390
                    Width =393
                    Height =300
                    FontSize =8
                    TabIndex =5
                    Name ="CmdVisualiser"
                    OnClick ="=ImpVisu(2)"
                    PictureData = Begin
                        0x28000000100000000f0000000100040000000000780000000000000000000000 ,
                        0x1000000010000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff0000000000000088000ffffffffff080000fffffff00007008 ,
                        0x0ffffff0788707880fffff0788e770880fffff08888780880fffff08e8878088 ,
                        0x0fffff07ee8770880ffffff0788708880fffffff000088880ffffffffff08888 ,
                        0x0fffffff000088880fffffff080888880fffffff008888880000000008888888 ,
                        0x0000000000000000
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3060
                    Width =333
                    Height =300
                    FontSize =8
                    TabIndex =6
                    Name ="CmdDetail"
                    Caption ="..."
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "SFrmExportationArchiveIRR.cls"
