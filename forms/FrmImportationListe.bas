Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    BorderStyle =1
    GridX =5
    GridY =5
    Width =4260
    ItemSuffix =39
    Left =5385
    Top =300
    Right =10005
    Bottom =6195
    RecSrcDt = Begin
        0x9d21564c4eede240
    End
    RecordSource ="SELECT IMPORTATION.* FROM IMPORTATION ORDER BY IMPORTATION.IMP_Date DESC; "
    Caption ="Liste des lots importés"
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =-15
                    Top =840
                    Width =975
                    Height =240
                    FontWeight =400
                    Name ="EtiIMP_Date"
                    Caption ="Date"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =975
                    Top =840
                    Width =330
                    Height =240
                    FontWeight =400
                    Name ="EtiIMP_Lot_Num"
                    Caption ="N°"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =2670
                    Top =840
                    Width =645
                    Height =240
                    FontWeight =400
                    Name ="EtiIMP_Lot_Contrat"
                    Caption ="Contrat"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =1305
                    Top =840
                    Width =675
                    Height =240
                    FontWeight =400
                    Name ="EtiIMP_Lot_Site"
                    Caption ="Site"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =1980
                    Top =840
                    Width =705
                    Height =240
                    FontWeight =400
                    Name ="EtiIMP_Lot_CodeGREF"
                    Caption ="Gref"
                End
            End
        End
        Begin Section
            Height =276
            BackColor =12632256
            Name ="Détail0"
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Width =975
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtIMP_Lot_Date"
                    ControlSource ="IMP_Lot_Date"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1020
                    Width =285
                    TabIndex =1
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtIMP_Lot_Num"
                    ControlSource ="IMP_Lot_Num"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =1350
                    Width =630
                    TabIndex =2
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtIMP_Lot_Site"
                    ControlSource ="IMP_Lot_Site"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2025
                    Width =630
                    TabIndex =3
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtIMP_Lot_CodeGREF"
                    ControlSource ="IMP_Lot_CodeGREF"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    Left =2700
                    Width =630
                    TabIndex =4
                    BackColor =255
                    ForeColor =16777215
                    Name ="TxtIMP_Lot_Contrat"
                    ControlSource ="IMP_Lot_Contrat"
                End
                Begin CommandButton
                    OverlapFlags =85
                    TextFontFamily =34
                    Left =3435
                    Top =30
                    Width =246
                    Height =231
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
                        0x0fffffff000088880fffffff080888880fffffff008888880000000008888888
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
                    Left =3690
                    Top =30
                    Width =246
                    Height =246
                    FontSize =8
                    TabIndex =6
                    Name ="CmdImprimer"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="=ImpVisu(0)"
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
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff00000000ff0000ffff00000000ff00ff00ff0000ffff00 ,
                        0xffffff0000000000
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    TextFontFamily =34
                    Left =3960
                    Top =30
                    Width =246
                    Height =246
                    FontSize =8
                    TabIndex =7
                    Name ="CmdExporter"
                    Caption ="&i"
                    StatusBarText ="Imprimer le rapport"
                    OnClick ="=ImpVisu(5)"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000add0330880000330da ,
                        0xa0330880000330add0330880000330daa0330000000330add0333333333330da ,
                        0xa0330000000330add030fffffff030daa030fffcfff030add030ffcccff030da ,
                        0xa030fffcccfdadadd030ffffcccacadaa000000daccccdaddadadadadacccada ,
                        0xadadadadaccccdad
                    End
                    FontName ="Arial"
                    ObjectPalette = Begin
                        0x0003100000000000800000000080000080800000000080008000800000808000 ,
                        0x80808000c0c0c000ff000000c0c0c000ffff00000000ff00c0c0c00000ffff00 ,
                        0xffffff0000000000
                    End
                End
            End
        End
        Begin FormFooter
            Height =226
            BackColor =12632256
            Name ="PiedFormulaire2"
        End
    End
End
CodeBehindForm
' See "FrmImportationListe.cls"
