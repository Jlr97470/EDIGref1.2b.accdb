Operation =1
Option =0
Begin InputTables
    Name ="CLIENT"
    Name ="EXPORTATION_IRR"
    Name ="DEBITEUR"
    Name ="EXPORTATION_IRR_DETAIL"
End
Begin OutputColumns
    Expression ="EXPORTATION_IRR.*"
    Expression ="EXPORTATION_IRR_DETAIL.*"
    Expression ="CLIENT.*"
    Expression ="DEBITEUR.*"
End
Begin Joins
    LeftTable ="DEBITEUR"
    RightTable ="EXPORTATION_IRR_DETAIL"
    Expression ="DEBITEUR.Déb_Lot_EID_Id=EXPORTATION_IRR_DETAIL.EID_Id"
    Flag =1
    LeftTable ="EXPORTATION_IRR"
    RightTable ="EXPORTATION_IRR_DETAIL"
    Expression ="EXPORTATION_IRR.EXPI_Id=EXPORTATION_IRR_DETAIL.EID_Code"
    Flag =1
    LeftTable ="CLIENT"
    RightTable ="DEBITEUR"
    Expression ="CLIENT.Cli_Code=DEBITEUR.Déb_Code_Client"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
End
Begin
    State =0
    Left =47
    Top =69
    Right =1002
    Bottom =389
    Left =-1
    Top =-1
    Right =948
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =440
        Top =6
        Right =536
        Bottom =120
        Top =0
        Name ="EXPORTATION_IRR"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="CLIENT"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="DEBITEUR"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="EXPORTATION_IRR_DETAIL"
        Name =""
    End
End
