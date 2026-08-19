Operation =1
Option =0
Begin InputTables
    Name ="REGLEMENT"
    Name ="REGLEMENT_DETAIL"
End
Begin OutputColumns
    Expression ="REGLEMENT.*"
    Expression ="REGLEMENT_DETAIL.*"
End
Begin Joins
    LeftTable ="REGLEMENT"
    RightTable ="REGLEMENT_DETAIL"
    Expression ="REGLEMENT.Reg_Id=REGLEMENT_DETAIL.RegD_Reg_Id"
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="REGLEMENT"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="REGLEMENT_DETAIL"
        Name =""
    End
End
