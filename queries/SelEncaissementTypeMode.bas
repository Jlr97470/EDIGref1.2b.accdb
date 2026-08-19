Operation =1
Option =0
Begin InputTables
    Name ="ENCAISSEMENT_MODE"
    Name ="ENCAISSEMENT_TYPE"
    Name ="ENCAISSEMENT"
End
Begin OutputColumns
    Expression ="ENCAISSEMENT.*"
    Expression ="ENCAISSEMENT_TYPE.*"
    Expression ="ENCAISSEMENT_MODE.*"
End
Begin Joins
    LeftTable ="ENCAISSEMENT_MODE"
    RightTable ="ENCAISSEMENT_TYPE"
    Expression ="ENCAISSEMENT_MODE.EMo_Code=ENCAISSEMENT_TYPE.ETy_Mode"
    Flag =1
    LeftTable ="ENCAISSEMENT_TYPE"
    RightTable ="ENCAISSEMENT"
    Expression ="ENCAISSEMENT_TYPE.ETy_Code=ENCAISSEMENT.Enc_Type"
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
    Left =76
    Top =98
    Right =1002
    Bottom =418
    Left =-1
    Top =-1
    Right =919
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =454
        Top =7
        Right =606
        Bottom =91
        Top =0
        Name ="ENCAISSEMENT_MODE"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =105
        Top =0
        Name ="ENCAISSEMENT_TYPE"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="ENCAISSEMENT"
        Name =""
    End
End
