Operation =1
Option =0
Begin InputTables
    Name ="CODE_GESTION"
    Name ="DEBITEUR"
    Name ="ENCAISSEMENT"
End
Begin OutputColumns
    Expression ="DEBITEUR.*"
    Expression ="CODE_GESTION.*"
    Expression ="ENCAISSEMENT.*"
End
Begin Joins
    LeftTable ="DEBITEUR"
    RightTable ="ENCAISSEMENT"
    Expression ="DEBITEUR.Déb_Code=ENCAISSEMENT.Enc_Code_Débiteur"
    Flag =1
    LeftTable ="CODE_GESTION"
    RightTable ="DEBITEUR"
    Expression ="CODE_GESTION.CG_Code=DEBITEUR.Déb_Situation"
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
    Left =105
    Top =127
    Right =1002
    Bottom =447
    Left =-1
    Top =-1
    Right =890
    Bottom =144
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="CODE_GESTION"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="DEBITEUR"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="ENCAISSEMENT"
        Name =""
    End
End
