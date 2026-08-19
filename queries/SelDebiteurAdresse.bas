Operation =1
Option =0
Where ="(((ADRESSE_DEBITEUR.AD_TypeAdresse_ID)=DLookUp(\"Adresse_Courrier\",\"PARAMETRE_"
    "GREF\")))"
Begin InputTables
    Name ="DEBITEUR"
    Name ="ADRESSE_DEBITEUR"
End
Begin OutputColumns
    Expression ="DEBITEUR.*"
    Expression ="ADRESSE_DEBITEUR.*"
End
Begin Joins
    LeftTable ="DEBITEUR"
    RightTable ="ADRESSE_DEBITEUR"
    Expression ="DEBITEUR.Déb_Code=ADRESSE_DEBITEUR.AD_Code_Débiteur"
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
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="DEBITEUR"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="ADRESSE_DEBITEUR"
        Name =""
    End
End
