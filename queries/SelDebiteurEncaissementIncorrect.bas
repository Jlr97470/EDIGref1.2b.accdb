Operation =1
Option =0
Where ="(((DEBITEUR.Déb_Situation)=\"99\") AND (((SELECT Round(Nz(Sum(Enc_Prin_m1),0),2)"
    " FROM ENCAISSEMENT WHERE Enc_Code_Débiteur=Déb_Code;))<Round([Déb_Mont_Prin_m1],"
    "2)) AND (((SELECT Round(Nz(Sum(Enc_Prin_m2),0),2) FROM ENCAISSEMENT WHERE Enc_Co"
    "de_Débiteur=Déb_Code;))<Round([Déb_Mont_Prin_m2],2)))"
Begin InputTables
    Name ="DEBITEUR"
End
Begin OutputColumns
    Expression ="DEBITEUR.Déb_Code"
    Expression ="DEBITEUR.Déb_Situation"
    Expression ="DEBITEUR.Déb_Date_Cré"
    Expression ="DEBITEUR.Déb_Mont_Prin_m1"
    Expression ="DEBITEUR.Déb_Mont_Prin_m2"
    Expression ="DEBITEUR.Déb_MonnaieEnCours"
    Alias ="Déb_Enc_Prin_m1"
    Expression ="(SELECT Round(Nz(Sum(Enc_Prin_m1),0),2) FROM ENCAISSEMENT WHERE Enc_Code_Débiteu"
        "r=Déb_Code;)"
    Alias ="Déb_Enc_Prin_m2"
    Expression ="(SELECT Round(Nz(Sum(Enc_Prin_m2),0),2) FROM ENCAISSEMENT WHERE Enc_Code_Débiteu"
        "r=Déb_Code;)"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="DEBITEUR.Déb_Code"
        dbInteger "ColumnWidth" ="1230"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DEBITEUR.Déb_Mont_Prin_m1"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="DEBITEUR.Déb_Mont_Prin_m2"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Déb_Enc_Prin_m1"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Déb_Enc_Prin_m2"
        dbInteger "ColumnWidth" ="1800"
        dbBoolean "ColumnHidden" ="0"
    End
End
Begin
    State =0
    Left =-29
    Top =112
    Right =889
    Bottom =516
    Left =-1
    Top =-1
    Right =911
    Bottom =228
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =202
        Bottom =195
        Top =6
        Name ="DEBITEUR"
        Name =""
    End
End
