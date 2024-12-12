Attribute VB_Name = "RechecheV"

Function RECHERCHEV(Valeur_Cherchee As Variant, Table_matrice As Range, No_index_col As Single, Optional Valeur_proche As Boolean)
'par Excel-Malin.com ( https://excel-malin.com/ )

On Error GoTo RECHERCHEVerror
    RECHERCHEV = Application.VLookup(Valeur_Cherchee, Table_matrice, No_index_col, Valeur_proche)
    If IsError(RECHERCHEV) Then RECHERCHEV = 0
    
Exit Function
RECHERCHEVerror:
    RECHERCHEV = 0
End Function



