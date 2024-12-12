Attribute VB_Name = "Ech_complet"
Sub Ech_complet()

'Renommer feuille

ActiveSheet.Name = "Data"


'Libellé masque des données

Worksheets("Data").Range("A1").Value = "Nom Client"
Worksheets("Data").Range("A2").Value = "Montant"
Worksheets("Data").Range("A3").Value = "Taux"
Worksheets("Data").Range("A4").Value = "Date Début"
Worksheets("Data").Range("A5").Value = "Durée (Mois)"
Worksheets("Data").Range("A6").Value = "Fréquence remboursement"
Worksheets("Data").Range("A7").Value = "Type remboursement"
Worksheets("Data").Range("A8").Value = "TRI"


'Libellé masque des données complémentaires générées

Worksheets("Data").Range("d4").Value = "Donc fin en :"
Worksheets("Data").Range("d5").Value = "Nbre Ech Tot"
Worksheets("Data").Range("e5").Value = "Nbre Ech / an"


'Libellé masque tableau Echéancier
Worksheets("Data").Range("a10").Value = "# Echéance"
Worksheets("Data").Range("b10").Value = "Date Echéance"
Worksheets("Data").Range("c10").Value = "Capital Restant"
Worksheets("Data").Range("d10").Value = "Mon_Capital"
Worksheets("Data").Range("e10").Value = "Mon_Intérêts"
Worksheets("Data").Range("f10").Value = "Mon_Echéance"
Worksheets("Data").Range("g10").Value = "KRD Fin"
   
'Mise en Forme
Worksheets("Data").Range("A1:A9").Font.Bold = True

Worksheets("Data").Range("e4").Font.Color = vbRed
Worksheets("Data").Range("d5:e5").Font.Italic = True
    With Worksheets("Data").Range("d6:e6")
    .Font.Italic = True
    .Font.Color = vbRed
    End With
    
    With Worksheets("Data").Range("A10:g10")
    .Font.Bold = True
    .Interior.ColorIndex = 16
    End With
    

'Définition des variables

    'Définition variables Inputs
    
    Dim clt As String
    Dim Mon As Double
    Dim Tx As Double
    Dim DateSt As Date
    Dim Dur As Integer
    Dim Freq As Integer
    Dim TYP As String
    
    'Définition variables Compl calculées
    
    Dim NBEch As Integer 'Nombre échéances total
    Dim NbA As Integer 'Nombre Echéances par an
    Dim DateF As Date 'Date dernière échéance
    
    'Définition variables Echéancier
    
    Dim NumEch As Integer
    Dim DateEch As Date
    Dim KRDB As Double
    Dim K As Double
    Dim Ints As Double
    Dim MonEch As Double
    Dim KRDF As Double
    
    Dim lastcell As Range
    Dim TotK As String
    Dim TotInts As String
    Dim TotMonEch As String

    Dim TRI As Double
    
    'Définition variables Calculées pour Encours Moyens
    Dim DateMIN As Date '1er JAN de l'annès déblocage
    Dim DateMAX As Date '31 DEC de l'annès dernier remboursement
    Dim DateScope As Integer 'Nombre de jours entre MAX et MIN
    Dim DateSeq As Date 'Date de chaque ligne
    Dim SEQ As Integer '#
    Dim SEQx As Integer 'c'est l'équivalent séquence pour l'échéancier
    Dim CltSEQ As String
    Dim KMoyD As Double
    Dim KMoyRem As Double
    Dim KMoyFin As Double
    Dim IntsMoy As Double
    Dim RenMoy As Double
        
    
' INPUTS
    'Valeur INPUTS
    clt = "Sirine "
    Mon = 1000000
    Tx = 0.05
    DateSt = #2/15/2022#
    Dur = 12 'En mois'
    Freq = 2   'chaque x mois
    TYP = "AC"  ' "AC" Annuités constantes ; "KC" Capital constant

    'Affect INPUTS
    Range("b1").Value = clt
    Range("b2").Value = Mon
    Range("b3").Value = Tx
    Range("b4").Value = DateSt
    Range("b5").Value = Dur
    Range("b6").Value = Freq
    Range("b7").Value = TYP


'Calcul variables Complémentaires
    'Valeurs Var Compl
    NBEch = Dur / Freq
    NbA = 12 / Freq
    DateF = DateAdd("m", Dur, DateSt)
    
    'Affect Var compl
    Range("e4").Value = DateF
    Range("d6").Value = NBEch
    Range("e6").Value = NbA

    'Calcul variables EncoursMoy

    DateMIN = DateSerial(Year(DateSt), 1, 1)
    DateMAX = DateSerial(Year(DateF), 12, 31)
    DateScope = DateDiff("d", DateMIN, DateMAX) + 1




'Calcul Echéancier

    'figer numero Echéance 1
    NumEch = Range("A11").Row - 10
    
    'Loop numéro échéance (à partir de la ligne #11) jusqu'à Nombre Echéances
    
    For NumEch = 1 To (NBEch)
    Cells(NumEch + 10, 2).Value = NumEch

    
        'Calcul Date Echéance
        DateEch = DateAdd("m", (NumEch * Freq), DateSt)
        
        'Calcul KRDB
        
            'Condition KRDB
            If NumEch = 1 Then
            KRDB = Mon
            Else
            KRDB = KRDF
            End If
        
        'Calcul Intérêt
        Ints = KRDB * Tx / NbA
        
        'Condition Type de remboursement
        
            'Hyp : Capital constant
            
            If TYP = "KC" Then
                
                'Calcul Capital Echéance
                K = Mon / NBEch
                
                'Calcul Montant Echéance
                MonEch = K + Ints
                
            'Hyp : Echéance constante
            Else: TYP = "AC"
                
                'Calcul Montant Echéance
                MonEch = Pmt((Tx / NbA), NBEch, -Mon)
                                
                'Calcul Capital Echéance
                K = MonEch - Ints
                
            End If
            
            'Calcul KRD fin
            KRDF = KRDB - K
            
            'Calcul Seq (utile pour le calcul des EncoursMoy
            SEQx = NumEch + DateDiff("d", DateMIN, DateEch)
            
    'Affecter valeurs échéancier
    
    Cells(NumEch + 10, 1).Value = NumEch
    Cells(NumEch + 10, 2).Value = DateEch
    Cells(NumEch + 10, 3).Value = KRDB
    Cells(NumEch + 10, 4).Value = K
    Cells(NumEch + 10, 5).Value = Ints
    Cells(NumEch + 10, 6).Value = MonEch
    Cells(NumEch + 10, 7).Value = KRDF
    Cells(NumEch + 10, 8).Value = SEQx
    
    Next NumEch
        
'Calcul des Totaux

    'Total Capital Remboursé
    Set lastcell = Range("D11").End(xlDown)
    lastcell.Select
    ActiveCell.Offset(2).Select
    TotK = "=sum(E11:" & lastcell.Address(False, False) & ")"
    ActiveCell.Formula = TotK
    ActiveCell.Font.Bold = True 'mise en forme Gras


    'Total Intérêts Remboursés
    Set lastcell = Range("E11").End(xlDown)
    lastcell.Select
    ActiveCell.Offset(2).Select
    TotInts = "=sum(f11:" & lastcell.Address(False, False) & ")"
    ActiveCell.Formula = TotInts
    ActiveCell.Font.Bold = True 'mise en forme Gras

    'Total Echéances Remboursées
    Set lastcell = Range("F11").End(xlDown)
    lastcell.Select
    ActiveCell.Offset(2).Select
    TotEch = "=sum(g11:" & lastcell.Address(False, False) & ")"
    ActiveCell.Formula = TotEch
    ActiveCell.Font.Bold = True 'mise en forme Gras


'Mise en forme nombres
Range("b2").NumberFormatLocal = "# ##0,00"
Range("b3").NumberFormatLocal = "0,00%"
Range(Cells(11, 4), Cells((NBEch + 10 + 2), 8)).NumberFormatLocal = "# ##0,00"

'Calcul TRI

    'Récupérer les valeurs de flux en colonne M
        'Copier/Coller valeurs Cash Out (-Mon)
        Range("Z1").Value = -Mon
        'Copier/Coller valeurs des remboursements
        Range("F11", Cells(NBEch + 10, 6)).Copy Range("Z2")
    
    'Insérer formule TRI (pas trouvé solution à utiliser directement formule IRR)
    Range("Y1").FormulaArray = "=TRI(Z:Z,0.1)"
    Range("Y1").Select
    ActiveCell.FormulaR1C1 = "=IRR(C[1],0.1)" 'enlever souci #NOM
    Range("Y2").Select
    
    'Affecter valeur TRI à Cellule
    Range("B8").Value = Range("y1").Value * NbA
    Range("B8").Font.Color = vbRed
    Range("b8").NumberFormatLocal = "0,00%"
    'Affecter valeur TRI à variable
    TRI = Range("B8").Value

    'Supprimer Données TRI (colonnes L et M)
    Columns("y:z").Select
    Selection.ClearContents


'Accessoires
Worksheets("Data").Columns("A:H").EntireColumn.AutoFit
Range("a1").Select

'Encours Moyens
    'Libellé masque tableau Encours Moyens
    Worksheets("Data").Range("K2").Value = "Client"
    Worksheets("Data").Range("L2").Value = "Seq"
    Worksheets("Data").Range("M2").Value = "Date"
    Worksheets("Data").Range("N2").Value = "KRD Déb"
    Worksheets("Data").Range("O2").Value = "Capital Remboursé"
    Worksheets("Data").Range("P2").Value = "KRD Fin"
    Worksheets("Data").Range("Q2").Value = "Intérêts"
    Worksheets("Data").Range("R2").Value = "Rendement"
       
    'Mise en Forme
    
    Worksheets("Data").Range("K1:U1").Font.Italic = True
       
        With Worksheets("Data").Range("K2:S2")
        .Font.Bold = True
        .Interior.ColorIndex = 16
        End With
        
    
    'Définition des variables entête pour mémoire
    Range("K1").Value = DateMIN
    Range("L1").Value = DateMAX
    Range("M1").Value = DateScope

    'Loop séquence
    
    For SEQ = 1 To DateScope
    
    'Date Sequence
    DateSeq = DateMIN + SEQ - 1
    CltSEQ = clt
        
    'Récupérer le K remboursé sur base de VLOOKUP et IF
    
        'Copier le KRD from Echéancier pour créer plage de recherche
        Range("H11", Cells(NBEch + 10, 8)).Copy Range("AA1")
        Range("C11", Cells(NBEch + 10, 3)).Copy Range("AB1")
        Range("D11", Cells(NBEch + 10, 4)).Copy Range("AC1")
        Range("G11", Cells(NBEch + 10, 7)).Copy Range("AD1")
    
        'Récupérer remboursement Capital KMoy
        Kmoy = RECHERCHEV(SEQ, Range("AA1:AD120"), 3)
        
    'Récupérer KRD départ pour EncoursMoyD
        'sur base d'un recherchevnume séquenciel et montant
        Range("AF1") = DateDiff("d", DateMIN, DateSt) + 1
        
        'récupérer valeur montat origine
        Range("AG1") = Mon
        
        'lancer rechercheV si jour déblocage = Montant sinon KRDfin
        
        KMoyD = RECHERCHEV(SEQ, Range("AF1:AG1"), 2)
        
        If KMoyD = 0 Then
        KMoyD = KMoyFin
        End If
        
    'Récupérer KRD Fin de période KMoyFin
    KMoyFin = KMoyD - Kmoy
    
    'Récupérer Intérêt par jourde période KMoyFin
    IntsMoy = KMoyD * TRI / 360
    
    
    'Affectation des variables
    Cells(SEQ + 2, 11).Value = CltSEQ
    Cells(SEQ + 2, 12).Value = SEQ
    Cells(SEQ + 2, 13).Value = DateSeq
    Cells(SEQ + 2, 14).Value = KMoyD
    Cells(SEQ + 2, 15).Value = Kmoy
    Cells(SEQ + 2, 16).Value = KMoyFin
    Cells(SEQ + 2, 17).Value = IntsMoy
    Cells(SEQ + 2, 18).Value = RenMoy
    
    
    Next SEQ

'Mise en forme nombres
Range("N:Q").NumberFormatLocal = "# ##0,00"
Range("R:R").NumberFormatLocal = "0,00%"





End Sub
