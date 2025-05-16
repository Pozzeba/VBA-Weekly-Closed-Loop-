# VBA-Weekly-Closed-Loop-
Format the weekly closed loop
Sub ClosedLoop()
'
' ClosedLoop Macro
' Ajuste da planilha semanal de closed loop
'

'
Dim n As Long
n = Range("A1").CurrentRegion.Rows.Count

    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
    ActiveCell.FormulaR1C1 = "=WEEKNUM(RC[1],21)"
    Range("A2").Select
    Selection.AutoFill Range("A2:A" & n)
    Range("A1").FormulaR1C1 = "Week"
    Range("B1").FormulaR1C1 = "BR - Collection sucess time"
    Range("C1").FormulaR1C1 = "Provider name"
    Range("D1").FormulaR1C1 = "Channel Name"
    Range("E1").FormulaR1C1 = "Warehouse_cn"
    Range("F1").FormulaR1C1 = "Warehouse_en"
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("G1").FormulaR1C1 = "Seller ID"
    Range("H1").FormulaR1C1 = "Provider Waybill number"
    Range("I1").FormulaR1C1 = "Subpackage number"
    Range("J1").FormulaR1C1 = "Receipt number"
    Range("K1").FormulaR1C1 = "Order number"
    Columns("L:L").Select
    Selection.Delete Shift:=xlToLeft
    ActiveWindow.SmallScroll ToRight:=4
    Columns("M:M").Select
    Selection.Delete Shift:=xlToLeft
    Range("L1").FormulaR1C1 = "The collection track sent by provider"
    Range("M1").FormulaR1C1 = "Status"
    Range("N1").FormulaR1C1 = "Parcel Status"
    Range("O1").FormulaR1C1 = "Package Value (BRL)"
    Range("P1").FormulaR1C1 = "Freight (BRL)"
    Range("Q1").FormulaR1C1 = "Collection Driver"
    Range("R1").FormulaR1C1 = "License Plate"
    Range("S1").FormulaR1C1 = "Transfer Driver"
    Range("T1").FormulaR1C1 = "License Plate"
    Range("U1").FormulaR1C1 = "Track Status update"
    Range("V1").FormulaR1C1 = "Closed Loop Confirmation Status"


    Range("F2").FormulaR1C1 = _
        "=VLOOKUP(RC[-1],'[support file.xlsx]Sheet1'!R2C1:R10C2,2,FALSE)"
Range("F2").AutoFill Range("F2:F" & n)
 
    Range("A1:T1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Range("U1:V1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Range("V1").Select
    Selection.AutoFilter
    
        Range("L2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("O2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlToLeft
    Range("F2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.End(xlUp).Select
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

