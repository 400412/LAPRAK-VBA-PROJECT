Attribute VB_Name = "Macro_Module"
Sub Data_Majemuk()
Attribute Data_Majemuk.VB_ProcData.VB_Invoke_Func = "O\n14"
'
' Data_Majemuk Macro
'
' Keyboard Shortcut: Ctrl+Shift+O
'
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]^2"
    ActiveCell.Offset(1, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = ChrW(&H3A3)
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    ActiveCell.Offset(0, 1).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-5]C:R[-1]C)"
    ActiveCell.Offset(1, -2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Rata-rata"
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=AVERAGE(R[-6]C:R[-2]C)"
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = ChrW(&H394)
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=DELTA_5(R[-2]C,R[-2]C[1])"
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "KSR (%)"
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=ksr(R[-2]C,R[-1]C)"
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Angka Penting"
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=angka_penting(R[-1]C)"
    ActiveCell.Offset(1, -1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "Hasil Akhir"
    ActiveCell.Offset(0, 1).Range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Application.CutCopyMode = False
    ActiveCell.Formula2R1C1 = "=HASIL(R[-4]C,R[-3]C,R[-1]C)"
    ActiveCell.Offset(-11, -1).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Font
        .Name = "New Era Casual"
        .Size = 12
        .Strikethrough = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveCell.Offset(6, 0).Range("A1:A6").Select
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    ActiveCell.Offset(0, 1).Range("A1:B6").Select
    With Selection.Font
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.249977111117893
    End With
    ActiveCell.Offset(0, -1).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1:C1").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


