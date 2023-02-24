Attribute VB_Name = "Module2"
Sub EBIT_CommonSize()
Attribute EBIT_CommonSize.VB_Description = "Create EBIT and Common size income statement\n"
Attribute EBIT_CommonSize.VB_ProcData.VB_Invoke_Func = "I\n14"
'
' EBIT_CommonSize Macro
' Create EBIT and Common size income statement
'
' Keyboard Shortcut: Ctrl+Shift+I
'
    Sheets.Add After:=ActiveSheet
    Range("B3").Select
    ActiveCell.Formula2R1C1 = "=Inc_2021!R[-1]C:RC"
    Range("B3:F4").Select
    Selection.FillRight
    Selection.Columns.AutoFit
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Gross Profit"
    Range("C5").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-2]C-R[-1]C"
    Range("C5:F5").Select
    Selection.FillRight
    Selection.Font.Bold = True
    Range("B5:F5").Select
    Selection.Font.Bold = True
    Range("B6").Select
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "EBIT_CommonSize"
    Range("B6").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[-1]C"
    Range("B6:F6").Select
    Selection.FillRight
    Range("C6:F6").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B5").Select
    ActiveCell.FormulaR1C1 = "Net sales"
    Range("B7").Select
    ActiveCell.FormulaR1C1 = "Gross profit"
    Range("B7").Select
    Selection.Font.Bold = True
    Range("C7").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-2]C-R[-1]C"
    Range("C7:F7").Select
    Selection.FillRight
    Selection.Font.Bold = True
    Range("B8").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[3]C"
    Range("B9").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[3]C"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!RC"
    Range("B8:B10").Select
    Selection.AutoFill Destination:=Range("B8:F10"), Type:=xlFillDefault
    Range("B8:F10").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B11").Select
    ActiveCell.FormulaR1C1 = "EBIT"
    Range("B11").Select
    Selection.Font.Bold = True
    Range("C11").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C-R[-3]C-R[-2]C+R[-1]C"
    Range("C11:F11").Select
    Selection.FillRight
    Selection.Font.Bold = True
    Range("B3:F4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Sheets("Inc_2021").Select
    Range("B2:F3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -4.99893185216834E-02
        .PatternTintAndShade = 0
    End With
    Sheets("EBIT_CommonSize").Select
    Range("B6:F6").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    Range("B13").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("B10").Select
    Sheets("Inc_2021").Select
    Range("B5:F5").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("B7").Select
    Sheets("EBIT_CommonSize").Select
    Range("B8:F10").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("B12").Select
    Sheets("EBIT_CommonSize").Select
    Range("B10:F10").Select
    Selection.Cut
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Cut
    Range("B12").Select
    ActiveSheet.Paste
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[6]C"
    Range("B10:F10").Select
    Selection.FillRight
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Sheets("Inc_2021").Select
    Range("B14:F16").Select
    Range("B16").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("B11:F12").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Sheets("EBIT_CommonSize").Select
    Range("C11").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-4]C-R[-3]C-R[-2]C+R[-1]C"
    Range("C11:F11").Select
    Selection.FillRight
    Range("B12").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[-5]C"
    Range("B13").Select
    ActiveCell.Formula2R1C1 = "=Inc_2021!R[-5]C:R[-4]C"
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[-5]C"
    Range("B12:B15").Select
    Selection.AutoFill Destination:=Range("B12:F15"), Type:=xlFillDefault
    Range("B12:F15").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B16").Select
    ActiveCell.FormulaR1C1 = "EBT"
    Range("B16").Select
    Selection.Font.Bold = True
    Range("C16").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-5]C+R[-4]C-R[-3]C+R[-1]C"
    Range("C16:F16").Select
    Selection.FillRight
    Selection.Font.Bold = True
    Range("B17").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[1]C"
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!R[1]C"
    Range("B17:B18").Select
    Selection.AutoFill Destination:=Range("B17:F18"), Type:=xlFillDefault
    Range("B17:F18").Select
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "EAT"
    Range("B19").Select
    Selection.Font.Bold = True
    Range("C19").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-3]C-R[-2]C-R[-1]C"
    Range("C19:F19").Select
    Selection.FillRight
    Range("B20").Select
    Sheets("EBIT_CommonSize").Select
    ActiveCell.FormulaR1C1 = "check"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C-Inc_2021!RC"
    Range("C20:F20").Select
    Selection.FillRight
    Range("B21").Select
    ActiveCell.FormulaR1C1 = "=Inc_2021!RC"
    Range("B21:F21").Select
    Selection.FillRight
    Selection.Style = "Comma"
    Selection.NumberFormat = "_-* #,##0.0_-;-* #,##0.0_-;_-* ""-""??_-;_-@_-"
    Selection.NumberFormat = "_-* #,##0_-;-* #,##0_-;_-* ""-""??_-;_-@_-"
    Range("B12:F15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("C19:F19").Select
    Selection.Font.Bold = True
    Range("B17:F18").Select
    Range("F17").Activate
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -9.99786370433668E-02
        .PatternTintAndShade = 0
    End With
    Range("B21:F21").Select
    Selection.Font.Bold = True
    Range("B3:F5").Select
    Range("B3:F7").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B3:F11").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Range("B12:F16").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B17:F21").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B23").Select
    ActiveCell.FormulaR1C1 = "Common Size"
    Range("B24").Select
    Sheets("EBIT_CommonSize").Select
    Range("B5:F21").Select
    Selection.Copy
    Range("B25").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=13
    Range("B25:F27").Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("B25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-20]C"
    Range("B25:B41").Select
    ActiveWindow.SmallScroll Down:=-12
    Range("B5:B21").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=13
    Range("B25").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("C25").Select
    Application.CutCopyMode = False
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=R[-20]C/R5C"
    Range("C25").Select
    Selection.Style = "Percent"
    Range("C25:C41").Select
    Range("C25").Select
    Selection.Copy
    Range("C26:C41").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("C25:C41").Select
    Selection.Copy
    Range("D25:F41").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.0%"
    Selection.NumberFormat = "0.00%"
    Range("G27").Select
    ActiveWindow.SmallScroll Down:=6
    Range("D39").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("H27").Select
    ActiveWorkbook.Save
End Sub
