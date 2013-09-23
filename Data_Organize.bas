Attribute VB_Name = "Data_Organize"
Sub BKorganize()
Attribute BKorganize.VB_ProcData.VB_Invoke_Func = " \n14"
'
' BKorganize Macro
'

'
    Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = "Compiled Data"
    Sheets("Compiled Data").Select
    
    Range("B5").Value = "LAeq"
    Range("B7").Value = "12.5"
    Range("B8").Value = "16"
    Range("B9").Value = "20"
    Range("B10").Value = "25"
    Range("B11").Value = "31.5"
    Range("B12").Value = "40"
    Range("B13").Value = "50"
    Range("B14").Value = "63"
    Range("B15").Value = "80"
    Range("B16").Value = "100"
    Range("B17").Value = "125"
    Range("B18").Value = "160"
    Range("B19").Value = "200"
    Range("B20").Value = "250"
    Range("B21").Value = "315"
    Range("B22").Value = "400"
    Range("B23").Value = "500"
    Range("B24").Value = "630"
    Range("B25").Value = "800"
    Range("B26").Value = "1000"
    Range("B27").Value = "1250"
    Range("B28").Value = "1600"
    Range("B29").Value = "2000"
    Range("B30").Value = "2500"
    Range("B31").Value = "3150"
    Range("B32").Value = "4000"
    Range("B33").Value = "5000"
    Range("B34").Value = "6300"
    Range("B35").Value = "8000"
    Range("B36").Value = "10000"
    Range("B37").Value = "12500"
    Range("B38").Value = "16000"
    Range("B39").Value = "2000"
    Range("B41").Value = "LAeq"
    Range("B43").Value = "16"
    Range("B44").Value = "31.5"
    Range("B45").Value = "63"
    Range("B46").Value = "125"
    Range("B47").Value = "250"
    Range("B48").Value = "500"
    Range("B49").Value = "1000"
    Range("B50").Value = "2000"
    Range("B51").Value = "4000"
    Range("B52").Value = "8000"
    Range("B53").Value = "16000"
    Range("C41").Value = C5
    
    
    Range("C3").Value = "1"
    Range("D3").Value = "2"
    Range("C3:D3").Select
    Selection.AutoFill Destination:=Range("C3:CX3"), Type:=xlFillDefault

    Columns("B:B").Select
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Rows("3:3").Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Range("A1", "CX60").Select
    Selection.HorizontalAlignment = xlCenter
    Range("C5", "CX60").Select
    Selection.NumberFormat = "0.0"
    
    Sheets("TotalBB").Select
    Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Select
    Selection.Find(What:="LAeq").Select
    Range(Selection.Offset(1, 0), Selection.End(xlDown)).Copy
    Sheets("Compiled Data").Select
    Range("C5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
      
    Sheets("TotalSpectra").Select
    Range(Cells(1, 1), Cells(1, 1).End(xlToRight)).Select
    Selection.Find(What:="LZeq 12.5Hz").Select
    Range(Selection.Offset(1, 0), Selection.Offset(1, 0).End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Copy
    Sheets("Compiled Data").Select
    Range("C7").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=True
    
    Range("C5", "CX5").Copy Destination:=Range("C41", "CX41")
    
    Range("C43").Formula = "=10*LOG(10^(C7/10)+10^(C8/10)+10^(C9/10))"
    Range("C44").Formula = "=10*LOG(10^(C10/10)+10^(C11/10)+10^(C12/10))"
    Range("C45").Formula = "=10*LOG(10^(C13/10)+10^(C14/10)+10^(C15/10))"
    Range("C46").Formula = "=10*LOG(10^(C16/10)+10^(C17/10)+10^(C18/10))"
    Range("C47").Formula = "=10*LOG(10^(C19/10)+10^(C20/10)+10^(C21/10))"
    Range("C48").Formula = "=10*LOG(10^(C22/10)+10^(C23/10)+10^(C24/10))"
    Range("C49").Formula = "=10*LOG(10^(C25/10)+10^(C26/10)+10^(C27/10))"
    Range("C50").Formula = "=10*LOG(10^(C28/10)+10^(C29/10)+10^(C30/10))"
    Range("C51").Formula = "=10*LOG(10^(C31/10)+10^(C32/10)+10^(C33/10))"
    Range("C52").Formula = "=10*LOG(10^(C34/10)+10^(C35/10)+10^(C36/10))"
    Range("C53").Formula = "=10*LOG(10^(C37/10)+10^(C38/10)+10^(C39/10))"

    
    Range("C43:C53").Select
    Selection.AutoFill Destination:=Range("C43:CX53"), Type:=xlFillDefault
    
    Range("C5", Range("C5").End(xlToRight)).Select
    numSelected = WorksheetFunction.CountA(Selection)
    Range(Range("C1").Offset(0, numSelected), "CX53").Delete

    
    
End Sub


Sub KingstonCompile()
'
' KingstonCompile Macro
'
' Keyboard Shortcut: Ctrl+Shift+K
'
Application.ScreenUpdating = False

    Sheets.Add After:=Sheets(Sheets.Count)
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Compiled Data"
    Sheets("OBA").Select
    Range("B2:M2").Select
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("OBA").Select
    Range("B3:M3").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("C3").Select
    ActiveSheet.Paste
    Sheets("OBA").Select
    Range("B8:AK8").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("C18").Select
    ActiveSheet.Paste
    Range("B3").Select
    Sheets("OBA").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Sheets("Compiled Data").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Leq"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "Leq"
    Range("B20").Select
    Sheets("OBA").Select
    Range("B9:AK9").Select
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("C19").Select
    ActiveSheet.Paste
    Sheets("OBA").Select
    Range("F27").Select
    ActiveWindow.SmallScroll Down:=15
    Sheets("Time History").Select
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 72
    Sheets("Compiled Data").Select
    Range("B18:AL19").Select
    Application.CutCopyMode = False
    Selection.Cut
    Range("R2").Select
    ActiveSheet.Paste
    Range("H41").Select
    Sheets("Time History").Select
    Range("CB3:DK2550").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=15
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("S5").Select
    ActiveSheet.Paste
    Sheets("Measurement History").Select
    Range("I41").Select
    Sheets("Time History").Select
    ActiveWindow.ScrollColumn = 76
    ActiveWindow.ScrollColumn = 75
    ActiveWindow.ScrollColumn = 74
    ActiveWindow.ScrollColumn = 73
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.ScrollColumn = 71
    ActiveWindow.ScrollColumn = 70
    ActiveWindow.ScrollColumn = 69
    ActiveWindow.ScrollColumn = 68
    ActiveWindow.ScrollColumn = 67
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 65
    ActiveWindow.ScrollColumn = 64
    ActiveWindow.ScrollColumn = 63
    ActiveWindow.ScrollColumn = 62
    ActiveWindow.ScrollColumn = 61
    ActiveWindow.ScrollColumn = 60
    ActiveWindow.ScrollColumn = 59
    ActiveWindow.ScrollColumn = 58
    ActiveWindow.ScrollColumn = 57
    ActiveWindow.ScrollColumn = 56
    ActiveWindow.ScrollColumn = 55
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 53
    ActiveWindow.ScrollColumn = 52
    ActiveWindow.ScrollColumn = 51
    ActiveWindow.ScrollColumn = 50
    ActiveWindow.ScrollColumn = 49
    ActiveWindow.ScrollColumn = 48
    ActiveWindow.ScrollColumn = 47
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 45
    ActiveWindow.ScrollColumn = 44
    ActiveWindow.ScrollColumn = 43
    ActiveWindow.ScrollColumn = 42
    ActiveWindow.ScrollColumn = 41
    ActiveWindow.ScrollColumn = 40
    ActiveWindow.ScrollColumn = 39
    ActiveWindow.ScrollColumn = 38
    ActiveWindow.ScrollColumn = 37
    ActiveWindow.ScrollColumn = 36
    ActiveWindow.ScrollColumn = 35
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 33
    ActiveWindow.ScrollColumn = 32
    ActiveWindow.ScrollColumn = 31
    ActiveWindow.ScrollColumn = 30
    ActiveWindow.ScrollColumn = 29
    ActiveWindow.ScrollColumn = 28
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 26
    ActiveWindow.ScrollColumn = 25
    ActiveWindow.ScrollColumn = 24
    ActiveWindow.ScrollColumn = 23
    ActiveWindow.ScrollColumn = 22
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 20
    ActiveWindow.ScrollColumn = 19
    ActiveWindow.ScrollColumn = 18
    ActiveWindow.ScrollColumn = 17
    ActiveWindow.ScrollColumn = 16
    ActiveWindow.ScrollColumn = 15
    ActiveWindow.ScrollColumn = 14
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 12
    ActiveWindow.ScrollColumn = 11
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 9
    ActiveWindow.ScrollColumn = 8
    ActiveWindow.ScrollColumn = 7
    ActiveWindow.ScrollColumn = 6
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 3
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.SmallScroll Down:=-216
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2718").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2677").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2636").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2595").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2554").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2513").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2472").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2431").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2390").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2349").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2308").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2267").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2226").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2185").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2144").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2103").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2062").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB2021").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1980").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1939").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1898").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1857").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1816").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1775").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1734").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1693").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1652").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1611").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1570").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1529").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1488").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1447").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1406").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1365").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1324").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1283").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1242").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1201").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1160").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1119").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1078").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB1037").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB996").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB955").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB914").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB873").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB832").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB791").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB750").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB709").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB668").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB627").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB586").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB545").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB504").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB463").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB422").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB381").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB340").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB299").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB258").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB217").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB176").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB135").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB94").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB53").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("CB12").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("AF3:AQ3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("C5").Select
    ActiveSheet.Paste
    Sheets("Time History").Select
    Range("Q2981").Select
    ActiveWindow.SmallScroll Down:=-15
    ActiveSheet.Previous.Select
    Selection.End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Sheets("Time History").Select
    Range("D2977").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    Range("C3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compiled Data").Select
    Columns("A:A").Select
    Columns("A:C").Select
    Application.CutCopyMode = False
    Columns("A:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("C10").Select
    Sheets("Time History").Select
    Selection.Copy
    Range("C3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.End(xlToLeft).Select
    Range("C3").Select
    Selection.End(xlToRight).Select
    Selection.End(xlToLeft).Select
    Range("C3:E3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("B5").Select
    ActiveSheet.Paste
    Columns("B:B").EntireColumn.AutoFit
    Range("D4").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Laeq"
    Range("D3").Select
    ActiveCell.FormulaR1C1 = "Overall"
    Range("D4").Select
    ActiveCell.FormulaR1C1 = "Laeq"
    Rows("4:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("H4").Select
    Sheets("Time History").Select
    ActiveWindow.SmallScroll Down:=12
    ActiveWindow.ScrollColumn = 4
    ActiveWindow.ScrollColumn = 5
    ActiveWindow.ScrollColumn = 10
    ActiveWindow.ScrollColumn = 13
    ActiveWindow.ScrollColumn = 21
    ActiveWindow.ScrollColumn = 27
    ActiveWindow.ScrollColumn = 34
    ActiveWindow.ScrollColumn = 46
    ActiveWindow.ScrollColumn = 54
    ActiveWindow.ScrollColumn = 66
    ActiveWindow.ScrollColumn = 72
    ActiveWindow.SmallScroll Down:=-24
    ActiveWindow.LargeScroll Down:=-1
    Range("C2907").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2866").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2825").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2784").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2743").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2702").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2661").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2620").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2579").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2538").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2497").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2456").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2415").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2374").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2333").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2292").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2251").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2210").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2169").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2128").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2087").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2046").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C2005").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1964").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1923").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1882").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1841").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1800").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1759").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1718").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1677").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1636").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1595").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1554").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1513").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1472").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1431").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1390").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1349").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1308").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1267").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1226").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1185").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1144").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1103").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1062").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C1021").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C980").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C939").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C898").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C857").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C816").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C775").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C734").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C693").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C652").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C611").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C570").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C529").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C488").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C447").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C406").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C365").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C324").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C283").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C242").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C201").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C160").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C119").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C78").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C37").Select
    ActiveWindow.LargeScroll Down:=-1
    Range("C37").Select
    Sheets("OBA").Select
    ActiveWindow.SmallScroll Down:=-24
    Sheets("Summary").Select
    ActiveWindow.SmallScroll Down:=21
    Range("B39").Select
    Selection.Copy
    Sheets("Compiled Data").Select
    Range("D3").Select
    ActiveSheet.Paste
    Range("D2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "Laeq"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "Overall"
    Range("B4").Select
    Sheets("Summary").Select
    ActiveWindow.SmallScroll Down:=15
    Range("A55:B57").Select
    Selection.Copy
    Sheets("Time History").Select
    ActiveWindow.LargeScroll ToRight:=-2
    Sheets("Compiled Data").Select
    ActiveWindow.SmallScroll Down:=-9
    Range("S6").Select
    ActiveSheet.Paste
    Range("S14").Select
End Sub
Sub CopyOver()
'
' CopyOver Macro
'
' Keyboard Shortcut: Ctrl+Shift+C
'
    Sheets("Compiled Data").Select
    Sheets("Compiled Data").Copy After:=Workbooks("Book1").Sheets(1)
End Sub
Sub Macro3()
'
' Macro3 Macro
'

'
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
End Sub
Sub Macro9()
'
' Macro9 Macro
'
' Keyboard Shortcut: Ctrl+m
'
    
    'ActiveSheet.Move _
      ' After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)


    
    i = 6
    For i = 6 To 21
    ActiveWorkbook.Sheets(i).Select
    'ActiveSheet.Name = "dlf" & i
    ActiveSheet.Name = i + 4
    Next
    
    
End Sub

