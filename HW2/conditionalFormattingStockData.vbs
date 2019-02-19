Sub setConditionalFormatting()

    Dim rg As Range
    Dim cond1 As FormatCondition, cond2 As FormatCondition

    'define the range that requires formatting
    Set rg = Range("J2", Range("J2").End(xlDown))
 
    'clear any existing conditional formatting
    rg.FormatConditions.Delete
 
    'define the rule for each conditional format
    Set cond1 = rg.FormatConditions.Add(xlCellValue, xlGreaterEqual, "0")
    Set cond2 = rg.FormatConditions.Add(xlCellValue, xlLess, "0")
 
    'define the format applied for each conditional format
    With cond1
       .Interior.Color = vbGreen
    End With
 
    With cond2
        .Interior.Color = vbRed
    End With
  
End Sub
