Sub runHWForAllWorksheets()

        'Declare Current as a worksheet object variable.
        Dim Current As Worksheet

        Application.ScreenUpdating = False
        
         'Loop through all of the worksheets in the active workbook.
         For Each Current In Worksheets
            'select the current worksheet
            Current.Select
            
            'run the code to summarize stock data by stock
            Call testHardHW(Current.Name)
            
            'format worksheet
            Current.Columns("A:P").AutoFit
            Call setConditionalFormatting
            
         Next
         
         Application.ScreenUpdating = True


End Sub
