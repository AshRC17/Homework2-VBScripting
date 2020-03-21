Sub WorksheetCalculations():

    'Count the amount of active/open worksheets open in the workbook
    wsCount = ActiveWorkbook.Worksheets.Count
    
    'Quick check to output the number of workshsheets open
    MsgBox (ActiveWorkbook.Worksheets.Count)
    
    'Now to go through each worksheet and call functions to start populating data
    For Each ws In Worksheets
        'if you don't do this step excel will just keep adding columns to the first worksheet you had selected
        ws.Activate
        Call AddLabels
        Call Symbols
        Call CalculateHardShit
        
    Next ws

End Sub

'This will generate all headers
'for the data accordingly
Function AddLabels():

    'Labels for the first set of data
    Range("J1").Value = "Ticker Symbol"
    Range("K1").Value = "Yearly change"
    Range("L1").Value = "Yearly Change %"
    Range("M1").Value = "Total Volume"
    
    'Labels for the harder set of data
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
 
    
End Function


'Function to calculate Ticker symbol,
'Yearly Change, Yearly Change % and Total Volume
'and change the color of the cell for Yearly
'change and Change percentage accordingly
Function Symbols():
    'Length of Column variable
    Dim loc As Long
    
    'Summary table location variable
    Dim STableR As Integer
    
    'Ticker symbol variable
    Dim tname As String
    
    'Volume total variable
    Dim VTotal As Double
    
    'Opening value variable
    Dim ovalue As Double
    
    'Closing value variable
    Dim cvalue As Double
    
    'Yearly change variable
    Dim ychange As Double
     
    'Initialize the counter for the row where the results table will start
    STableR = 2
    
    'Calculate the length of the Symbol column
    loc = nor("A2")
    
    For i = 2 To loc
        'Check to see if we not at same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'If so then set the ticker name
            tname = Cells(i, 1).Value
            
            'Add up the total Volume
            VTotal = VTotal + Cells(i, 7)
            
            'Store the closing value to variable
            cvalue = Cells(i, 6).Value
            
            'Print ticker name in the summary table
            Range("J" & STableR).Value = tname
            
            'Print yearly change in the summary table
            ychange = cvalue - ovalue
            Range("K" & STableR).Value = ychange
            If ychange >= 0 Then
                Range("K" & STableR).Interior.ColorIndex = 4
            Else
                Range("K" & STableR).Interior.ColorIndex = 3
            End If
            
            'Format cell as percentage and then print yearly change percentage in the summary table
            Range("L" & STableR).NumberFormat = "0.00%"
            
            'If statement to catch 0/0 overflow error message and Div/0 error message
            If ychange = 0 And ovalue = 0 Then
                Range("L" & STableR).Value = 0
            ElseIf ychange > 0 And ovalue = 0 Then
                Range("L" & STableR).Value = ychange
            Else
                Range("L" & STableR).Value = ychange / ovalue
            End If
            
            'Change the color of the cells according to the value in the cell
            If ychange >= 0 Then
                Range("L" & STableR).Interior.ColorIndex = 4
            Else
                Range("L" & STableR).Interior.ColorIndex = 3
            End If
            
            'Print total volume in the summary table
            Range("M" & STableR).Value = VTotal
            
            'Increase summary table row by 1
            STableR = STableR + 1
            
            'Reset the volume total
            VTotal = 0
            
        
        'See if the ticker we are at is not the same as the cell before.
        ElseIf Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
            'If not store the opening value
            ovalue = Cells(i, 3)
            
            'Keep adding the volume total together
            VTotal = VTotal + Cells(i, 7).Value
        
        'Else
        Else
        
            'Keep adding the volume total together
            VTotal = VTotal + Cells(i, 7).Value
            
        End If

    Next i
    
End Function

'Function to calculate the number of rows
Function nor(stCell):
    'Variable to hold the nunber of rows
    Dim nOfRows As Long
    
    'Variable to hold the starting cell
    Dim setR As String
    setR = stCell
    'MsgBox (setR)
    
    'If loop to count the number of rows in a column
    If IsEmpty(Range(setR)) Then
        nOfRows = 0
        nor = 0
        'MsgBox (Str(nOfRows) + " rows available")
    ElseIf IsEmpty(Range(setR).Offset(1, 0)) Then
        nOfRows = 1
        nor = 1
        'MsgBox ("Only " + Str(numOfRows) + " row")
    Else
        nOfRows = Range(setR, Range(setR).End(xlDown)).Rows.Count
        nor = Range(setR, Range(setR).End(xlDown)).Rows.Count
        'MsgBox (nOfRows)
    End If
End Function

'Function to print out list of tickers symbols in that sheet
Function TickerList():
    'VAriable for length of column
    Dim loc As Long
    
    'Variable for summary table row start point
    Dim STableR As Integer
    
    'Variable for ticker symbol/name
    Dim tname As String
    STableR = 2
    loc = nor("A2")
    
    For i = 2 To loc
        'Check to see if we are will at same ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'If so then set the ticker name
            tname = Cells(i, 1).Value
            
            'Print ticker name in the summary table
            Range("J" & STableR).Value = tname
            
            'Increase summary table row by 1
            STableR = STableR + 1
            
        End If
        
    Next i
           
End Function

Function CalculateHardShit():
    
    'Calculate HardShit number of rows
    Dim HSNoR As Integer
    HSNoR = nor("J2")
    
    'Create the variables to store the value and symbol of the greatest % increase
    Dim gpiv As Double
    Dim gpis As String
        
    'Create the variables to store the value and symbol of the greatest % decrease
    Dim gpdv As Double
    Dim gpds As String
    
    'Create the variables to store the value and symbol of the greatest total volume
    Dim gtvv As Double
    Dim gtvs As String
    
    'MsgBox (HSNoR)
    
    'Loop through all individual ticker symbols
    For tic = 2 To HSNoR + 1
        If Range("L" & tic).Value > gpiv Then
            gpiv = Range("L" & tic).Value
            gpis = Range("J" & tic).Value
        ElseIf Range("L" & tic).Value < gpdv Then
            gpdv = Range("L" & tic).Value
            gpds = Range("J" & tic).Value
        End If
        If Range("M" & tic).Value > gtvv Then
            gtvv = Range("M" & tic).Value
            gtvs = Range("J" & tic).Value
        End If
    Next tic
        
    'Assign the symbol and value for the greatest increase
    Range("Q2").NumberFormat = "0.00%"
    Range("P2").Value = gpis
    Range("Q2").Value = gpiv
    
    'Assign the symbol and value for the greatest decrease
    Range("Q3").NumberFormat = "0.00%"
    Range("P3").Value = gpds
    Range("Q3").Value = gpdv
    
    'Assign the symbol and value for the greatest total volume
    Range("P4").Value = gtvs
    Range("Q4").Value = gtvv
        
End Function

