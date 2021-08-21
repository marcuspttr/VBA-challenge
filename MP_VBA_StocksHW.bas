Attribute VB_Name = "Module11"
Sub stockcleanup():

'It will run these commmands for each worksheet
For Each ws In Worksheets

    'Declaring variables that will be needed
    Dim stockname As String
    Dim stockopen As Double
    Dim stockclose As Double
    Dim stockvolume As Double

    'These are separate as one reads through the stocks and the other is for placing each summary
    Dim rowtracker As Long
    rowtracker = 2
    Dim summarytracker As Long
    summarytracker = 2


    
    'Setting headings for the summary data
    ws.Range("I1") = "<ticker>"
    ws.Range("J1") = "<yearly change>"
    ws.Range("K1") = "<percent change>"
    ws.Range("L1") = "<total stock volume>"

    ws.Range("O2") = "Greatest % Increase"
    ws.Range("O3") = "Greatest % Decrease"
    ws.Range("O4") = "Greatest Total Volume"
    ws.Range("P1") = "<ticker>"
    ws.Range("Q1") = "<value>"
    
    'Setting initial values for the initial top of the list stock.
    stockname = Range("A2").Value
    stockopen = Range("C2").Value
    stockvolume = Range("G2").Value

   

    'As long as there is data it will keep reading down the page
    While Not IsEmpty(Cells(rowtracker, 1))
        
        'It reads to see if it has come to the end of a particular stock.
        'If not then add to stock volume
        If ws.Cells(rowtracker + 1, 1) = stockname Then
            stockvolume = stockvolume + Cells(rowtracker + 1, 7).Value

        'If it detects a change to a new stock it will note the final closing value, calculate the change, and show the final volume.
        Else:
            stockclose = ws.Cells(rowtracker, 6).Value
            
            ws.Cells(summarytracker, 9) = stockname
            ws.Cells(summarytracker, 10) = stockclose - stockopen
            ws.Cells(summarytracker, 12) = stockvolume
            
            'If the change in the stock is negative it formats the cell red, positive is green
            If ws.Cells(summarytracker, 10) < 0 Then
                ws.Cells(summarytracker, 10).Interior.ColorIndex = 3
            
            Else
                ws.Cells(summarytracker, 10).Interior.ColorIndex = 4
            
            End If
            
            'If a stock opens with 0 value, then you cannot divide by that.
            'This detects that a as a potential issue and just reports out 0% change.
            If stockopen = 0 Then
                ws.Cells(summarytracker, 11) = 0
            
            ElseIf stockopen <> 0 Then
                ws.Cells(summarytracker, 11) = Format((stockclose - stockopen) / stockopen, "percent")
            
            End If
            
            'Captures the name, opening value, and initial stock volume of the new, different stock.
            stockname = ws.Cells(rowtracker + 1, 1).Value
            stockopen = ws.Cells(rowtracker + 1, 3).Value
            stockvolume = ws.Cells(rowtracker + 1, 7).Value
        
            'Used so the next summary is not placed on top of the data.
            summarytracker = summarytracker + 1
        
        End If
        
        'Moves down the data list
        rowtracker = rowtracker + 1
    
    Wend
    
    'Now that each stock's summary is available, this is setting up the bonus analysis
    'Rowtracker resets back to the top of the page
    rowtracker = 2
    Dim greatestname(2) As String
    Dim greatestpercent(1) As Double
    Dim greatestvolume As Double
    
    'Takes initial values from first stock summary.
    greatestname(0) = greatestname(1) = greatestname(2) = ws.Cells(rowtracker, 9)
    greatestpercent(0) = greatestpercent(1) = ws.Cells(rowtracker, 11)
    greatestvolume = ws.Cells(rowtracker, 12)
    
    'Starts back at the top now reading through the summary info.
    'Constantly comparing values to find greatest % growth and volume. Lowest growth as well.
    While Not IsEmpty(ws.Cells(rowtracker, 9))
        
        If ws.Cells(rowtracker + 1, 11) > greatestpercent(0) Then
            greatestname(0) = ws.Cells(rowtracker + 1, 9)
            greatestpercent(0) = ws.Cells(rowtracker + 1, 11)
        
        ElseIf ws.Cells(rowtracker + 1, 11) < greatestpercent(1) Then
            greatestname(1) = ws.Cells(rowtracker + 1, 9)
            greatestpercent(1) = ws.Cells(rowtracker + 1, 11)
        
        ElseIf ws.Cells(rowtracker + 1, 12) > greatestvolume Then
            greatestname(2) = ws.Cells(rowtracker + 1, 9)
            greatestvolume = ws.Cells(rowtracker + 1, 12)
    
    End If
    
    rowtracker = rowtracker + 1
    
    Wend
    
    'Spits out it's findings
    ws.Range("P2") = greatestname(0)
    ws.Range("Q2") = Format(greatestpercent(0), "percent")
    ws.Range("P3") = greatestname(1)
    ws.Range("Q3") = Format(greatestpercent(1), "percent")
    ws.Range("P4") = greatestname(2)
    ws.Range("Q4") = greatestvolume

'Once it has finished reading all of the data and summarizing the worksheet, it goes to the next.
Next ws

End Sub
