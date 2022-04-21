Attribute VB_Name = "LargeClaims"
Option Explicit

Sub Large_Claim_Auto()
'this code will take 3 columns of data (Month, ID, Claims) and produce an additive matrix of claims
'will work for any given period length (i.e. plan year is 12 months, but this will work for any positive number of months)

    Dim Month() As Date
    Dim ID() As String
    Dim Amount() As Variant
    Const StartRow As Long = 2
    Dim lastRow As Long
    Dim chDate As Date
    Dim j As Long
    Dim k As Long
    Dim l As Long
    Dim perStart As Date
    Dim perLength As Long
    Dim monthLabel() As Date
    Dim uniqueID() As String
    Dim claimants As Long
    Dim newStart As Long
    Dim rowSum As Variant
    Dim subValue As Variant
    Dim lastMonth As Date
    Dim lastClaimants() As String
    Dim threshDrop As VbMsgBoxResult
    Dim activeMonths As Long
    Dim indvCount As Long
    Dim sh As Worksheet
    Dim shCount As Long
    Dim rowTracker As Variant
    
    On Error GoTo ErrorHandle
    
    'error handling for checking if "Input" and "Output" tabs are available
    shCount = Sheets.Count
    l = 0
    For j = 1 To shCount
        If Sheets(j).Name = "Input" Or Sheets(j).Name = "Output" Then
            l = l + 1
        End If
    Next j
    
    If l <> 2 Then
        MsgBox "Input and/or Output tab(s) are missing, please use exact spelling." & vbNewLine & "Toggle Help for further information."
        Exit Sub
    End If
    
    Sheets("Input").Activate
    'gathering some basic information (period start & period length)
    perStart = Range("G1").Value
    perLength = Range("G2").Value
    
    'error handling for period start and length
    'most of the errors here will be handled by ErrorHandle, but this will serve as backup
    If IsNumeric(perLength) = False Or perLength < 1 Then
        MsgBox "Error: please fix the period length and try again."
        Exit Sub
    End If
    
    
    'this will redimension the arrays as the correct size
    'LastRow = Range("A2").End(xlDown).Row (don't use this method in case there is a month entered as a blank)
    lastRow = Range("A" & Rows.Count).End(xlUp).Row
    ReDim Month(StartRow To lastRow)
    ReDim ID(StartRow To lastRow)
    ReDim uniqueID(StartRow To lastRow)
    ReDim Amount(StartRow To lastRow)
    
    'this will fill the initial arrays with the data
    For j = 2 To lastRow
        chDate = Range("A" & j).Value
        'error handling if incorrectly entered data makes it past ErrorHandle
        If IsDate(chDate) = False Or IsEmpty(chDate) = True Or chDate < perStart Or chDate > WorksheetFunction.EDate(perStart, perLength) Then
            MsgBox "Error: please fix dates and try again."
            Exit Sub
        End If
        Month(j) = Range("A" & j).Value
        ID(j) = Range("B" & j).Value
        Amount(j) = CDec(Range("C" & j).Value)
    Next j
    
    'switch to the output sheet and start inputting information
    With Sheets("Output")
        .Activate
        .Cells.ClearContents
        .Cells.ClearFormats
    End With

    With Range("A1")
        .Value = "ID"
        .HorizontalAlignment = xlCenter
    End With
    
    'redimension monthLabel to length of period
    ReDim monthLabel(1 To perLength)
    
    'fill monthLabel with appropriate months
    For j = 1 To perLength
        monthLabel(j) = WorksheetFunction.EDate(perStart, j - 1)
    Next j
    
    'paste monthLabel to Output sheet
    With Range(Cells(1, 2), Cells(1, 1 + perLength))
        .Value = monthLabel
        .NumberFormat = "m/d/yyy"
        .HorizontalAlignment = xlCenter
    End With
    
    'paste IDs and then make unique
    With Range("A2:A" & lastRow)
        .Value = Excel.WorksheetFunction.Transpose(ID)
        .RemoveDuplicates Columns:=1, Header:=xlNo
        .HorizontalAlignment = xlCenter
    End With
    
    'calculate claimants to create and fill uniqueID array
    claimants = Range("A2").End(xlDown).Row
    
    'this will prevent error if only 1 claimant used
    If IsEmpty(Range("A3")) Then
        claimants = 2
    End If

    
    ReDim uniqueID(StartRow To claimants)
       
    'print unique ID values and fill matrix with total monthly values
    'rows = j, columns = k, data array scan = l
    'reminder: cells(rows, columns)
    For j = 2 To claimants
        uniqueID(j) = Range("A" & j).Value
        For k = 2 To perLength + 1
            For l = 2 To lastRow
                If ID(l) = Cells(j, 1).Value And Month(l) = Cells(1, k).Value Then
                    Cells(j, k).Value = CDec(Amount(l))
                End If
            Next l
        Next k
    Next j
    
    'this will check if a claimant drops off the list (i.e. drops below threshold)
    'logic: if a claimant shows up in a prior month, but not in a later month, they get flagged
    lastMonth = Month(UBound(Month))
    ReDim lastClaimants(2 To claimants)
    For j = 2 To claimants
        For k = 2 To UBound(Month)
            If Month(k) = lastMonth And uniqueID(j) = ID(k) Then
                lastClaimants(j) = "Yes"
            End If
        Next k
    Next j
    

    rowTracker = 1
    For j = 2 To claimants
        rowTracker = rowTracker + 1
        If lastClaimants(j) = "" Then
            threshDrop = MsgBox("Claimant " & uniqueID(j) & " has dropped below threshold. Would you like to remove them?" & _
            vbNewLine & vbNewLine & "Select No to return to Input Page.", vbYesNo + vbDefaultButton1, "Threshold Violation")
            If threshDrop = vbYes Then
                Range("A" & rowTracker).EntireRow.Delete
                rowTracker = rowTracker - 1
                claimants = claimants - 1
            Else
                Sheets("Output").Cells.ClearContents
                Sheets("Input").Activate
                Exit Sub
            End If
        End If
    Next j

    
    'this will check if a claimant is entered twice in one month
    activeMonths = VBA.DateDiff("m", perStart, lastMonth) + 1
    indvCount = 0
    For j = 2 To claimants
        For k = 2 To UBound(ID)
            If uniqueID(j) = ID(k) Then
                indvCount = indvCount + 1
            End If
        Next k
        If indvCount > activeMonths Then
            MsgBox "Claimant " & uniqueID(j) & "has multiple data entries during a single month." & vbNewLine & _
            vbNewLine & "Please fix and try again.", , "Data Error"
            Sheets("Input").Activate
            Exit Sub
        End If
        indvCount = 0
    Next j
    
    'now we will turn the total monthly values into their additive monthly values
    'copy the matrix and paste below
    'newStart will give the starting position of the new matrix
    newStart = claimants + 2
    Range("A1").CurrentRegion.Copy
    Range("A" & newStart).PasteSpecial
    
    'we'll have the additive amounts in the matrix below
    'this code will perform the subtractions
    'reminder: cells(r, c)
    For j = 2 To claimants
        For k = 3 To perLength + 1
            If Cells(j + newStart - 1, k - 1).Value <> "" Then
                rowSum = CDec(WorksheetFunction.Sum(Range(Cells(j + newStart - 1, 2), Cells(j + newStart - 1, k - 1))))
                subValue = CDec(Cells(j, k).Value - rowSum)
                If subValue <= 0 And Cells(j, k).Value = "" Then
                    Cells(j + newStart - 1, k).Value = 0
                Else
                    Cells(j + newStart - 1, k).Value = CDec(subValue)
                End If
            End If
        Next k
    Next j

    
    'now replace the monthly totals with the monthly additives
    Range("A1:A" & newStart - 1).EntireRow.Delete
    Range("A1").Select
    
    'add total column
    With Cells(1, perLength + 2)
        .Value = "Total"
        .HorizontalAlignment = xlCenter
    End With
    For j = 2 To claimants
        Cells(j, perLength + 2).Value = WorksheetFunction.Sum(Range(Cells(j, 2), Cells(j, perLength + 1)))
    Next j
    
    'add monthly / period totals and counts
    With Cells(claimants + 1, 1)
        .Value = "Total $"
        .HorizontalAlignment = xlCenter
    End With
    With Cells(claimants + 2, 1)
        .Value = "Total #"
        .HorizontalAlignment = xlCenter
    End With
    For j = 2 To perLength + 2
        Cells(claimants + 1, j).Value = WorksheetFunction.Sum(Range(Cells(2, j), Cells(claimants, j)))
        Cells(claimants + 2, j).Value = WorksheetFunction.CountA(Range(Cells(2, j), Cells(claimants, j)))
    Next j
    
    'final formatting
    With Range("B2", Cells(claimants + 1, perLength + 2))
        .HorizontalAlignment = xlCenter
        .NumberFormat = "#,##0_);[Red](#,##0)"
        
    End With
    With Range(Cells(claimants + 1, 1), Cells(claimants + 2, perLength + 2))
        .HorizontalAlignment = xlCenter
        .Interior.Color = 14277081
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Font.Bold = True
    End With
    With Range("A1", Cells(claimants + 2, 1))
        .Interior.Color = 14277081
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Font.Bold = True
    End With
    With Range("A1", Cells(1, perLength + 2))
        .Interior.Color = 14277081
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Font.Bold = True
    End With
        With Range(Cells(1, perLength + 2), Cells(claimants + 2, perLength + 2))
        .Interior.Color = 14277081
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Font.Bold = True
    End With
    
    Range("A2", Cells(claimants, perLength + 2)).Sort Key1:=Range(Cells(2, perLength + 2), Cells(claimants, perLength + 2)), Order1:=xlDescending
    
    'ideas for future versions:
    'option to donate your data
    'option to add in more information (how many columns of data would you like?; dx, plan, etc.)
    'would you like to name the headers?
    'adds title to page (client name + date through when)
    'add graph option
    

ErrorHandle:
    Select Case Err.Number
        Case 13 'data type mismatch
            MsgBox "There was an error with the type of data. Please check formatting", , "Data Error"
            Exit Sub
    End Select

End Sub
