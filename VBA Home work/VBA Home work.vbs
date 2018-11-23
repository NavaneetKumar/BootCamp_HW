Sub CleanUp()
    For Each sh In Worksheets
        sh.Activate
        
        With sh
            .Range("I1").Select
            If .Range("I1") = "" Then
                GoTo DoNothing
            End If
            
            .Range(Selection, Selection.End(xlToRight)).Select
            .Range(Selection, Selection.End(xlDown)).Select
            Selection.Delete
            
            
        
        End With
DoNothing:
    Next
End Sub

Sub BuildReport()
    
    Dim sh As Worksheet
    Dim rn As Range
    Dim rowmax As Long
    Dim rowcnt As Long
    Dim rptcnt As Long
    Dim vol As Double
    Dim openpt, closept
    
    'Loop through each worksheet
    For Each sh In Worksheets
        'Set the report header
        sh.Range("I1") = "Ticker"
        sh.Range("J1") = "Yearly Change"
        sh.Range("K1") = "Percentage Change"
        sh.Range("L1") = "Total Stock Volume"
            
        'Initialize the volume total
        vol = 0
        
        'Get the end of row
        Set rn = sh.UsedRange
        rowmax = rn.Rows.Count + rn.Row - 1
        
        'intitialize the Report counter
        rptcnt = 2
        
        'Loop through Ticker
        With sh
        
        'set the open value for the first ticker
        openpt = .Range("C2")
        
        For rowcnt = 2 To rowmax
            'Check if Ticker is changing
            If .Range("A" & rowcnt) <> .Range("A" & rowcnt + 1) Then
             'if in then we are at the last row of ticker
             
             'Add the Ticker in Report
             .Range("I" & rptcnt) = .Range("A" & rowcnt)
             
             'Add the last volume
             vol = vol + .Range("G" & rowcnt)
             
             'Add the final volume to the report
             .Range("L" & rptcnt) = vol
             
             'reset the volume total
             vol = 0
                     
             'set the close value from the last record of each ticker
             closept = .Range("F" & rowcnt)
             
             'set the difference in report
             .Range("J" & rptcnt) = closept - openpt
             
             'set the color formatting for difference
             If (closept >= openpt) Then
                .Range("J" & rptcnt).Interior.ColorIndex = 4
             Else
                .Range("J" & rptcnt).Interior.ColorIndex = 3
             End If
             
             'set the change percentage and it's formatting
             If openpt = 0 Then
                .Range("K" & rptcnt) = 1
             Else
                .Range("K" & rptcnt) = (closept - openpt) / openpt
             End If
             .Range("K" & rptcnt).NumberFormat = "0.00%"
             
             'set the open value from the first record of each ticker
             openpt = .Range("C" & rowcnt + 1)
             
             'running only for9 ticker
             'If rptcnt = 10 Then
             '   GoTo DoNothing
             'End If
             
             'increment the report counter
             rptcnt = rptcnt + 1
            End If
            
            'Keep on adding the volume for same ticker
            vol = vol + .Range("G" & rowcnt)
        Next rowcnt
        End With
'DoNothing:
        BuildSecondReport sh
    Next
End Sub

Sub BuildSecondReport(sh As Worksheet)
    GrtInc = 0
    GrtIncTik = ""
    GrtDec = 0
    GrtDecTik = ""
    GrtVol = 0
    GrtVolTik = ""
    With sh
        .Range("P1") = "Ticket"
        .Range("Q1") = "Volume"
        .Range("O2") = "Greatest % increase"
        .Range("O3") = "Greatest % Decrease"
        .Range("O4") = "Greatest total volume"
        
        LastRow = .Range("I1").CurrentRegion.Rows.Count
        For rowcnt = 2 To LastRow
            If GrtInc < .Range("K" & rowcnt) Then
                GrtInc = .Range("K" & rowcnt)
                GrtIncTik = .Range("I" & rowcnt)
            End If
            
            If GrtDec > .Range("K" & rowcnt) Then
                GrtDec = .Range("K" & rowcnt)
                GrtDecTik = .Range("I" & rowcnt)
            End If
            
            If GrtVol < .Range("L" & rowcnt) Then
                GrtVol = .Range("L" & rowcnt)
                GrtVolTik = .Range("I" & rowcnt)
            End If
            
        Next
        .Range("Q2") = GrtInc
        .Range("Q2").NumberFormat = "0.00%"
        .Range("P2") = GrtIncTik
        .Range("Q3") = GrtDec
        .Range("Q3").NumberFormat = "0.00%"
        .Range("P3") = GrtDecTik
        .Range("Q4") = GrtVol
        .Range("P4") = GrtVolTik
    End With
End Sub

