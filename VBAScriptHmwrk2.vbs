Attribute VB_Name = "Module1"
Sub BP_Macro()
    For Each W In Worksheets
    
    'Declare variable to hold Ticket and Total Volumne for each ticker
    Dim Ticker As String
    Dim VolumeTotal As Double
    VolumeTotal = 0
    
    'Set location for each Ticker in the summary table
    Range("J1").Value = "Ticker"
    Range("K1").Value = "TotalVolume"
    Dim SummaryRow As Integer
    SummaryRow = 2 'account for header row
    
    'Count the number of rows in data set
    iSummaryRow = Cells(Rows.Count, 1).End(xlUp).Row 'Count the number of records

    'Loop through all Ticker rows
    For r = 2 To iSummaryRow
    
        'Check if still within the same Ticker
        If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then 'Next cell <> current cell
           
           'Set Ticker Name
           TickerName = Cells(r, 1).Value
           
           'Add to the Volume total
           VTotal = VTotal + Cells(r, 7).Value
           
           'Add Ticker Name and Total Volume to Summary Table
           Range("J" & SummaryRow).Value = TickerName
           Range("K" & SummaryRow).Value = Format(VTotal, "Currency") 'Format Currency
           
           'Add one to Summary Row counter
           SummaryRow = SummaryRow + 1
           
           'Reset Card Total
           VTotal = 0
           
        Else 'If following cell is the same card
           VTotal = VTotal + Cells(r, 7).Value
           
        End If
    Next r
    
    Next W
    
End Sub
