Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Sub stockMarketLoop()

	' Worksheet loop
	For Each ws In Worksheets
		
		' Naming headers
		ws.Range("I1").Value = "Ticker"
		ws.Range("J1").Value = "Yearly Change"
		ws.Range("K1").Value = "Percent Change"
		ws.Range("L1").Value = "Total Stock Volume"
		ws.Range("O2").Value = "Greatest % Increase"
		ws.Range("O3").Value = "Greatest % Decrease"
		ws.Range("O4").Value = "Greatest Total Volume"
		ws.Range("P1").Value = "Ticker"
		ws.Range("Q1").Value = "Value"
		
		' Variable Declaration
		Dim lastRow As Long
		Dim tickerName As String
		Dim yearlyChange As Double
		Dim percentChange As Double
		Dim totalStockVolume As Double
		Dim summaryTableRow As Long
		Dim yearlyOpen As Double
		Dim yearlyClose As Double
		Dim tickerOpeningRow As Long
		Dim greatestIncrease As Double
		Dim greatestDecrease As Double
		Dim lastRowValue As Long
		
		' Default Values
		totalStockVolume = 0
		summaryTableRow = 2
		tickerOpeningRow = 2
		greatestIncrease = 0		
		greatestTotalVolume = 0
						
		' Determine Last Row
		lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
		
		' Begin Loop
		For i = 2 to lastRow
		
			' Add Total Stock Volume per Ticker
			totalStockVolume = totalStockVolume + ws.Cells(i, 7).Value
			If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
			
				' Print Stock Results to Summary Table and reset Stock Counter
				tickerName = ws.Cells(i, 1).Value
				ws.Range("I" & summaryTableRow).Value = tickerName
				ws.Range("L" & summaryTableRow).Value = totalStockVolume
				totalStockVolume = 0
				
				' Calculate Yearly Change
				yearlyOpen = ws.Range("C" & tickerOpeningRow)
				yearlyClose = ws.Range("F" & i)
				yearlyChange = yearlyClose - yearlyOpen
				ws.Range("J" & summaryTableRow).Value = yearlyChange
				
				' Calculate Percent Change & Formatting, Prevent Div by Zero Err
				If yearlyOpen = 0 Then
					percentChange = 0
				Else
					yearlyOpen = ws.Range("C" & tickerOpeningRow)
					percentChange = yearlyChange / yearlyOpen
				End If
				ws.Range("K" & summaryTableRow).NumberFormat = "0.00%"
				ws.Range("K" & summaryTableRow).Value = percentChange
				
				' Conditional Formatting Green/Red
				If ws.Range("J" & summaryTableRow).Value >= 0 Then
					ws.Range("J" & summaryTableRow).Interior.Color = RGB(0, 255, 0)
				Else
					ws.Range("J" & summaryTableRow).Interior.Color = RGB(255, 0, 0)
				End If
				
				' Prepare next row for Summary Table
				summaryTableRow = summaryTableRow + 1
				tickerOpeningRow = i + 1
				End If
			Next i
			
			' Challenge Section
			' Summary Table Row Length Determination
			lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
			
			' Shoop da Loop
			For i = 2 to lastRow
				If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
					ws.Range("Q2").Value = ws.Range("K" & i).Value
					ws.Range("P2").Value = ws.Range("I" & i).Value
				End If
				
				If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
					ws.Range("Q3").Value = ws.Range("K" & i).Value
					ws.Range("P3").Value = ws.Range("I" & i).Value
				End If
				
				If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
					ws.Range("Q4").Value = ws.Range("L" & i).Value
					ws.Range("P4").Value = ws.Range("I" & i).Value
				End If
			Next i

			' Formatting For Final Presentation
			ws.Range("Q2").NumberFormat = "0.00%"
			ws.Range("Q3").NumberFormat = "0.00%"
			ws.Columns("A:Q").AutoFit
			
		Next ws
		
End Sub
