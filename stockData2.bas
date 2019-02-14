Attribute VB_Name = "stockData2"
Sub stockData2()
    ' LOOP THROUGH ALL SHEETS or (Tabs)
    ' --------------------------------------------
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        ' Determine the Last Row in the worksheet
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Adding Heading for summary section
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        'Creating Variable to hold Value
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Initial Open Price in column C
        Open_Price = Cells(2, Column + 2).Value
         ' Loop through all ticker symbol in column A
        
        For i = 2 To LastRow
         ' Check to see if we are still within the same ticker symbol, if it is not then do something...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                ' Ready to Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Set Close Price from column F
                Close_Price = Cells(i, Column + 5).Value
                ' Calculating Yearly change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' This will Calculate Total Volumn
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' increase summary table row by Adding one to the summary table row
                Row = Row + 1
                ' Intialize open Price by reseting the Open Price
                Open_Price = Cells(i + 1, Column + 2)
                ' Initialize the volume Total by reseting the Volumn Total here
                Volume = 0
            'if cells are the same ticker (Checking to see if cells are the same ticker)
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
            'Repeat loop
        Next i
        
        ' Determining the Last Row of Yearly Change per WS - by the way YC stands for Yearly Change
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
        ' Setting the Cell Colors color index of 10 = Green, 3 = Red
        For j = 2 To YCLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
        
    Next WS
        
End Sub

