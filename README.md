#Solution For My VBA Project Titled: VBA Stock Market Analysis by Anoje Janathanan  

#The Purpose of this code is to analyze the provided stock data across the three provided worksheets in Microsoft Excel using VBA. This code loops throughout all 3 worksheets and calculates different metrics per stock (Percent Change, Yearly Change, and Total Stock Volume) based off of the tickers provided. Moreover, the data results from the script are summarized/outputted in the table alongside extra information for further analysis purposes: These include 'Greatest % Increase', 'Greatest % Decrease', and 'Greatest Total Volume'. Once again, the script is put together in a way that the information is presented in correspondence to each sheet. Here is the code with comments provided for further clarification.





















Sub ModuleTwo() 'Initialize variables 'Created by Anoje J
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Highest_Ticker_Result As String
    Dim Lowest_Ticker_Result As String
    Dim Highest_Volume_Ticker_Result As String
    Dim Summary_Table_Row As Integer
    Dim lastRow As Long
    Dim OpenedPrice As Double
    Dim ClosedPrice As Double
    Dim PercentageChange As Double
    Dim YearlyChange As Double
    Dim TS_Volume As Double
    Dim Highest_Result As Double
    Dim Lowest_Result As Double
    Dim Highest_Volume As Double

    For Each ws In ThisWorkbook.Worksheets  'This for loop permits this body of code to loop throughout all of the worksheets
        Highest_Result = 0
        Lowest_Result = 0
        Highest_Volume = 0
        TS_Volume = 0 'TS represents total stock volume
        Summary_Table_Row = 2  'Initializes table row for data summary
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

        OpenedPrice = ws.Cells(2, 3).Value 'This is a given cell in the spreadsheet, where Open price = Column 3, row 2. Which is already provided in the workbook
        
        For i = 2 To lastRow 'Loops/Sums according to each row in column 7 which is the volume column
            Ticker = ws.Cells(i, 1).Value
            TS_Volume = TS_Volume + ws.Cells(i, 7).Value

            If ws.Cells(i + 1, 1).Value <> Ticker Then
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ClosedPrice = ws.Cells(i, 6).Value

                If OpenedPrice <> 0 Then   'Allows the calculation of the % change as long as the open price is not equal to zero
                    PercentageChange = ((ClosedPrice - OpenedPrice) / OpenedPrice) * 100 'Percent Change equates to this formula
                Else
                    PercentageChange = 0
                End If

                YearlyChange = ClosedPrice - OpenedPrice  'This represents that yearly change result is based off the difference of closed price and opened price

                ws.Range("J" & Summary_Table_Row).Value = YearlyChange 'With respect to indexing, Column 11 is where yearly change data is assigned to on the spreadsheet
                ws.Range("K" & Summary_Table_Row).Value = PercentageChange 'With respect to indexing, Column 12 is where percent change data is assigned on the spreadsheet
                ws.Range("L" & Summary_Table_Row).Value = TS_Volume 'With respect to indexing, Column 13 is where the total stock volume data is assigned to on the spreadsheet

                If YearlyChange <= 0 Then     'If the yearly change is less than or equal to zero, the color in column 11 is set to red. Vice versa, otherwise, assign green
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Range("J" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0)
                End If

                If PercentageChange > Highest_Result Then 'These codes are set to showcase the percentages as long as they satisfy the conditions below and updates the values accordingly
                    Highest_Result = PercentageChange
                    Highest_Ticker_Result = Ticker
                End If

                If PercentageChange < Lowest_Result Then
                    Lowest_Result = PercentageChange
                    Lowest_Ticker_Result = Ticker
                End If

                If TS_Volume > Highest_Volume Then
                    Highest_Volume = TS_Volume
                    Highest_Volume_Ticker_Result = Ticker
                End If

                PercentageChange = 0
                OpenedPrice = ws.Cells(i + 1, 3).Value
                Summary_Table_Row = Summary_Table_Row + 1
                TS_Volume = 0
            End If
        Next i

        ws.Range("S3").Value = Format(Highest_Result, "0.00") & "%"  'Formats and Assigns the values to the cells listed
        ws.Range("S4").Value = Format(Lowest_Result, "0.00") & "%"
        ws.Range("S5").Value = Highest_Volume
        ws.Range("R3").Value = Highest_Ticker_Result 'Greatest result
        ws.Range("R4").Value = Lowest_Ticker_Result 'Below 0, as its a negative percentage
        ws.Range("R5").Value = Highest_Volume_Ticker_Result 'Greatest volume

        ws.Range("A1:L1").Offset(0, 0).HorizontalAlignment = xlCenter 'Horizontally aligns the text accordingly, excludes the percentage increase/decrease values on the far right under columns R/S
    Next ws
End Sub
