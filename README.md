# VBA-challenge
The VBA of Wall Street
Sub Stock_Data()

    'Declare Variables
    Dim WS As Worksheet
    Dim ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_ChangePrice As Double
    Yearly_ChangePrice = 0
    Dim Total_Volume As Double
    Total_Volume = 0
    Dim YrChange_Percent As Double
    Dim CountofRows As Long
    Dim GreatestPer_Increase As Double
    Dim GreatestPer_IncreaseTicker As String
    Dim GreatestPer_Decrease As Double
    Dim GreatestPer_DecreaseTicker As String
    Dim Greatest_Increase_Volume As Double
    Dim Greatest_Increase_VolumeTicker As String
    'Count number of sheets in this workbook
    Dim CountofSheets As Long
    Dim WsWorkbook As Workbook
    CountofSheets = Worksheets.Count
    'Sum of Valume for each unique ticker
    Dim SumofTotalVolume As Double
    Dim Summary_Table_Row As Integer 'keep track of the value and store in the summary table
    Summary_Table_Row = 2
  
    'For loop for counting sheets
    For Each WS In Worksheets  'for loop to run the script for all the worksheets in the workbook
                      
        CountofRows = WS.Cells(Rows.Count, "A").End(xlUp).Row
            
        'Print the headers in the summary table
        WS.Range("I1").Value = "Ticker"
        WS.Range("J1").Value = "Yearly_Change"
        WS.Range("K1").Value = "Percent_Change"
        WS.Range("L1").Value = "Total_Stock_Volume"
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("O3").Value = "Greatest % Decrease"
        WS.Range("O4").Value = "Greatest Total Volume"
        WS.Range("O2").Value = "Greatest % Increase"
        WS.Range("P1").Value = "Ticker"
        WS.Range("Q1").Value = "Value"
            
        'Reset the sum once the ticker is found
        Open_Price = 0
        Close_Price = Range("F2") 'get the next open price
        SumofTotalVolume = 0
        Summary_Table_Row = 2
        GreatestPer_Increase = 0
        GreatestPer_Decrease = 0
            
        'For loop to count of rows
        For i = 2 To CountofRows
              
            If Open_Price = 0 Then
                Open_Price = Range("c" & i) 'get new open price for new ticker
            End If
            'Sum of unique ticker for total volume
            SumofTotalVolume = SumofTotalVolume + WS.Cells(i, 7)  'store the total volume for each ticker
            'If condition to go throw each row until the new ticker is found
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                'If the new ticker is found, print the previous ticker, yearly change in price and total volume in summary table
                ticker = WS.Cells(i, 1).Value
                WS.Range("I" & Summary_Table_Row).Value = ticker
                Close_Price = WS.Cells(i, 6).Value
                Yearly_ChangePrice = Close_Price - Open_Price 'calculate the yearly change in price
                WS.Range("J" & Summary_Table_Row).Value = Yearly_ChangePrice
                WS.Range("L" & Summary_Table_Row).Value = SumofTotalVolume
                'to find if there is Zero value for Open price. Since we can not devide by Zero
            If Open_Price <> 0 Then
                'Calculate Yearly change in Percent and store the value
                YrChange_Percent = (Yearly_ChangePrice) / Open_Price
                WS.Range("K" & Summary_Table_Row).Value = YrChange_Percent 'print yearly change in percentage
            If (YrChange_Percent) >= 0 Then
                WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4 'change it to green for all positive change
            End If
            If (YrChange_Percent) < 0 Then
                WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3 'change it to red for all negative change
            End If
            Else
                YrChange_Percent = 0 'reset yearly change in percent to 0
            End If
            If (YrChange_Percent > GreatestPer_Increase) Then 'find the greatest incease in percent and print
                GreatestPer_Increase = YrChange_Percent
                GreatestPer_IncreaseTicker = ticker
            ElseIf (YrChange_Percent < GreatestPer_Decrease) Then 'find the greatest decrease in percent and print
                GreatestPer_Decrease = YrChange_Percent
                GreatestPer_DecreaseTicker = ticker
            End If
            If (SumofTotalVolume > Greatest_Increase_Volume) Then 'find the greatest volume and print
                Greatest_Increase_Volume = SumofTotalVolume
                Greatest_Increase_VolumeTicker = ticker
            End If
                'print yearly change percent and format to %
                 WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%" 'format the percent change
                        
                Open_Price = 0
                Close_Price = WS.Cells(i + 1, 6).Value
                SumofTotalVolume = 0
                Summary_Table_Row = Summary_Table_Row + 1
            End If
                  
            Next i
                'Print value in summary table
                WS.Range("P2").Value = GreatestPer_IncreaseTicker
                WS.Range("P3").Value = GreatestPer_DecreaseTicker
                WS.Range("Q2").Value = GreatestPer_Increase
                WS.Range("Q2").NumberFormat = "0.00%"
                WS.Range("Q3").Value = GreatestPer_Decrease
                WS.Range("Q3").NumberFormat = "0.00%"
                WS.Range("Q4").Value = Greatest_Increase_Volume
    Next
    

    

End Sub
