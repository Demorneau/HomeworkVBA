Attribute VB_Name = "Module1"
Sub Stock_Volume_Count()
' Laura De Morneau Feb.26.2019 Homework 2
'Calculation of the Total Stock Volume for each Ticker ID on 3 worksheets in one workbook
'this work includes the Yearly Change for the open-close of stock market and the Greatest
'and Lowest Percentage Change.

'Dimension assigned by variables
'Amount of worksheets is indicated by
Dim Sheet_Count As Integer
'Amount of volume added based on the label ticker
Dim Total_Stock_Volume As Double
'Ticker Data are characters therefore they are assigned as strings
Dim Ticker_ID As String
'New Ticker will be used to store Ticker ID for the Greatest Total Volume
Dim New_Ticker As String
'Ticker IP to store data for Greatest Percentage Increase Ticker ID
Dim Ticker_IP As String
'Ticker Dp for Greatest Percentage Decrease Ticker ID
Dim Dp_Ticker As String
'Number of lines in a column
Dim Line_Count_Max As Long
Dim Record_Count As Long
'Row count for the heading
Dim Position_Count_R As Integer
'Column count for the heading
Dim Postion_Count_C As Integer
'Assignation for the counter of rows
Dim row_ct As Long
'Arrays to store Total Value Data per each active sheet
Dim Greatest_Total_Value(3) As Double
Dim Greatest_Percent_Increase(3) As Double
Dim Greatest_Percent_Decrease(3) As Double
'Single Maximun or Minimun value
Dim Greatest_Value As Double
Dim Percent_Increase As Double
Dim Percent_Decrease As Double
'Assigning values for open and close Market Yearly Change
Dim open_value As Double
Dim end_value As Double
'To run the last for loop
Dim i As Double
open_value = 0
end_value = 0
Sheet_Count = 1
'Coding first for to activate one by one a worksheet
For Sheet_Count = 1 To 3
    Worksheets(Sheet_Count).Activate
    'Position for the header first value (2,9)
    Position_Count_R = 2
    Position_Count_C = 9
    row_ct = 2
    Total_Stock_Volume = 0
    open_value = 0
    end_value = 0
    'Data Characters for the Ticker column
    Ticker_ID = Cells(row_ct, 1).Value
    Line_Count_Max = WorksheetFunction.CountA(Range("A2:A797711"))
    'Second for to move inside a worksheet
    For Record_Count = 2 To (Line_Count_Max - 1)
        Cells(1, 9) = "Ticker ID"
        Cells(1, 10) = "Yearly Change"
        Cells(1, 11) = "Percent Change"
        Cells(1, 12) = "Total Stock Volume"
        open_value = Cells(row_ct, 3).Value
        'Addition of value while the ticker label remains the same
            Do While Cells(row_ct, 1).Value = Ticker_ID
            Total_Stock_Volume = Total_Stock_Volume + Cells(row_ct, 7).Value
            row_ct = row_ct + 1
            Loop
        end_value = Cells((row_ct - 1), 6).Value
        Cells(Position_Count_R, Position_Count_C) = Ticker_ID
        Cells(Position_Count_R, (Position_Count_C + 3)) = Total_Stock_Volume
        Cells(Position_Count_R, (Position_Count_C + 1)) = (end_value - open_value)
            If open_value > 0 Then
                Cells(Position_Count_R, (Position_Count_C + 2)) = ((end_value - open_value) / open_value)
                'Calculate percent change if close is not zero
                'Placing select color in cells based on greater or lower than zero
                If (end_value - open_value) > 0 Then
                    'GREEN cells color for values greater than zero
                    Cells(Position_Count_R, (Position_Count_C + 1)).Interior.Color = RGB(198, 212, 60)
                    Else
                    'RED cells color for values less than zero (zero values cells are blank)
                    Cells(Position_Count_R, (Position_Count_C + 1)).Interior.Color = RGB(255, 0, 0)
                End If
            
                Else
                Cells(Position_Count_R, (Position_Count_C + 2)) = 0
            End If
            
            Total_Stock_Volume = 0
            Position_Count_R = Position_Count_R + 1
            Ticker_ID = Cells(row_ct, 1).Value
            Record_Count = row_ct
            
           
            
        Next

Worksheets(Sheet_Count).Columns("J:Q").AutoFit
Greatest_Total_Value(Sheet_Count) = WorksheetFunction.Max(Range("L2:L43398"))
Greatest_Percent_Increase(Sheet_Count) = WorksheetFunction.Max(Range("K2:K3200"))
Greatest_Percent_Decrease(Sheet_Count) = WorksheetFunction.Min(Range("K2:K3200"))



Next
'Calculating the maximun and minimun increase and decrease
For Sheet_Count = 1 To 3
Cells(Sheet_Count + 2, 19) = Greatest_Total_Value(Sheet_Count)
Cells(Sheet_Count + 2, 20) = Greatest_Percent_Decrease(Sheet_Count)
Cells(Sheet_Count + 2, 21) = Greatest_Percent_Increase(Sheet_Count)
Greatest_Value = WorksheetFunction.Max(Range("S:S"))
Percent_Increase = WorksheetFunction.Max(Range("U:U"))
Percent_Decrease = WorksheetFunction.Min(Range("T:T"))

Next

'Third and fourth for loop to find the corresponding ticker label to maximun and minimun increase or decrease
For Sheet_Count = 1 To 3
Worksheets(Sheet_Count).Activate
Line_Count_Max = WorksheetFunction.CountA(Range("A2:A797711"))
    For i = 2 To (Line_Count_Max - 1)

    If Greatest_Value = Cells(i, 12).Value Then
    New_Ticker = Cells(i, 9).Value
    ElseIf Percent_Increase = Cells(i, 11).Value Then
    Ticker_IP = Cells(i, 9).Value
    End If
    
    If Percent_Decrease = Cells(i, 11).Value Then
    Dp_Ticker = Cells(i, 9).Value
    End If
              
    Next
    
Next

Cells(1, 15) = "                      "
Cells(1, 16) = "                      "
Cells(1, 17) = "                      "
Cells(1, 16) = "Ticker ID"
Cells(1, 17) = "Value        "
Cells(2, 15) = "Greater % Increase"
Cells(3, 15) = "Greater % Decrease"
Cells(4, 15) = "Greatest Total Volume"
Cells(4, 17) = Greatest_Value
Cells(4, 16) = New_Ticker
Cells(2, 17) = Percent_Increase * 100 & "%"
Cells(2, 16) = Ticker_IP
Cells(3, 17) = Percent_Decrease * 100 & "%"
Cells(3, 16) = Dp_Ticker

'to clear columns needed to store arrays
Columns("S:U").Clear



End Sub

