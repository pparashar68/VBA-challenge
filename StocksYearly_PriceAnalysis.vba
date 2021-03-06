Attribute VB_Name = "Module1"
Sub VbaChallenge()

For Each ws In Worksheets
        ws.Activate
        Call SetTitle                               'This subroutine initialize the cells and add Title heading to workshees or workbook
        Call CalculateSummary              ' This subroutine go through each rows stocks values and calculate the summary of stock tickers
        Call GreatestChange                  ' This subroutine calculate the summary values of stocks for each year
    Next ws



End Sub

Sub SetTitle()
    Range("I:Q").Value = ""
    Range("I:Q").Interior.ColorIndex = 0
' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    'this is for challenge only
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("I:O").Columns.AutoFit
End Sub

Sub CalculateSummary()
Dim ticker As String
Dim totalvolume              As Double
Dim startcounter         As Double
Dim endcounter           As Double
Dim year_open_price     As Double
Dim year_close_price    As Double
Dim percentage_change   As Double
Dim yearly_price_change As Double
Dim counter As Double


Dim cnt As Double



    startcounter = 2
    year_open_price = 0
    year_close_price = 0
    percentage_change = 0
    totalvolume = 0
    ticker = ""
    yearly_price_change = 0
    
    endcounter = Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (endcounter)
    cnt = 2
    totalvolume = 0
    year_open = 0
    
 
    For counter = startcounter To endcounter
            
            If year_open_price = 0 Then
                year_open_price = Cells(counter, 3).Value
                
            End If
            
            ticker = Cells(counter, 1).Value
        
            
            If Cells(counter, 1).Value = Cells(counter + 1, 1) Then
                
                totalvolume = totalvolume + Cells(counter, 7).Value
                
            Else

                year_close_price = Cells(counter, 6).Value
                
                yearly_price_change = year_close_price - year_open_price
                If year_close_price = 0 And year_open_price <> 0 Then
                    percentage_change = -100
                ElseIf year_close_price = 0 And year_open_price = 0 Then
                    percentage_change = 0
                Else
                    percentage_change = (yearly_price_change) / year_open_price
                End If
                
                
                
                
                         Cells(cnt, "K").Value = Cells(counter, 6).Value
                         Cells(cnt, "I").Value = ticker
                         Cells(cnt, "J").Value = yearly_price_change
                         
                         ' format cells positive as green and negative change as red base on cell value
                                If Cells(cnt, "J").Value > 0 Then
                                            Cells(cnt, "J").Interior.Color = vbGreen
                                ElseIf Cells(cnt, "J").Value < 0 Then
                                           Cells(cnt, "J").Interior.Color = vbRed
                                Else
                                           Cells(cnt, "J").Interior.Color = vbBlue
                                
                                End If
                         
                         Cells(cnt, "K").Value = percentage_change
                         
                         Cells(cnt, "K").NumberFormat = "00.00%"
                        
                        ' format cells positive as green and negative change as red base on cell value
                                 If Cells(cnt, "K").Value > 0 Then
                                            Cells(cnt, "K").Interior.Color = vbGreen
                                ElseIf Cells(cnt, "K").Value < 0 Then
                                           Cells(cnt, "K").Interior.Color = vbRed
                                Else
                                           Cells(cnt, "K").Interior.Color = vbBlue
                                
                                End If
                         Cells(cnt, "L").Value = totalvolume
                               
                          ' initialize variables for next ticker of the stock in the spreadsheet
                                             cnt = cnt + 1
                                             totalvolume = 0
                                             year_close_price = 0
                                             yearly_price_change = 0
                                             percentage_change = 0
                                             year_open_price = 0
                                
            
            End If
    
    
    
    Next counter




End Sub

Sub GreatestChange()

Dim greatest_percentage_increased  As Double
Dim greatest_percentage_decreased   As Double
Dim greatest_total_volume   As Double
Dim total_rows  As Double
Dim maxValue As Double
Dim rng As Range
Dim rng_volume  As Range


total_rows = Cells(Rows.Count, "I").End(xlUp).Row
start_rows = 2
greatest_percentage_increased = Application.WorksheetFunction.Max(Columns("K"))
greatest_percentage_decreased = Application.WorksheetFunction.Min(Columns("K"))
greatest_total_volume = Application.WorksheetFunction.Max(Columns("L"))
Range("Q2").Value = greatest_percentage_increased
If Range("Q2").Value > 0 Then
        Range("Q2").Interior.Color = vbGreen
        Else
        Range("Q2").Interior.Color = vbRed
End If
Range("Q3").Value = greatest_percentage_decreased
If Range("Q3").Value > 0 Then
Range("Q3").Interior.Color = vbGreen
Else
Range("Q3").Interior.Color = vbRed
End If

Cells(2, "Q").NumberFormat = "00.00%"
Cells(3, "Q").NumberFormat = "00.00%"
Range("Q4").Value = greatest_total_volume

Range("P2").Value = "=INDEX(I2" & ":I" & total_rows & ",MATCH(Q" & 2 & ",K" & 2 & ":K" & total_rows & ",0))"
Range("P3").Value = "=INDEX(I2" & ":I" & total_rows & ",MATCH(Q" & 3 & ",K" & 2 & ":K" & total_rows & ",0))"
Range("P4").Value = "=INDEX(I2" & ":I" & total_rows & ",MATCH(Q" & 4 & ",L" & 2 & ":L" & total_rows & ",0))"
Cells(4, "Q").NumberFormat = "0,000"

    

    

                                


End Sub


