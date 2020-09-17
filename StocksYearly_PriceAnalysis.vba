Attribute VB_Name = "Module1"
Sub VbaChallenge()

For Each ws In Worksheets
        ws.Activate
        Call SetTitle
       Call CalculateSummary
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
Dim year_open_price     As Integer
Dim year_close_price    As Integer
Dim percentage_change   As Integer
Dim yearly_price_change As Integer
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
                percentage_change = (yearly_price_change * 100) / year_close_price
                
                
                Cells(cnt, "K").Value = Cells(counter, 6).Value
                ticker_YearFirst_day_closePrice = 0
                Cells(cnt, "I").Value = ticker
                Cells(cnt, "J").Value = yearly_price_change
                Cells(cnt, "K").Value = percentage_change
                Cells(cnt, "L").Value = totalvolume
                      
                cnt = cnt + 1
                totalvolume = 0
            
            
            End If
    
    
    Next counter

   

'Next wks


End Sub

