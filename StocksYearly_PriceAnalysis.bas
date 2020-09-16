Attribute VB_Name = "Module1"
Sub stocksummary()
Dim ticker As String
Dim YearFirstDay As String
Dim YearLastDay As String
Dim totalvolume  As Double
Dim pricechange As Integer
Dim startcounter As Integer
Dim endcounter    As Double
Dim ticker_YearFirst_day_closePrice As Integer
Dim cnt As Double

' WorkSheets variables to check all the worksheets in workbook
Dim wbk As Workbook
Dim wks As Worksheet

'Set wbk = ThisWorkbook

'For Each wks In wbk.Worksheets              ' this code is going to execute for each worksheet in the workbook
'    wks.Name = wks.Name & "Test"

    startcounter = 2
    endcounter = Cells(Rows.Count, "A").End(xlUp).Row
    'MsgBox (endcounter)
    cnt = 2
    totalvolume = 0
 
    For counter = startcounter To endcounter
            ticker = Cells(counter, 1).Value
            YearFirstDay = Cells(counter, 2).Value
            ticker_YearFirst_day_closePrice = Cells(counter, 6).Value
            'totalvolume = Cells(counter, 7).Value

            
            If Cells(counter, 1).Value = Cells(counter + 1, 1) Then
                
                totalvolume = totalvolume + Cells(counter, 7).Value
                'Cells(counter, "Q").Value = Cells(counter, 1).Value
                
            Else

                Cells(cnt, "I").Value = ticker
                Cells(cnt, "N").Value = totalvolume
                Cells(cnt, "J").Value = ticker_YearFirst_day_closePrice
                Cells(cnt, "K").Value = Cells(counter, 6).Value
                ticker_YearFirst_day_closePrice = 0
                
                
                cnt = cnt + 1
            
            
            End If
    
    
    Next counter

   

'Next wks


End Sub
