Attribute VB_Name = "Module1"
Sub VBA_Challenge():

Dim headers() As Variant

Dim MainWs As Worksheet

Dim wb As Workbook

Set wb = ActiveWorkbook

headers() = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly_Change", "Percent_Change", "Stock_Volume", " ", " ", " ", "Ticker", "Value")

For Each MainWs In wb.Sheets
    With MainWs
    .Rows(1).Value = ""
    For i = LBound(headers()) To UBound(headers())
    .Cells(1, 1 + i).Value = headers(i)
    
    Next i
    .Rows(1).Font.Bold = True
    End With
    
Next MainWs

    For Each MainWs In Worksheets
    
    
    Dim Ticker_Name As String
    Ticker_Name = " "
    
    Dim Total_Ticker_Volume As Double
    Total_Ticker_Volume = 0
    
    Dim Start_Price As Double
    Start_Price = 0
    
    Dim Close_Price As Double
    Close_Price = 0
    
    Dim Yearly_Price_Change As Double
    Yearly_Price_Change = 0
    
    Dim Yearly_Price_Change_Percent As Double
    Yearly_Price_Change_Percent = 0
    
    Dim Max_Ticker_Name As String
    Max_Ticker_Name = " "
    
    Dim Min_Ticker_Name As String
    Min_Ticker_Name = " "
    
    Dim Max_Percent As Double
    Max_Percent = 0
    
    Dim Min_Percent As Double
    Min_Percent = 0
    
    Dim Max_Volume_Ticker_Name As String
    Max_Volume_Ticker_Name = " "
    
    Dim Max_Volume As Double
    Max_Volume = 0
    
    Dim Table_Row As Long
    Table_Row = 2
    
    Dim Lastrow As Long
     Lastrow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row
     
    Start_Price = MainWs.Cells(2, 3).Value
    
    For i = 2 To Lastrow
    
        If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
            Ticker_Name = MainWs.Cells(i, 1).Value
            
            Close_Price = MainWs.Cells(i, 6).Value
            Yearly_Price_Change = Close_Price - Start_Price
            
            If Start_Price <> 0 Then
                Yearly_Price_Change_Percent = (Yearly_Price_Change / Start_Price) * 100
                
            End If
            
                Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
                
            MainWs.Range("I" & Table_Row).Value = Ticker_Name
                
            MainWs.Range("J" & Table_Row).Value = Yearly_Price_Change
            
            If (Yearly_Price_Change > 0) Then
             MainWs.Range("J" & Table_Row).Interior.ColorIndex = 4
             
            ElseIf (Yearly_Price_Change <= 0) Then
             MainWs.Range("J" & Table_Row).Interior.ColorIndex = 3
            End If
            
            MainWs.Range("K" & Table_Row).Value = (CStr(Yearly_Price_Change_Percent) & "%")
            MainWs.Range("L" & Table_Row).Value = Total_Ticker_Volume
            
            Table_Row = Table_Row + 1
            
            Start_Price = MainWs.Cells(i + 1, 3).Value
            
            If (Yearly_Price_Change_Percent > Max_Percent) Then
             Max_Percent = Yearly_Price_Change_Percent
             Max_Ticker_Name = Ticker_Name
             
            ElseIf (Yearly_Price_Change_Percent < Min_Percent) Then
             Min_Percent = Yearly_Price_Change_Percent
             Min_Ticker_Name = Ticker_Name
             
            End If
            
            If (Total_Ticker_Volume > Max_Volume) Then
             Max_Volume = Total_Ticker_Volume
             Max_Volume_Ticker_Name = Ticker_Name
             
             End If
             
            Yearly_Price_Change_Percent = 0
            Total_Ticker_Volume = 0
            
        Else
        
            Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
            
        End If
        
        
    
    Next i
    
        MainWs.Range("Q2").Value = (CStr(Max_Percent) & "%")
        MainWs.Range("Q3").Value = (CStr(Min_Percent) & "%")
        MainWs.Range("Q4").Value = Max_Volume
        MainWs.Range("P2").Value = Max_Ticker_Name
        MainWs.Range("P3").Value = Min_Ticker_Name
        MainWs.Range("P4").Value = Max_Volume_Ticker_Name
        MainWs.Range("O2").Value = "Greatest % Increase"
        MainWs.Range("O3").Value = "Greatest % Decrease"
        MainWs.Range("O4").Value = "Greatest Total Volume"
        
    Next MainWs
            
            
End Sub

