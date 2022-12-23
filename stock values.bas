Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data():

Dim MainWs As Worksheet

For Each MainWs In Worksheets

Dim Ticker_Name As String

Dim Max_Ticker_Name As String

Dim Min_Ticker_Name As String

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim Total_Stock_Volume As Integer
Total_Stock_Volume = 0

Dim Stock_Open As Double
Stock_Open = 0

Dim Max_Stock As Double
Max_Stock = 0

Dim Min_Stock As Double
Min_Stock = 0

Dim Stock_Close As Double
Stock_Close = 0

Dim Max_Percent As Double
Max_Percent = 0

Dim Min_Percent As Double
Min_Percent = 0


Dim mainLastRow As Long
Dim mainLastCol As Long

Next MainWs


Dim Summary_Table_Row As Long
Summary_Table_Row = 2

Dim Lastrow As Long


Lastrow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row

Stock_Open = MainWs.Cells(2, 3).Value


    For i = 2 To Lastrow
    
        If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
    
            Ticker_Name = MainWs.Cells(i, 1).Value
    
             Stock_Close = MainWs.Cells(i, 6).Value
             Stock_Open = MainWs.Cells(i, 3).Value
            yearly_change = Stock_Close - Stock_Open
    
            If Stock_Open <> 0 Then
                percent_change = (yearly_change / Stock_Open) * 100
        
            End If
    
            Total_Stock_Volume = Total_Stock_Volume + MainWs.Cells(i, 7).Value
    
            MainWs.Range("I" & Summary_Table_Row).Value = Ticker_Name
    
            MainWs.Range("J" & Summary_Table_Row).Value = yearly_change
    
        ' this is green for positive'
        
         If (yearly_change > 0) Then
         MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
        ' this is for negative'
        
            ElseIf (yearly_change < 0) Then
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            Else
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 0
            
            
        End If
        
            
            MainWs.Range("J" & Summary_TableRow).Value = Total_Stock_Volume
    
            Summary_Table_Row = Summary_Table_Row + 1
    
            Stock_Open = MainWs.Cells(i + 1, 3).Value
    
            Change_Percent = 0
            Ticker_Volume = 0
        
   
   
   Else: Total_Stock_Volume = Total_Stock_Volume + MainWs.Cells(i, 7).Value
        
   End If
    
   Next i
   
   ' finding the min and max percent change'
   
   
    If (yearly_change > Max_Percent) Then
            Min_Percent = percent_change
            Max_Stock = Ticker_Name
    
        ElseIf (yearly_change < Min_Percent) Then
            Min_Percent = percent_change
             Min_Stock = Ticker_Name
    
        End If
    
    If (Total_Stock_Volume > Max_Stock) Then
        Max_Stock = Total_Stock_Volume
        Max_Ticker_Name = Ticker_Name
        
    End If
    
    If (Ticker_Volume > Max_Stock) Then
        Max_Stock = Total_Stock_Volume
        Max_Ticker_Name = Ticker_Name

    End If
    
 



End Sub

