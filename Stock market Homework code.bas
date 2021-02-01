Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim ws As Worksheet

'Loop through all woorksheets

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
   
    
    Dim Ticker As String
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 4
    Dim Yearly_Change As Double
        Yearly_Change = 0
    Dim Yearly_Change_Percent As Double
        Yearly_Change_Percent = 0
    Dim Total_Stock_Volume As LongLong
         Total_Stock_Volume = 0
    
  Cells(3, 10) = "Ticker Symbol"
  Cells(3, 11) = "Yearly Change"
  Cells(3, 12) = "% Change"
  Cells(3, 13) = "Total Stock Volume"
  
   Range("J3:M3").Font.Bold = True

  For i = 2 To Range("A1").End(xlDown).Row
 
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        Ticker = Cells(i, 1).Value
        
         Yearly_Change = Yearly_Change + (Cells(i, 3).Value - Cells(i, 6).Value)
        
        If (Cells(i, 3) = 0 And Cells(i, 6) = 0) Then
            Yearly_Change_Percent = 0
            Else
            Yearly_Change_Percent = Yearly_Change_Percent + ((Cells(i, 3).Value - Cells(i, 6).Value) / Cells(i, 3).Value)
            
        End If
        
        Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
            Range("J" & Summary_Table_Row).Value = Ticker
            
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            
            Range("L" & Summary_Table_Row).Value = Yearly_Change_Percent
            
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
            
            Range("M" & Summary_Table_Row).Value = Total_Stock_Volume
                 
                 Range("M" & Summary_Table_Row).NumberFormat = "#,###"
             
        Summary_Table_Row = Summary_Table_Row + 1
            
            Yearly_Change = 0
            
            Yearly_Change_Percent = 0
            
            Total_Stock_Volume = 0
            
        Else
             Yearly_Change = Yearly_Change + (Cells(i, 3).Value - Cells(i, 6).Value)
        
              If (Cells(i, 3) = 0 And Cells(i, 6) = 0) Then
                Yearly_Change_Percent = 0
                Else
                    Yearly_Change_Percent = Yearly_Change_Percent + ((Cells(i, 3).Value - Cells(i, 6).Value) / Cells(i, 3).Value)

              End If
        
             Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
        
            
            
        
            
    End If
    If Range("K" & Summary_Table_Row).Value > 0 Then
    
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        
        Else
        Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
 Next i
 
Next ws

    MsgBox ("All worksheets complete")

End Sub

