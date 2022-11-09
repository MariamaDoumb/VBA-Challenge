Attribute VB_Name = "Module1"
Sub Stock_Data()
    [H1:I1] = [{"Ticker","Total Stock Volume"}]

Dim Ticker As String
Dim Summary_Table As Integer
Summary_Table = 2
Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0


 Dim LastRow As Long
    LastRow = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row

    Dim i As Long, Total As Long
    
    For i = 2 To LastRow
    
     If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
     
     Ticker = Cells(i, 1).Value
     
     Range("H" & Summary_Table).Value = Ticker

     Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
     Range("I" & Summary_Table).Value = Total_Stock_Volume
     

   
     Summary_Table = Summary_Table + 1
     Total_Stock_Volume = 0
     
    End If
    
    Next i
    
End Sub
