VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_checker_all_tabs()

    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim RowCounter As Long
    Dim j As Long
    Dim changeColor As Long
    
   
    For Each ws In ThisWorkbook.Worksheets
       
        If InStr(ws.Name, "Q") > 0 Then
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Quarterly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
            QuarterlyChange = 0
            PercentChange = 0
            TotalStockVolume = 0
            RowCounter = 0
            j = 2
            
            For i = 2 To LastRow
                RowCounter = RowCounter + 1
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                ClosingPrice = ws.Cells(i, 6).Value
                
               
                If Ticker <> ws.Cells(i - 1, 1).Value Then
               
                    If RowCounter > 1 Then
                        
                        QuarterlyChange = ClosingPrice - OpeningPrice
                        If OpeningPrice <> 0 Then
                            PercentChange = (QuarterlyChange / OpeningPrice) * 100
                        Else
                            PercentChange = 0
                        End If
                        
                      
                        ws.Cells(j, 9).Value = Ticker
                        ws.Cells(j, 10).Value = QuarterlyChange
                        ws.Cells(j, 11).Value = PercentChange
                        ws.Cells(j, 12).Value = TotalStockVolume
                        
                        
                        If QuarterlyChange > 0 Then
                            changeColor = RGB(0, 255, 0)
                        ElseIf QuarterlyChange < 0 Then
                            changeColor = RGB(255, 0, 0)
                        Else
                            changeColor = RGB(255, 255, 255)
                        End If
                        ws.Cells(j, 10).Interior.Color = changeColor
                        
                        j = j + 1
                    End If
                
                    TotalStockVolume = 0
                    RowCounter = 0
                End If
                
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            Next i
        End If
    Next ws
End Sub
