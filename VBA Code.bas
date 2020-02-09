Attribute VB_Name = "Module1"
Sub TakeSymbol()
    
    
    Dim WsName As String
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentChange As Double
    Dim TTSV As Double 'Total Stock Volume

    Dim YearlyChange As Double
    Dim LastRow As Long
    Dim t_r As Long  'ticker row index
    Dim st_r As Long 'sum ticker row index
    Dim st_c As Long 'sum ticker column index
    Dim op_r As Long 'OpenPrice index

    Dim G_In_Tic As String
    Dim G_De_Tic As String
    Dim G_TTSV_Tic As String
    Dim G_In_Val As Double
    Dim G_De_Val As Double
    Dim G_TTSV_Val As Double
    
    st_c = 9 'summary ticker index column
    op_r = 2 'open price index row
    
    G_In_Tic_r = 2
    G_In_Tic_c = 14
    G_In_r = 2
    G_In_c = 15
    G_De_r = 3
    G_De_c = 15
    G_TTSV_r = 4
    G_TTSV_c = 15
    
    
    

    
    For Each ws In Worksheets
    
            st_r = 2
            op_r = 2
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
                      
            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
             'LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row


        For t_r = 2 To LastRow
            
            If ws.Cells(t_r, 1).Value = ws.Cells(t_r + 1, 1) Then
                
                ' Collect Open Price
                OpenPrice = ws.Cells(op_r, 3).Value
                
                ' Collect Ticker
                Ticker = ws.Cells(t_r, 1).Value
                
                'Accumulate Stock_Volume
                TTSV = TTSV + ws.Cells(t_r, 7).Value
    
            Else
                'Collect Close Price
                ClosePrice = ws.Cells(t_r, 6).Value
                
                'Calculate Year Change
                YearlyChange = ClosePrice - OpenPrice
                
                'Accumulate Stock_Volumn
                TTSV = TTSV + ws.Cells(t_r, 7).Value
                
                'Open_Price_Row_Index =  next row
                op_r = t_r + 1
                
                'Summary_Table_Ticker = Last value of the same collection
                ws.Cells(st_r, st_c).Value = Ticker
                
                'Summary_Table_YearlyChange = Calculate Year Change
                ws.Cells(st_r, st_c + 1).Value = YearlyChange
                
                '-----------------------------
                'Avoid Error for devided by 0
                '-----------------------------
                If YearlyChange = 0 Or OpenPrice = 0 Then
                    ws.Cells(st_r, st_c + 2).Value = 0
                Else
                    PercentChange = (YearlyChange / OpenPrice) * 100
                    ws.Cells(st_r, st_c + 2).Value = Format((YearlyChange / OpenPrice), "Percent")
                    
                End If
                
                'End Avoid Error--------------
                
                
                '-----------------------------
                'Conditional Format Cells ----
                '-----------------------------
                If YearlyChange <= 0 Then
                    ws.Range("J" & st_r).Interior.ColorIndex = 3
                Else
                    ws.Range("J" & st_r).Interior.ColorIndex = 10
                End If
                
                'End Conditional Format Cells-
                
                'Total_Stock_Volume = Accumate Stock_Volume
                ws.Cells(st_r, st_c + 3).Value = TTSV
  
                
                'Set Stock_Volume to 0 for accumating next ticker
                TTSV = 0
                
                'Summary_Table Row_Indext + 1
                st_r = st_r + 1
                
             End If
                             

             
        Next t_r
        ws.Cells.EntireColumn.AutoFit
        
    Next ws
    
    Sheets(1).Activate
    
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Increase"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    Range("O2").Value = G_In_Tic
    Range("O3").Value = G_De_Tic
    Range("O4").Value = G_TTSV_Tic
    
    
    Range("P2").Value = G_In
    Range("P3").Value = G_De
    Range("P4").Value = G_TTSV_Val
    '---------------------
    'AutoFit Entire Column
    '---------------------
    Cells.EntireColumn.AutoFit
    
   For Each ws In Worksheets
        
        G_In_Val = Range("K2").Value
        G_In_Val = Range("K2").Value
        G_TTSV_Val = Range("L2").Value
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        For t_r = 2 To LastRow
        
            'ws.Range("N5").Value = (t_r)
            If ws.Range("K" & t_r).Value >= G_In_Val Then
                G_In_Val = ws.Range("K" & t_r).Value
                G_In_Tic = ws.Range("I" & t_r).Value
                'ws.Range("N6").Value = "G_In_Val : " & G_In_Val
                'ws.Range("O6").Value = "G_In_Tic : " & G_In_Tic
                
            End If
            
            If ws.Range("K" & t_r).Value < G_De_Val Then
                G_De_Val = ws.Range("K" & t_r).Value
                G_De_Tic = ws.Range("I" & t_r).Value
                'ws.Range("N7").Value = "G_De_Val : " & G_De_Val
                'ws.Range("O7").Value = "G_De_Tic : " & G_De_Tic
            End If
            
            If ws.Range("L" & t_r).Value >= G_TTSV_Val Then
                G_TTSV_Val = ws.Range("L" & t_r).Value
                G_TTSV_Tic = ws.Range("I" & t_r).Value
                'MsgBox (t_r)
                'MsgBox (G_TTSV_Tic)
                'MsgBox (G_TTSV_Val)
                'ws.Range("N8").Value = "G_TTSV_Val : " & G_TTSV_Val
                'ws.Range("O8").Value = "G_TTSV_Tic : " & G_TTSV_Tic
            End If
        Next t_r
        
    Range("O2").Value = G_In_Tic
    Range("P2").Value = Format(G_In_Val, "Standard")
    
    
    Range("O3").Value = G_De_Tic
    Range("P3").Value = Format(G_De_Val, "Standard")
    
    Range("O4").Value = G_TTSV_Tic
    Range("P4").Value = G_TTSV_Val
    
    Cells.EntireColumn.AutoFit
    
     
   Next ws
       
    
End Sub


