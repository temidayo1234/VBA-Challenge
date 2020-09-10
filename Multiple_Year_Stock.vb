Sub Multiple_Year_Stock()

    'Declare LastRow as a variable
    Dim LastRow As Long
    
    'Declare TickerHolder as a vaiable
    Dim TickerHolder As Integer
    
    'Declare Closing as a variable
    Dim Closing As Double
    
    'Declare Opening as a variable
    Dim Opening As Double
    
    'Declare Yearly_Change as a variable
    Dim Yearly_Change As Double
    
    'Declare Yearly_Change_Holder as variable
    Dim Yearly_Change_Holder As Integer
    
    'Declare Percent_Change_Holder as variable
    Dim Percent_Change As Double
    
    'Declare Percent_Change as variable
    Dim Percent_Change_Holder As Integer
        
    'Declare Volume_Total as variable
    'Dim Volume_Total As Long
    
    'Declare Volume_Total_Holder as variable
    Dim Volume_Total_Holder As Integer
    
    'Declare Greatest_Percent_Increase as variable
    Dim Greatest_Percent_Increase As Double
    
    'Declare Greatest_Percent_Decrease as variable
    Dim Greatest_Percent_Decrease As Double
    
                    
    'Parse through each worksheet
    For Each ws In Worksheets
        
        'Assign heading to greatest change increase section
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
                
        'Define LastRow to indicate end of row in worksheets
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Define TickerHolder to hold the destination for the ticker
        TickerHolder = 2
        
        'Define Yearly_Change_Holder to hold the destination for the Yearly_Change
        Yearly_Change_Holder = 2
        
        'Define Percent_Change_Holder to hold the destination for the Percent_Change
        Percent_Change_Holder = 2
        
        'Assigning value to the first opening value
        Opening = ws.Cells(2, 3).Value
        
        'Define Volume_Total_Holder
        Volume_Total_Holder = 2
        
        'Define Volume_Total
        Volume_Total = 0
        
        'Define Greatest_Percent_Increase
        Greatest_Percent_Increase = 0
        
        'Define Greatest_Percent_Decrease
        Greatest_Percent_Decrease = 0
        
        'Define Greatest_Percent_Decrease
        Greatest_Volume = 0
        
        'Parse through each row from A2 to the last row
        For i = 2 To LastRow
        
            'Compare a row to the future row to evaluate if they are no the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Moving each unique Ticker to a new location
                ws.Cells(TickerHolder, 9).Value = ws.Cells(i, 1).Value
                TickerHolder = TickerHolder + 1
            
                'Yearly change calculation
                Closing = ws.Cells(i, 6).Value
                Yearly_Change = Closing - Opening
                ws.Cells(Yearly_Change_Holder, 10).Value = Yearly_Change
                    
                    'Adding color formatting based on positive and negative values
                    If Yearly_Change < 0 Then
                        ws.Cells(Yearly_Change_Holder, 10).Interior.ColorIndex = 3
                    Else
                         ws.Cells(Yearly_Change_Holder, 10).Interior.ColorIndex = 4
                    End If
                    
                Yearly_Change_Holder = Yearly_Change_Holder + 1
            
                'Calculate the percent of the change
                If Opening = 0 Then
                    Percent_Change = 0
                    Percent_Change_Holder = Percent_Change_Holder + 1
                Else
                Percent_Change = (Yearly_Change / Opening) * 100
                ws.Cells(Percent_Change_Holder, 11).Value = Percent_Change & "%"
                Percent_Change_Holder = Percent_Change_Holder + 1
                
                End If
                
                
                If Percent_Change > Greatest_Percent_Increase Then
                    Greatest_Percent_Increase = Percent_Change
                    ws.Cells(2, 16).Value = Greatest_Percent_Increase
                    ws.Cells(2, 15).Value = ws.Cells(i, 1).Value
                    ws.Cells(2, 14).Value = "Greatest % Increase"
                 Else
                
                 End If
                
                 If Percent_Change < Greatest_Percent_Decrease Then
                    Greatest_Percent_Decrease = Percent_Change
                    ws.Cells(3, 16).Value = Greatest_Percent_Decrease
                    ws.Cells(3, 15).Value = ws.Cells(i, 1).Value
                    ws.Cells(3, 14).Value = "Greatest % Decrease"
                    
                 Else
                
                 End If
    
            
                'Add the last Volume total the Volume_Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                ws.Cells(Volume_Total_Holder, 12).Value = Volume_Total
                Volume_Total_Holder = Volume_Total_Holder + 1
                
                If Volume_Total > Greatest_Volume Then
                    Greatest_Volume = Volume_Total
                    ws.Cells(4, 16).Value = Greatest_Volume
                    ws.Cells(4, 15).Value = ws.Cells(i, 1).Value
                    ws.Cells(4, 14).Value = "Greatest Volume"
                Else
                
                End If
    
                Volume_Total = 0
                

                Opening = ws.Cells(i + 1, 3).Value
                                    
            Else
                'Add Volume_Total
                Volume_Total = Volume_Total + ws.Cells(i, 7).Value
                
                If ws.Cells(i, 3).Value = 0 And ws.Cells(i + 1, 3).Value <> 0 Then
                   Opening = ws.Cells(i + 1, 3).Value
                End If
            End If
            
        
            
        'Move to the next row
        Next i
                                
    'Move to the next worsksheet
    Next ws



End Sub

