Attribute VB_Name = "Module2"
Sub WallStreet()

         'variables to loop worksheet, rows
         Dim Current As Worksheet
         Dim l_lastrow As Long
         Dim l_displayrow    As Integer
         
         'Variables to calculate stock values per ticker
         Dim v_totalvolume   As Double
         Dim v_ticker        As String
         Dim v_openprice     As Double
         Dim v_variance      As Double
         Dim v_percentage    As Double

         'Variables for min,max values
         Dim v_max_ticker  As String
         Dim v_max         As Double
         Dim v_min_ticker  As String
         Dim v_min         As Double
         Dim v_vol_ticker  As String
         Dim v_vol         As Double
         
        

         For Each Current In Worksheets

            l_lastrow = Current.Cells(Rows.Count, 1).End(xlUp).Row
            l_displayrow = 2
            
            'Headings
            Current.Range("I1").Value = "Ticker"
            Current.Range("J1").Value = "Yearly Change"
            Current.Range("K1").Value = "Percent Change"
            Current.Range("L1").Value = "Total Stock Volume"
            Current.Range("O2").Value = "Greatest % Increase"
            Current.Range("O3").Value = "Greatest % Decrease"
            Current.Range("O4").Value = "Greatest Total Volume"
            Current.Range("P1").Value = "Ticker"
            Current.Range("Q1").Value = "Value"
            
            
            
            For I = 2 To l_lastrow
    
                   If Current.Cells(I + 1, 1).Value <> Current.Cells(I, 1).Value Then
                      
                      v_ticker = Current.Cells(I, 1).Value
                      v_totalvolume = v_totalvolume + Current.Cells(I, 7)
                      v_variance = Current.Cells(I, 6).Value - v_openprice
                    
                      Current.Range("I" & l_displayrow).Value = v_ticker
                      Current.Range("L" & l_displayrow).Value = v_totalvolume
                      Current.Range("J" & l_displayrow).Value = v_variance

                      If v_openprice = 0 Or IsEmpty(v_openprice) = True Then
                         v_percentage = 0
                      ElseIf v_openprice <> 0 Then
                         v_percentage = v_variance / v_openprice
                      End If
                      
                      Current.Range("K" & l_displayrow).NumberFormat = "0.00%"
                      Current.Range("K" & l_displayrow).Value = v_percentage


                      If v_variance < 0 Then
                         Current.Range("J" & l_displayrow).Interior.ColorIndex = 3
                      ElseIf v_variance > 0 Then
                         Current.Range("J" & l_displayrow).Interior.ColorIndex = 4
                      End If


                      
                        'Calculating greates volume
                        If IsEmpty(v_vol) = True Then
                           v_vol_ticker = v_ticker
                           v_vol = v_totalvolume
                        
                        Else
                            If v_totalvolume > v_vol Then
                                v_vol_ticker = v_ticker
                                v_vol = v_totalvolume
                            End If
                        End If
                    
                      
                      'Calculating greatest increase/decrease
                      If IsEmpty(v_percentage) = True Then
                           v_max_ticker = v_ticker
                           v_max = v_percentage
                          
                           v_min_ticker = v_ticker
                           v_min = v_percentage
                         
                        
                        Else
                        
                            If v_percentage > v_max Then
                                v_max_ticker = v_ticker
                                v_max = v_percentage
                             End If
                             
                        
                           If v_percentage < v_min Then
                                v_min_ticker = v_ticker
                                v_min = v_percentage
                           End If
                        End If
                      
                          
                      
                      'Reseting values for ticker
                      l_displayrow = l_displayrow + 1
                      v_totalvolume = 0
                      v_openprice = Empty
                      v_variance = Empty
                      v_ticker = Empty
                     
                
                  Else
            
                        v_totalvolume = (v_totalvolume) + Current.Cells(I, 7)
                                        
                    
                        If v_openprice = Empty Then
                           v_openprice = Current.Cells(I, 3)
        
                        End If
                           
                    
                  End If
  
            Next I


            'Setting greates values for the worksheet
            Current.Range("P4").Value = v_vol_ticker
            Current.Range("Q4").Value = v_vol
            
            Current.Range("P2").Value = v_max_ticker
            Current.Range("Q2").Value = v_max
            Current.Range("Q2").NumberFormat = "0.00%"
            
            
            Current.Range("P3").Value = v_min_ticker
            Current.Range("Q3").Value = v_min
            Current.Range("Q3").NumberFormat = "0.00%"
            
            Current.Columns(15).AutoFit
            Current.Columns(16).AutoFit
            Current.Columns(17).AutoFit
            Current.Columns(9).AutoFit
            Current.Columns(10).AutoFit
            Current.Columns(11).AutoFit
            Current.Columns(12).AutoFit
            
            
            'REsetting worksheet variables
            v_vol_ticker = Empty
            v_vol = Empty
            v_max_ticker = Empty
            v_max = Empty
            v_min_ticker = Empty
            v_min = Empty
            
         Next
         
         
         
End Sub

