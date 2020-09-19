Attribute VB_Name = "Module4"
Sub loop_worksheets():

Dim sheet As Worksheet
Application.ScreenUpdating = False

For Each sheet In Worksheets
    sheet.Select
    Call VBA_assignment

Next
Application.ScreenUpdating = True

End Sub

Sub VBA_assignment():

Dim ticker_type As String
Dim output_index As Integer
Dim yearly_change As Double
Dim new_row As Boolean
Dim open_value As Double
Dim closing_value As Double
Dim percentage_change As Double
Dim total_stock_volume As Double
Dim last_row As Double

Dim last_row2 As Double
Dim p As Integer
Dim max As Double
Dim max_ticker As String
Dim min As Double
Dim min_ticker As String
Dim max_stock_volume As Double
Dim max_stock_volume_ticker As String

output_index = 2
yearly_change = 0
new_row = True
open_value = 0
percentage_change = 0
total_stock_volume = 0
last_row = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To last_row
        
        If new_row = True Then
            total_stock_volume = Cells(i + 1, 7).Value
            open_value = Cells(i, 3)
            new_row = False
        
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1) Then
            ticker_type = Cells(i, 1).Value
            closing_value = Cells(i, 6)
            yearly_change = (closing_value - open_value)
            total_stock_volume = total_stock_volume + Cells(i, 7)
                
                If open_value = 0 Then
                    percentage_change = 0
                
                Else: percentage_change = yearly_change / open_value
                

                End If
            
            Cells(output_index, 9).Value = ticker_type
            Cells(output_index, 10).Value = yearly_change
            Cells(output_index, 11).Value = percentage_change
            Cells(output_index, 11).NumberFormat = "0.00%"
            
            If yearly_change >= 0 Then
                Cells(output_index, 10).Interior.ColorIndex = 4
            
            Else
                Cells(output_index, 10).Interior.ColorIndex = 3
            
            End If
            
            new_row = True
            yearly_change = 0
            total_stock_volume = 0
            output_index = output_index + 1
            
        Else
            total_stock_volume = total_stock_volume + Cells(i + 1, 7).Value
            Cells(output_index, 12).Value = total_stock_volume
        
        End If
    
    Next i
    
    last_row2 = Cells(Rows.Count, 9).End(xlUp).Row
    max = 0
    min = 0
    max_stock_volume = 0

        For p = 2 To last_row2
                    
            If Cells(p, 11).Value > max Then
            max = Cells(p, 11).Value
            max_ticker_name = Cells(p, 9).Value
            Range("P2").Value = max
            Range("P2").NumberFormat = "0.00%"
            Range("O2").Value = max_ticker_name
            Else
                    
            End If
                    
            If Cells(p, 11).Value < min Then
            min = Cells(p, 11).Value
            min_ticker_name = Cells(p, 9).Value
            Range("P3").Value = min
            Range("P3").NumberFormat = "0.00%"
            Range("O3").Value = min_ticker_name
                    
            Else
                    
            End If
                
            If Cells(p, 12).Value > max_stock_volume Then
            max_stock_volume = Cells(p, 12).Value
            max_stock_volume_ticker = Cells(p, 9).Value
            Range("P4").Value = max_stock_volume
            Range("O4").Value = max_stock_volume_ticker
                    
            Else
                    
            End If


        Next p

End Sub

