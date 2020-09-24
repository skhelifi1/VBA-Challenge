Attribute VB_Name = "Module2014"
Sub Multipleyear_stock()

        Dim currentticker As String
        Dim opendate As Long
        Dim closedate As Long
        Dim openvalue As Double
        Dim closedvalue As Double
        
        
        Dim yearlychange As Double
        Dim percentchange As Double
        Dim totalvolume As Double
        Dim summarytable_Row As Integer
        
        lastrow = Cells(Rows.Count, 1).End(xlDown).Row
        summarytable_Row = 2
        opendate = Cells(2, 2).Value
        closedate = Cells(262, 2).Value
        totalvolume = 0
        
    For i = 2 To lastrow
    
        currentticker = Cells(i, 1).Value
        totalvolume = totalvolume + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> currentticker Then
        
        
        Range("I" & summarytable_Row).Value = currentticker
        Range("J" & summarytable_Row).Value = yearlychange
        Range("K" & summarytable_Row).Value = percentchange
        Range("L" & summarytable_Row).Value = totalvolume
        
        summarytable_Row = summarytable_Row + 1
        
        totalvolume = 0
        End If
        
        If Cells(i, 2).Value = opendate Then
        openvalue = Cells(i, 3).Value
        End If
        
        If Cells(i + 1, 2).Value = closedate Then
        closedvalue = Cells(i + 1, 6).Value
        End If
        
        yearlychange = closedvalue - openvalue
                
        percentchange = yearlychange / openvalue
            
        If yearlychange > 0 Then
        Range("J" & summarytable_Row).Interior.ColorIndex = 4
        End If
        If yearlychange < 0 Then
        Range("J" & summarytable_Row).Interior.ColorIndex = 3

        End If
        
          
  Next i
         

End Sub




