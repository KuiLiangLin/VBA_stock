Attribute VB_Name = "Module3"
Sub d_divation_single_Q()
       
Dim i%, r%, j%, k%, m%, x%, y%, w%
Dim arr1, arr2

For x = 1001 To 1948 Step 1

    If Sheets("collect_Q").Cells(x, 1) = "公司" Then
        'For k = 1 To 70 Step 1
        '    Sheets("collect_Q").Cells(x + 1000, k) = Sheets("collect_Q").Cells(x, k)
        'Next
        x = x + 1
    End If
    
    If Sheets("collect_Q").Cells(x, 1) = "代號" Then
        'For j = 1 To 70 Step 1
        '    Sheets("collect_Q").Cells(x + 1000, j) = Sheets("collect_Q").Cells(x, j)
        'Next
        x = x + 1
    End If
    
    If Sheets("collect_Q").Cells(x, 1) = "6541" Then
        x = x + 1
    End If
        
        
        
        
    Sheets("collect_Q").Cells(x + 1000, 1) = Sheets("collect_Q").Cells(x, 1)
    Sheets("collect_Q").Cells(x + 1000, 2) = Sheets("collect_Q").Cells(x, 2)
    
    For y = 3 To 50 Step 8
        
         Sheets("collect_Q").Cells(x + 1000, y) = Sheets("collect_Q").Cells(x, y)
         Sheets("collect_Q").Cells(x + 1000, y + 1) = Sheets("collect_Q").Cells(x, y + 1)
         
         Sheets("collect_Q").Cells(x + 1000, y + 2) = Sheets("collect_Q").Cells(x, y + 2) - Sheets("collect_Q").Cells(x, y)
         Sheets("collect_Q").Cells(x + 1000, y + 3) = Sheets("collect_Q").Cells(x, y + 3) - Sheets("collect_Q").Cells(x, y + 1)
         
         Sheets("collect_Q").Cells(x + 1000, y + 4) = Sheets("collect_Q").Cells(x, y + 4) - Sheets("collect_Q").Cells(x, y + 2)
         Sheets("collect_Q").Cells(x + 1000, y + 5) = Sheets("collect_Q").Cells(x, y + 5) - Sheets("collect_Q").Cells(x, y + 3)
         
         Sheets("collect_Q").Cells(x + 1000, y + 6) = Sheets("collect_Q").Cells(x, y + 6) - Sheets("collect_Q").Cells(x, y + 4)
         Sheets("collect_Q").Cells(x + 1000, y + 7) = Sheets("collect_Q").Cells(x, y + 7) - Sheets("collect_Q").Cells(x, y + 5)
    Next
 
Next

End Sub



