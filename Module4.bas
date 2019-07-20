Attribute VB_Name = "Module4"
Sub a_recent_4_month()
Attribute a_recent_4_month.VB_ProcData.VB_Invoke_Func = " \n14"
        
Dim i%, r%, x%, y%, w%
Dim arr1, arr2


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
w = 3    'stored data always starts at collect_M collumn C and D
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
For y = 8 To 4 Step -1   'sheet 4 TO sheet 7
'''''''''''''''''''''''''''''''''''''''''''''''''''''''


    arr1 = Sheets(y).Range("A12:K12")
    arr2 = Sheets(y).Range("A1:A1000")
    
    For x = 3 To 947
    
        For i = 1 To 30
            If arr1(1, i) = "上月比較" Then
            Exit For
            End If
        Next
        
        
        For r = 1 To UBound(arr2)
            If arr2(r, 1) = Sheets("collect_M").Range("A" & x).Value Then
            Exit For
            End If
        Next
        
        
        
        Sheets("collect_M").Cells(1, w) = Sheets(y).Name & Sheets(y).Cells(12, i)
        Sheets("collect_M").Cells(1, w + 1) = Sheets(y).Name & Sheets(y).Cells(12, i + 1)
        
        Sheets("collect_M").Cells(x, w) = Sheets(y).Cells(r, i)
        Sheets("collect_M").Cells(x, w + 1) = Sheets(y).Cells(r, i + 1)
        
    Next
    w = w + 2

Next

End Sub
