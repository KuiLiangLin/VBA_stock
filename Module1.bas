Attribute VB_Name = "Module1"
Sub history_quart()
       
Dim i%, r%, j%, k%, m%, x%, y%, w%
Dim arr1, arr2

''''''''''''''''''''''''''''''''''''''''''''''''''''''
w = 3     'stored data always starts at collect_Q collumn C and D
''''''''''''''''''''''''''''''''''''''''''''
For y = 9 To 10 Step 1 '34 ' sheet eps10802 TO sheet eps10801
''''''''''''''''''''''''''''''''''''''''''''


    arr2 = Sheets(y).Range("A1:A1000")
    For x = 3 To 948 Step 1

        For r = 1 To 999 'UBound(arr2)
        If arr2(r, 1) = Sheets("collect_Q").Range("A" & x).Value Then
            Exit For
            End If
        Next
        
        For i = r To 2 Step -1
            If arr2(i, 1) = "公司" Then
            Exit For
            End If
        Next
        
        arr1 = Sheets(y).Range(Sheets(y).Cells(i, 1), Sheets(y).Cells(i, 35))
            
        For j = 1 To 35 Step 1
            If arr1(1, j) Like "*利息淨收益*" Then
            Exit For
            End If
            If arr1(1, j) Like "*收益*" Then
            Exit For
            End If
            If arr1(1, j) Like "*營業收入*" Then
           Exit For
            End If
        Next
        
        For k = 1 To 35 Step 1
            If arr1(1, k) Like "*基本每股盈餘*" = True Then
            Exit For
            End If
        Next
        
        For m = x To 2 Step -1
            If Sheets("collect_Q").Range("A" & m).Value = "公司" Then
            Exit For
            End If
        Next
        
        
        Sheets("collect_Q").Cells(x, w) = Sheets(y).Cells(r, j)
        Sheets("collect_Q").Cells(x, w + 1) = Sheets(y).Cells(r, k)
        
        Sheets("collect_Q").Cells(m, w) = "rev" & Mid(Sheets(y).Name, 4, 8) 'Sheets(y).Cells(i, j)
        Sheets("collect_Q").Cells(m, w + 1) = "eps" & Mid(Sheets(y).Name, 4, 8) 'Sheets(y).Cells(i, k)
        
    Next
    w = w + 2

Next

End Sub

