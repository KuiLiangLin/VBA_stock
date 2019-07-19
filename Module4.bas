Attribute VB_Name = "Module4"
Sub recent_4_month()
Attribute recent_4_month.VB_ProcData.VB_Invoke_Func = " \n14"


        
Dim i%, r%, x%, y%, w%
Dim arr1, arr2

w = 0

For y = 3 To 6

    arr1 = Sheets(y).Range("A12:K12") '1计沮结倒计舱arr1
    arr2 = Sheets(y).Range("A1:A2000") '2计沮结倒计舱arr2
    
    For x = 3 To 947
    
        For i = 1 To 30 'Θ2︽计
            If arr1(1, i) = "るゑ耕" Then
            Exit For
            End If
        Next
        
        
        For r = 1 To UBound(arr2) 'Θ2︽计
            If arr2(r, 1) = Sheets("overview").Range("A" & x).Value Then
            Exit For
            End If
        Next
        
        
        
        Sheets("overview").Cells(1, y + w) = Sheets(y).Cells(12, i)
        Sheets("overview").Cells(1, y + w + 1) = Sheets(y).Cells(12, i + 1)
        
        'Sheets("overview").Range("C" & x) = Sheets(y).Cells(r, i)
        'Sheets("overview").Range("D" & x) = Sheets(y).Cells(r, i + 1)
        
        Sheets("overview").Cells(x, y + w) = Sheets(y).Cells(r, i)
        Sheets("overview").Cells(x, y + w + 1) = Sheets(y).Cells(r, i + 1)
        
    Next
    w = w + 1

Next

End Sub
