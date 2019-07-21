Attribute VB_Name = "Module3"
Sub d_divation_single_Q()
       
Dim i%, r%, j%, k%, m%, n%, x%, y%, z%, w%
Dim arr1(1000, 100), arr2(1000, 100), arr3(1000, 8)

For x = 1001 To 1948 Step 1 'xxxxxxxxxxxx

    If Sheets("collect_Q").Cells(x, 1) = "公司" Then
        Sheets("collect_Q").Cells(x + 1000, 1) = Sheets("collect_Q").Cells(x, 1)
        Sheets("collect_Q").Cells(x + 1000, 2) = Sheets("collect_Q").Cells(x, 2)
        Sheets("collect_Q").Cells(x + 1000, 3) = "AVG_Q1"
        Sheets("collect_Q").Cells(x + 1000, 4) = "DEV_Q1"
        Sheets("collect_Q").Cells(x + 1000, 5) = "AVG_Q2"
        Sheets("collect_Q").Cells(x + 1000, 6) = "DEV_Q2"
        Sheets("collect_Q").Cells(x + 1000, 7) = "AVG_Q3"
        Sheets("collect_Q").Cells(x + 1000, 8) = "DEV_Q3"
        Sheets("collect_Q").Cells(x + 1000, 9) = "AVG_Q4"
        Sheets("collect_Q").Cells(x + 1000, 10) = "DEV_Q4"
        x = x + 1
    End If
    
    If Sheets("collect_Q").Cells(x, 1) = "代號" Then
        Sheets("collect_Q").Cells(x + 1000, 2) = Sheets("collect_Q").Cells(x, 2)
        x = x + 1
    End If
           
        
        
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Q1
    w = 0 'wwwwwwwwwwwwww
    For y = 1 To 70 Step 1 'yyyyyyyyyyyyyyy
        If Sheets("collect_Q").Cells(1001, y) Like "rev*" Then
            If Sheets("collect_Q").Cells(1001, y) Like "*01" Then
                For i = 8 To 70 Step 8 'iiiiiiiiiiiiiiiiiiii
                    If Sheets("collect_Q").Cells(x, y + 1) = 0 = False Then
                        If Sheets("collect_Q").Cells(x, y + i) = 0 = False Then
                            arr1(x - 1001, w) = (Sheets("collect_Q").Cells(x, y + i + 1) / Sheets("collect_Q").Cells(x, y + 1)) * (Sheets("collect_Q").Cells(x, y) / Sheets("collect_Q").Cells(x, y + i))
                            w = w + 1
                        End If
                    End If
                Next
            End If
        End If
        If Sheets("collect_Q").Cells(1001, y) Like "*END*" Then
            arr1(x - 1001, w) = "END1000"
        End If
    Next
    
    arr1_tmp = 0
    For z = 0 To 100 Step 1 'zzzzzzzzzzzzzzzzzzz
        If arr1(x - 1001, z) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr1(x - 1001, z)
    Next
    
    If z = 0 Then
        z = 1
    End If
    arr1_avg_q1 = arr1_tmp / z
    
    For m = 0 To 100 Step 1 'mmmmmmmmmmmmmmmmmmmmm
        If arr1(x - 1001, m) Like "*END*" Then
            Exit For
        End If
            arr2(x - 1001, m) = (arr1(x - 1001, m) - arr1_avg_q1) ^ 2
    Next
    
    arr1_tmp = 0
    For n = 0 To 100 Step 1 'nnnnnnnnnnnnnnnnnnnnnnn
        If arr2(x - 1001, n) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr2(x - 1001, n)
    Next
    arr1_dev_q1 = (arr1_tmp / z) ^ (1 / 2)
    
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Q2
    w = 0 'wwwwwwwwwwwwww
    For y = 1 To 70 Step 1 'yyyyyyyyyyyyyyy
        If Sheets("collect_Q").Cells(1001, y) Like "rev*" Then
            If Sheets("collect_Q").Cells(1001, y) Like "*02" Then
                For i = 8 To 70 Step 8 'iiiiiiiiiiiiiiiiiiii
                    If Sheets("collect_Q").Cells(x, y + 1) = 0 = False Then
                        If Sheets("collect_Q").Cells(x, y + i) = 0 = False Then
                            arr1(x - 1001, w) = (Sheets("collect_Q").Cells(x, y + i + 1) / Sheets("collect_Q").Cells(x, y + 1)) * (Sheets("collect_Q").Cells(x, y) / Sheets("collect_Q").Cells(x, y + i))
                            w = w + 1
                        End If
                    End If
                Next
            End If
        End If
        If Sheets("collect_Q").Cells(1001, y) Like "*END*" Then
            arr1(x - 1001, w) = "END1000"
        End If
    Next
    
    arr1_tmp = 0
    For z = 0 To 100 Step 1 'zzzzzzzzzzzzzzzzzzz
        If arr1(x - 1001, z) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr1(x - 1001, z)
    Next
    
    If z = 0 Then
        z = 1
    End If
    arr1_avg_q2 = arr1_tmp / z
    
    For m = 0 To 100 Step 1 'mmmmmmmmmmmmmmmmmmmmm
        If arr1(x - 1001, m) Like "*END*" Then
            Exit For
        End If
            arr2(x - 1001, m) = (arr1(x - 1001, m) - arr1_avg_q2) ^ 2
    Next
    
    arr1_tmp = 0
    For n = 0 To 100 Step 1 'nnnnnnnnnnnnnnnnnnnnnnn
        If arr2(x - 1001, n) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr2(x - 1001, n)
    Next
    arr1_dev_q2 = (arr1_tmp / z) ^ (1 / 2)
    
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Q3
    w = 0 'wwwwwwwwwwwwww
    For y = 1 To 70 Step 1 'yyyyyyyyyyyyyyy
        If Sheets("collect_Q").Cells(1001, y) Like "rev*" Then
            If Sheets("collect_Q").Cells(1001, y) Like "*03" Then
                For i = 8 To 70 Step 8 'iiiiiiiiiiiiiiiiiiii
                    If Sheets("collect_Q").Cells(x, y + 1) = 0 = False Then
                        If Sheets("collect_Q").Cells(x, y + i) = 0 = False Then
                            arr1(x - 1001, w) = (Sheets("collect_Q").Cells(x, y + i + 1) / Sheets("collect_Q").Cells(x, y + 1)) * (Sheets("collect_Q").Cells(x, y) / Sheets("collect_Q").Cells(x, y + i))
                            w = w + 1
                        End If
                    End If
                Next
            End If
        End If
        If Sheets("collect_Q").Cells(1001, y) Like "*END*" Then
            arr1(x - 1001, w) = "END1000"
        End If
    Next
    
    arr1_tmp = 0
    For z = 0 To 100 Step 1 'zzzzzzzzzzzzzzzzzzz
        If arr1(x - 1001, z) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr1(x - 1001, z)
    Next
    
    If z = 0 Then
        z = 1
    End If
    arr1_avg_q3 = arr1_tmp / z
    
    For m = 0 To 100 Step 1 'mmmmmmmmmmmmmmmmmmmmm
        If arr1(x - 1001, m) Like "*END*" Then
            Exit For
        End If
            arr2(x - 1001, m) = (arr1(x - 1001, m) - arr1_avg_q3) ^ 2
    Next
    
    arr1_tmp = 0
    For n = 0 To 100 Step 1 'nnnnnnnnnnnnnnnnnnnnnnn
        If arr2(x - 1001, n) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr2(x - 1001, n)
    Next
    arr1_dev_q3 = (arr1_tmp / z) ^ (1 / 2)
    
    
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Q4
    w = 0 'wwwwwwwwwwwwww
    For y = 1 To 70 Step 1 'yyyyyyyyyyyyyyy
        If Sheets("collect_Q").Cells(1001, y) Like "rev*" Then
            If Sheets("collect_Q").Cells(1001, y) Like "*04" Then
                For i = 8 To 70 Step 8 'iiiiiiiiiiiiiiiiiiii
                    If Sheets("collect_Q").Cells(x, y + 1) = 0 = False Then
                        If Sheets("collect_Q").Cells(x, y + i) = 0 = False Then
                            arr1(x - 1001, w) = (Sheets("collect_Q").Cells(x, y + i + 1) / Sheets("collect_Q").Cells(x, y + 1)) * (Sheets("collect_Q").Cells(x, y) / Sheets("collect_Q").Cells(x, y + i))
                            w = w + 1
                        End If
                    End If
                Next
            End If
        End If
        If Sheets("collect_Q").Cells(1001, y) Like "*END*" Then
            arr1(x - 1001, w) = "END1000"
        End If
    Next
    
    arr1_tmp = 0
    For z = 0 To 100 Step 1 'zzzzzzzzzzzzzzzzzzz
        If arr1(x - 1001, z) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr1(x - 1001, z)
    Next
    
    If z = 0 Then
        z = 1
    End If
    arr1_avg_q4 = arr1_tmp / z
    
    For m = 0 To 100 Step 1 'mmmmmmmmmmmmmmmmmmmmm
        If arr1(x - 1001, m) Like "*END*" Then
            Exit For
        End If
            arr2(x - 1001, m) = (arr1(x - 1001, m) - arr1_avg_q4) ^ 2
    Next
    
    arr1_tmp = 0
    For n = 0 To 100 Step 1 'nnnnnnnnnnnnnnnnnnnnnnn
        If arr2(x - 1001, n) Like "*END*" Then
            Exit For
        End If
            arr1_tmp = arr1_tmp + arr2(x - 1001, n)
    Next
    arr1_dev_q4 = (arr1_tmp / z) ^ (1 / 2)
    

    
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''WRITE
    Sheets("collect_Q").Cells(x + 1000, 1) = Sheets("collect_Q").Cells(x, 1)
    Sheets("collect_Q").Cells(x + 1000, 2) = Sheets("collect_Q").Cells(x, 2)
    Sheets("collect_Q").Cells(x + 1000, 3) = arr1_avg_q1
    Sheets("collect_Q").Cells(x + 1000, 4) = arr1_dev_q1
    Sheets("collect_Q").Cells(x + 1000, 5) = arr1_avg_q2
    Sheets("collect_Q").Cells(x + 1000, 6) = arr1_dev_q2
    Sheets("collect_Q").Cells(x + 1000, 7) = arr1_avg_q3
    Sheets("collect_Q").Cells(x + 1000, 8) = arr1_dev_q3
    Sheets("collect_Q").Cells(x + 1000, 9) = arr1_avg_q4
    Sheets("collect_Q").Cells(x + 1000, 10) = arr1_dev_q4
 
Next

End Sub



