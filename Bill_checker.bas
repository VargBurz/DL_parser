Attribute VB_Name = "DL_parser"
Sub DL_parser()

Dim TN As String
Dim S_BILL As String
Dim BILL_COUNT As Integer
Dim LIST(50) As Long
Dim SUMM_BILLS() As String
Dim c As Integer
Dim RESULT_LIST(50) As Integer
Dim R_BILL As Integer
Dim new_tn As String
Dim tn_list() As String
Dim DL_BILL As Integer
Dim DL_COUNT As Integer

c = ActiveCell.Row

For k = c To c + 300

    TN = ActiveSheet.Cells(k, 16)
    Cells(k, 16).Select
    If (Len(TN) < 1 And Len(ActiveSheet.Cells(k, 1)) < 1) Then
        MsgBox ("Empty string. End of parsing!")
        Exit For
    End If
    
    If (ActiveSheet.Cells(k, 9) <> "Деловые линии") Then
        'MsgBox ("ne DL")
        GoTo NextIteration
    End If
    
    If (Len(TN) < 1) Then
        'MsgBox ("empty bill")
        GoTo NextIteration
    End If
    
    TN = Replace(TN, "-", "")
    TN = Replace(TN, " ", "")
    If (Len(TN) > 13) Then
        tn_list = Split(TN, ",")
        bill_number = UBound(tn_list) + 1
        
    Else
        bill_number = 1
    End If
    
    With CreateObject("Microsoft.XMLHTTP")
        
        If (Len(TN) > 13) Then
            'MsgBox (TN)
            new_tn = tn_list(1)
        Else
            new_tn = TN
        End If
        .Open "GET", "https://www.dellin.ru/cabinet/orders/" & new_tn, "False"
        .send
        If .statustext = "OK" Then
            S = .responsetext
            l = Len(S) - 19
            n = 1
            For i = 1 To l
                If (Mid(S, i, 19) Like "doc-transfer__price") Then
                    LIST(n) = i
                    n = n + 1
                End If
            Next
            dl_bill_number = (n - 1) / 2
            If bill_number <> dl_bill_number Then
                ActiveSheet.Cells(k, 16).Interior.Color = vbYellow
                MsgBox ("Diffrent number of bills!")
            End If
			
	Erase RESULT_LIST
            For i = 1 To dl_bill_number
                x = LIST(i * 2 - 1)
                y = LIST(i * 2)
                d = y - x
                watch_s = Mid(S, x, d)
                Ln = Len(watch_s) - 5
                point1 = 1
                point2 = 2
            
                For j = 1 To Ln
                    If (Mid(watch_s, j, 5) Like "<span") Then
                        point1 = j + 5
                        Exit For
                    End If
                Next
            
                xg = Len(watch_s) - point1
                shr_ws = Mid(watch_s, point1, xg)
                For j = 1 To Len(shr_ws):
                    If (Mid(shr_ws, j, 1) Like ">") Then
                        point2 = j
                        Exit For
                    End If
                Next
            
                point3 = 1
                For ind = 1 To Len(watch_s) - 7
                    If (Mid(watch_s, ind, 7) Like "</span>") Then
                        point3 = ind
                        Exit For
                    End If
                Next
                result = Mid(watch_s, point1 + point2, point3 - point2 - point1)
                result = Replace(result, " ", "")
                result = Replace(result, vbLf, "")
                result = CInt(result)
                
                'MsgBox (result)
                
                RESULT_LIST(i) = result
            Next
            
            DL_BILL = 0
            DL_COUNT = UBound(RESULT_LIST)
            For dl_b = 1 To DL_COUNT
		If RESULT_LIST(dl_b) = 0 Then
                    k_end = k_end + 1
                    If k_end > 1 Then
                        Exit For
                    End If
                End If
                DL_BILL = DL_BILL + RESULT_LIST(dl_b)
            Next
                        
            S_BILL = ActiveSheet.Cells(k, 17)
            S_BILL = Replace(S_BILL, "=", "")
            S_BILL = Replace(S_BILL, " ", "")
            R_BILL = 0
            If InStr(S_BILL, "+") > 0 Then
                SUMM_BILLS = Split(S_BILL, "+")
                BILL_COUNT = UBound(SUMM_BILLS) + 1
                For Count = 1 To BILL_COUNT
                    R_BILL = R_BILL + CInt(SUMM_BILLS(i))
                Next
            Else
                Dim rand_list(1) As String
                rand_list(1) = S_BILL
                SUMM_BILLS = rand_list
                R_BILL = CInt(S_BILL)
            End If
            
            'MsgBox (R_BILL)
            'MsgBox (DL_BILL)
            
            If R_BILL = DL_BILL Then
                MsgBox ("The amount of bills coincides!")
                ActiveSheet.Cells(k, 17).Interior.Color = vbWhite
            Else
                MsgBox ("The amount of bills doesn't coincide!")
                ActiveSheet.Cells(k, 17).Interior.Color = vbRed
                
            End If
        Else
            MsgBox ("Account Access Error. Check the number of TN!")
            ActiveSheet.Cells(k, 16).Interior.Color = vbRed
        End If
    
    End With
NextIteration:
Next

End Sub
