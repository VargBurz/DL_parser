Attribute VB_Name = "Module1"
Sub DL_parser()

Dim TN As String
Dim LIST(50) As Long

Dim c As Integer

c = ActiveCell.Row

For k = c To c + 100

    TN = ActiveSheet.Cells(k, 16)
    If (Len(TN) < 1 And Len(ActiveSheet.Cells(k, 1)) < 1) Then
        MsgBox ("Empty string. End of parsing!")
        Exit For
    End If
    If (Len(TN) < 1) Then
        MsgBox ("empty bill")
        GoTo NextIteration
    End If
    
    TN = Replace(TN, "-", "")
    If (Len(TN) > 13) Then
        TN = Split(TN, ",")
        bill_number = UBound(TN) + 1
    Else
        bill_number = 1
    End If
    
    With CreateObject("Microsoft.XMLHTTP")
        .Open "GET", "https://www.dellin.ru/cabinet/orders/" & TN, "False"
        .send
        If .statustext = "OK" Then
            s = .responsetext 'это и есть HTML-код страницы
            'ActiveSheet.Cells(36, 1) = Len(s)
            's = Mid(s, 38000, 35000)
            l = Len(s) - 19
            n = 1
            For i = 1 To l
                If (Mid(s, i, 19) Like "doc-transfer__price") Then
                    LIST(n) = i
                    n = n + 1
                End If
            Next
            
            For i = 1 To bill_number
                x = LIST(i)
                y = LIST(i + 1)
            d = y - x
            watch_s = Mid(s, x, d)
            'MsgBox (Len(watch_s))
            Ln = Len(watch_s) - 5
            point1 = 1
            point2 = 2
            
            For j = 1 To Ln
                If (Mid(watch_s, j, 5) Like "<span") Then
                    'MsgBox (Mid(watch_s, j, 5))
                    point1 = j + 5
                    'MsgBox (point1)
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
            'MsgBox (Mid(watch_s, point1, point2))
            
            point3 = 1
            For i = 1 To Len(watch_s) - 7
                If (Mid(watch_s, i, 7) Like "</span>") Then
                    point3 = i
                    Exit For
                End If
            Next
            'MsgBox (Mid(watch_s, point3, 7))
            result = Mid(watch_s, point1 + point2, point3 - point2 - point1)
            result = Replace(result, " ", "")
            result = Replace(result, vbLf, "")
            'MsgBox (result)
            'MsgBox (Len(result))
            result = CInt(result)
            'ActiveSheet.Cells(1, 18) = result
            If ActiveSheet.Cells(k, 17).Value = result Then
                MsgBox ("ok!")
                ActiveSheet.Cells(k, 17).Interior.Color = vbGreen
            Else
                MsgBox ("ne ok!")
                ActiveSheet.Cells(k, 17).Interior.Color = vbRed
                
            End If
        Else
            MsgBox ("WRONG http, check TN")
            ActiveSheet.Cells(k, 16).Interior.Color = vbRed
            'Exit For
        End If
    
    End With
NextIteration:
Next

End Sub
