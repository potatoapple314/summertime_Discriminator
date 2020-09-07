Option Explicit
Public Function summer(cell As String) As Boolean
    Dim ref As Date
    Dim toDate As String
    Dim toDate1 As String
    Dim toDate2 As String
    Dim AddDate As Date
    Dim I As Integer
    Dim summer_start  As Date
    Dim summer_end  As Date
    

    toDate = Format(cell, "yyyy/mm/dd")
    ref = CDate(toDate)
    'MsgBox ref
    toDate1 = Format(cell, "yyyy/3/7")
    'MsgBox toDate1
    For I = 1 To 7
        AddDate = DateAdd("d", I, CDate(toDate1))
        If Weekday(AddDate, 2) = 7 Then
            'MsgBox (AddDate)
            summer_start = AddDate
            Exit For
        End If
    Next I
    
    toDate2 = Format(cell, "yyyy/11/1")
    'MsgBox toDate2
    For I = 1 To 7
        AddDate = DateAdd("d", I, CDate(toDate2))
        If Weekday(AddDate, 2) = 7 Then
            'MsgBox (AddDate)
            summer_end = AddDate
            Exit For
        End If
    Next I
    
    If ref >= summer_start And ref <= summer_end Then
        summer = True
        Exit Function
    Else
        summer = False
        Exit Function
    End If

End Function

