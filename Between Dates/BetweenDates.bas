Attribute VB_Name = "BETWEENDATES"
Function BETWEENDATES(date1, Optional RoundToDays = False, Optional date2 = 0)

'   Displays the interval between two dates in text form.
'   If only one date value was provided, we calculate the interval from Current Time.
'   By default, both date and time of day are taken into consideration.
'   The time of day will be ignored when parameter RoundToDays is set to True.

Dim Interval As Double

If date2 = 0 Then
    date2 = Now
End If

If Not (IsDate(date1) And IsDate(date2)) Then
    BETWEENDATES = CVErr(xlErrValue)
    Exit Function
End If

If RoundToDays = True Then
    date1 = Int(date1)
    date2 = Int(date2)
End If

Interval = date2 - date1

'   If date1 is in the past relative to date2:
If Interval >= 0 Then

    If Interval < 1 Then
        
        If RoundToDays = True Then
                BETWEENDATES = "Today."
                Exit Function

        '   interval over 2 hours:
        ElseIf (Interval * 24) >= 2 Then
                BETWEENDATES = Round(Interval * 24) & " hours ago."
                Exit Function
                
        '   interval over 2 minutes:
        ElseIf (Interval * 24 * 60) > 2 Then
                BETWEENDATES = Round(Interval * 24 * 60) & " minutes ago."
                Exit Function
                
        '   interval under 2 minutes:
        Else
                BETWEENDATES = "Just now."
                Exit Function
        End If
        
    '   interval under ~26 hours:
    ElseIf Interval < 1.1 Then
        BETWEENDATES = "Yesterday."
        Exit Function
        
    '   interval under 48 hours:
    ElseIf Interval < 2 Then
        BETWEENDATES = "A day and " & (Round(Interval * 24)) - 24 & " hours ago."
        Exit Function
        
    '   interval over 48 hours:
    Else
        BETWEENDATES = (Round(Interval, 0)) & " days ago."
        Exit Function
        
    End If
    
'   If date1 is in the future relative to date2:
ElseIf Interval < 0 Then

    If Interval > -1 Then
          
        If RoundToDays = True Then
                BETWEENDATES = "Today."
                Exit Function
          
        '   interval over 2 hours:
        ElseIf (Interval * 24) <= -2 Then
                BETWEENDATES = Abs(Round(Interval * 24)) & " hours from now."
                Exit Function
                
        '   interval over 2 minute:
        ElseIf (Interval * 24 * 60) < -2 Then
                BETWEENDATES = Abs(Round(Interval * 24 * 60)) & " minutes from now."
                Exit Function
                
        '   interval under 2 minutes:
        Else
                BETWEENDATES = "Just now."
                Exit Function
                
        End If
    
    '   interval under ~26 hours:
    ElseIf Interval > -1.1 Then
        BETWEENDATES = "Tomorrow."
        Exit Function
        
    '   interval under 48 hours:
     ElseIf Interval > -2 Then
        BETWEENDATES = "A day and " & Abs(Round(Interval * 24)) - 24 & " hours from now."
        Exit Function
        
    '   interval over 48 hours:
    Else
        BETWEENDATES = Abs(Round(Interval, 0)) & " days from now."
        Exit Function
    End If
    
End If

End Function
