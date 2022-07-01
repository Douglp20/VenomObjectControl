Public Class DateTimeControl
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
#Region "Time and Date"

    Public Function DateIsNull(dte As System.Windows.Forms.DateTimePicker) As Object

        On Error GoTo Err
        If dte.Checked Then
            DateIsNull = dte.Value
        Else
            DateIsNull = DBNull.Value ''.ToString()
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function AddTimeToDateIsNull(dte As System.Windows.Forms.DateTimePicker, time As String) As Object
        On Error GoTo Err
        'time 11:00
        Dim StartDate As DateTime = dte.Value
        Dim strDateTime As String = CStr(Mid(StartDate, 1, 10))
        Dim strHour As String = CStr(Mid(time, 1, 2))
        Dim strMin As String = CStr(Mid(time, 4, 2))
        Dim lngHour As Integer = CInt(Mid(time, 1, 2))

        If dte.Checked Then
            If lngHour >= 12 Then
                strDateTime = "#" + strDateTime + " " + strHour + ":" + strMin + ":00 PM#"

            Else
                strDateTime = "#" + strDateTime + " " + strHour + ":" + strMin + ":00 AM#"
            End If

            AddTimeToDateIsNull = strDateTime

        Else
            AddTimeToDateIsNull = DBNull.Value

        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function


    Public Sub DatePickerValue(dte As System.Windows.Forms.DateTimePicker, DateValue As String)
        On Error GoTo Err
        If DateValue.ToString.Length > 0 Then
            dte.Value = DateValue
            If dte.ShowCheckBox Then dte.Checked = True
        Else
            If dte.ShowCheckBox Then dte.Checked = False
        End If

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub TimeInterval(StartTime As DateTime, Interval As Integer, cbo As System.Windows.Forms.ComboBox)
        On Error GoTo Err
        'time
        Dim EndTime As DateTime = "#23:30:00 PM#"


        cbo.Items.Clear()
        ' cbo.Items.Add("00:00")
        For i As Integer = 0 To 150
            cbo.Items.Add(StartTime.ToShortTimeString())
            StartTime = DateAdd(DateInterval.Minute, Interval, StartTime)
            If StartTime = EndTime Then Exit For
        Next


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub TimeIntervalWithBlank(StartTime As DateTime, Interval As Integer, cbo As System.Windows.Forms.ComboBox)
        On Error GoTo Err
        'time
        Dim EndTime As DateTime = "#23:30:00 PM#"


        cbo.Items.Clear()
        cbo.Items.Add("00:00")
        For i As Integer = 0 To 150
            cbo.Items.Add(StartTime.ToShortTimeString())
            StartTime = DateAdd(DateInterval.Minute, Interval, StartTime)
            If StartTime = EndTime Then Exit For
        Next


        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub

    Public Function GetTimeHours(StartTime As DateTime, EndTime As DateTime) As String
        On Error GoTo Err
        'time

        Dim TTF As New TimeSpan
        TTF = EndTime.Subtract(StartTime)
        GetTimeHours = TTF.Hours.ToString("00") + "." + TTF.Minutes.ToString("00")


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function GetTimeHours(StartTime As String, EndTime As String) As String
        On Error GoTo Err
        'time

        If StartTime.Length = 5 And EndTime.Length = 5 Then
            If StartTime.ToString().Replace(":", "").Trim = "0000" OrElse EndTime.ToString().Replace(":", "").Trim = "0000" = False Then
                Dim ParseStartTime As DateTime = Date.Parse(StartTime)
                Dim ParseEndTime As DateTime = Date.Parse(EndTime)

                Dim TTF As New TimeSpan
                TTF = ParseEndTime.Subtract(ParseStartTime)
                GetTimeHours = TTF.Hours.ToString("00") + "." + TTF.Minutes.ToString("00")
            Else
                GetTimeHours = "00:00"
            End If
        Else
            GetTimeHours = "00:00"
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function getTotalHours(Value As ArrayList) As String
        On Error GoTo Err
        Dim MinTotal As Integer = 0
        Dim FullTotalHours As String = "00.00"
        Dim Min As Integer = 0
        Dim Hour As Integer = 0
        Dim HourToMin As Integer = 0
        Dim arrValue As New ArrayList
        For i As Integer = 0 To Value.Count - 1
            If Len(Value(i)) = 5 Then
                arrValue = GetFormatSplitValue(Value(i), ".")
                HourToMin = Math.Floor(arrValue(0) * 60)
                Min = arrValue(1)
                MinTotal = MinTotal + (HourToMin + Min)
            End If
        Next
        If MinTotal > 20 Then
            'MinTotal = MinTotal Mod 1440
            Hour = MinTotal \ 60
            Min = MinTotal Mod 60
            Return String.Format("{0:00}.{1:00}", Hour, Min)
        Else
            Return String.Format("{0:00}.{1:00}", 0, 0)
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function TotalHours(Value As ArrayList) As String
        On Error GoTo Err
        'Dim MinTotal As Integer = 0
        'Dim FullTotalHours As String = "00.00"
        Dim Minute As Int32 = 0
        Dim Hour As Int32 = 0
        Dim MinuteTotal As Int32 = 0
        Dim arrValue As New ArrayList
        Dim newValue As String = String.Empty

        For i As Integer = 0 To Value.Count - 1
            If Len(Value(i)) = 5 Then
                newValue = Value(i).ToString().Replace(".", "").Trim
                If Int32.TryParse(newValue.Substring(0, 2), Hour) = True Then
                    MinuteTotal = MinuteTotal + Math.Floor(Hour * 60)
                End If
                If Int32.TryParse(newValue.Substring(2, 2), Minute) = True Then
                    MinuteTotal = MinuteTotal + Minute
                End If
            End If
        Next
        If MinuteTotal > 20 Then
            Hour = MinuteTotal \ 60
            Minute = MinuteTotal Mod 60
            Return String.Format("{0:00}.{1:00}", Hour, Minute)
        Else
            Return String.Format("{0:00}.{1:00}", 0, 0)
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function TotalHoursTest(Value As ArrayList) As String
        On Error GoTo Err
        Dim totMinTotal As Integer = 0
        ' Dim FullTotalHours As String = "00.00"
        Dim totMin As Integer = 0
        Dim totHourToMin As Integer = 0
        Dim arrValue As New ArrayList
        'time

        For i As Integer = 0 To Value.Count - 1
            If Len(Value(i)) = 5 Then
                arrValue = GetFormatSplitValue(Value(i), ".")
                totHourToMin = Math.Floor(arrValue(0) * 60)
                totMin = arrValue(1)
                totMinTotal = totMinTotal + (totHourToMin + totMin)
            End If
        Next
        If totMinTotal > 20 Then
            totMinTotal = totMinTotal Mod 1440
            totHourToMin = totMinTotal \ 60
            totMin = totMinTotal Mod 60
            ' FullTotalHours = String.Format("{0:00}:{1:00}", totHourToMin, totMin)
            '    Return String.Format("{0:00}:{1:00}", totHourToMin, totMin) 'FullTotalHours.ToString()
            'Else
            '    Return FullTotalHours.ToString()
        End If


        'Hour = totMinTotal \ 60
        'Minute = totMinTotal Mod 60


        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function AddTimeToDate(StartDate As DateTime, time As String) As DateTime
        On Error GoTo Err
        'time 11:00
        Dim strDateTime As String = CStr(Mid(StartDate, 1, 10))
        Dim strHour As String = CStr(Mid(time, 1, 2))
        Dim strMin As String = CStr(Mid(time, 4, 2))
        Dim lngHour As Integer = CInt(Mid(time, 1, 2))


        If lngHour >= 12 Then
            strDateTime = "#" + strDateTime + " " + strHour + ":" + strMin + ":00 PM#"
        Else
            strDateTime = "#" + strDateTime + " " + strHour + ":" + strMin + ":00 AM#"
        End If

        AddTimeToDate = strDateTime

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

#End Region
#Region "Format"
    Public Function GetFormatSplitValue(ByVal SplitValue As String, ByVal SplitItem As String) As ArrayList
        'check the Web site is valid
        Dim ReturnArray As New ArrayList
        Dim SplitArray As Array
        Dim SplitCount As Integer
        'Dim r As Integer
        On Error GoTo Err

        SplitArray = SplitValue.Split(SplitItem)
        SplitCount = SplitArray.Length
        For r As Integer = 0 To SplitCount - 1
            If SplitArray(r) = "All" Then
            Else
                ReturnArray.Add(SplitArray(r))
            End If
        Next


        Return ReturnArray

        Exit Function

Err:
        Dim rtn As String = "The error occur with " + SplitValue + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
#End Region
End Class
