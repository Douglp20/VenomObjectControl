Public Class ObjectControl

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
#Region "Null"
    Public Function StringIsNull(ByVal value As String) As String

        On Error GoTo Err
        If String.IsNullOrEmpty(value) Then
            StringIsNull = String.Empty
        Else
            StringIsNull = value
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
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

#End Region
#Region "Keypress"
    Public Function KeyPressNumeric(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        On Error GoTo Err
        If e.KeyChar = vbBack Then
            e.Handled = False
            KeyPressNumeric = False
        Else
            If Char.IsDigit(CChar(CStr(e.KeyChar))) = False Then
                e.Handled = True
                KeyPressNumeric = True
            End If
        End If



        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function
    Public Function KeyPressMoney(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        On Error GoTo Err
        If e.KeyChar = vbBack Then
            e.Handled = False
            KeyPressMoney = False
        Else
            If Asc(e.KeyChar) = 46 Then
                e.Handled = False
                KeyPressMoney = False
            Else
                If Char.IsDigit(CChar(CStr(e.KeyChar))) = False Then
                e.Handled = True
                    KeyPressMoney = True
                End If
            End If
        End If
        'If Asc(e.KeyChar) <> 8 Then
        '    If Asc(e.KeyChar) <> 46 Then
        '        If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
        '            e.Handled = True
        '            KeyPressMoney = True
        '        End If
        '    End If
        'End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

#End Region
#Region "Wording"
    Public Sub FullCapitalise(textbox As System.Windows.Forms.TextBox)
        On Error GoTo Err

        Dim txt = textbox.Text
        Dim len = textbox.Text.ToString().Length
        If len > 0 Then
            Dim rst As String = textbox.Text.ToString.ToUpper()
            textbox.Text = rst
            textbox.SelectionStart = len
        End If
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub


    Public Sub FirstCapitalise(textbox As System.Windows.Forms.TextBox)
        On Error GoTo Err

        Dim txt = textbox.Text
        Dim len As Integer = textbox.Text.ToString().Length
        Dim first_letter As String = ""
        Dim other_letter As String = ""
        Dim result As String = ""

        If len > 1 Then
            first_letter = txt.Substring(0, 1).ToUpper()
            other_letter = txt.Substring(1, len - 1).ToLower()
            result = first_letter.ToString() + other_letter.ToString()
        Else
            result = txt.Substring(0, 1).ToUpper()
        End If
        textbox.Text = result
        textbox.SelectionStart = len
        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
    Public Sub Initialise(Sourcetextbox As System.Windows.Forms.TextBox, Destextbox As System.Windows.Forms.TextBox)
        On Error GoTo Err

        Dim txt = Sourcetextbox.Text
        Dim len As Integer = Sourcetextbox.Text.ToString().Length

        Dim result As String = String.Empty
        Dim arrsplit() As String = txt.Split(" ")
        Dim splitCount As Integer = arrsplit.Count
        Dim myInitials As String = String.Empty

        If splitCount = 1 Then
            result = txt.Substring(0, 1)
        ElseIf splitCount = 2 Then
            result = txt.Substring(0, 1) & txt.Split(" ")(1).Substring(0, 1)
        End If
        Destextbox.Text = result

        Exit Sub

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Sub
#End Region
#Region "Time and Date"
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

#Region "Checks"

    Public Function KeyValidatingMaskedTime(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) As Boolean
        Dim SenderName As String = sender.Name.ToString()
        Dim value As String = String.Empty
        Dim hour As Int32
        Dim minute As Int32

        On Error GoTo Err
        If Len(sender.Text.ToString()) = 5 Then
            value = sender.Text.ToString().Replace(":", "").Trim
            If value = String.Empty Then Return KeyValidatingMaskedTime
            If Int32.TryParse(value.Substring(0, 2), hour) = False Then
                e.Cancel = True
                KeyValidatingMaskedTime = True
            Else
                If hour < 0 OrElse hour > 23 Then
                    e.Cancel = True
                    KeyValidatingMaskedTime = True
                End If
            End If
            If Int32.TryParse(value.Substring(2, 2), minute) = False Then
                e.Cancel = True
                KeyValidatingMaskedTime = True
            Else
                If minute < 0 OrElse minute > 59 Then
                    e.Cancel = True
                    KeyValidatingMaskedTime = True
                End If
            End If
        Else
            e.Cancel = True
            KeyValidatingMaskedTime = True
        End If

        Exit Function

Err:
            Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
            RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function

    Public Function EmailCheck(ByVal strValue As String) As Boolean
        'check the email is valid

        On Error GoTo Err

        If String.IsNullOrEmpty(strValue) = False Then
            Return strValue.Contains("@")
        End If


        Exit Function

Err:
        Dim rtn As String = "The error occur with " + strValue + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function WWWCheck(ByVal strValue As String) As Boolean
        'check the Web site is valid

        On Error GoTo Err

        Return strValue.Contains(".com") Or strValue.Contains(".uk")

        Exit Function

Err:
        Dim rtn As String = "The error occur with " + strValue + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
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
