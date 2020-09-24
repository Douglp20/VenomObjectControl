Public Class ObjectControl

    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Public Function KeyPressNumeric(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) As Boolean
        On Error GoTo Err

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
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

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) <> 46 Then
                If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
                    e.Handled = True
                    KeyPressMoney = True
                End If
            End If
        End If

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Function

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

    Public Sub TimeInterval(StartTime As DateTime, Interval As Integer, cbo As System.Windows.Forms.ComboBox)
        On Error GoTo Err
        'time
        Dim EndTime As DateTime = "#23:30:00 PM#"

        cbo.Items.Clear()


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
    Public Function EmailCheck(ByVal strValue As String) As Boolean
        'check the email is valid

        On Error GoTo Err


        Return strValue.Contains("@")


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
End Class
