Public Class ComboBox
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Public Function FillComboWithMonthName(ByRef cbo As System.Windows.Forms.ComboBox)
        On Error GoTo Err
        Dim startYear As Int32 = Now.Year.ToString
        With cbo
            .Items.Clear()
            .Items.Add("")
            .Items.Add("January")
            .Items.Add("February")
            .Items.Add("March")
            .Items.Add("April")
            .Items.Add("May")
            .Items.Add("June")
            .Items.Add("July")
            .Items.Add("August")
            .Items.Add("September")
            .Items.Add("October")
            .Items.Add("November")
            .Items.Add("December")
        End With

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function FillComboWithYear(ByRef cbo As System.Windows.Forms.ComboBox, ByRef AmountFromYear As Int16, ByRef AmountToYear As Int16)

        On Error GoTo Err

        Dim startYear As Int32 = Now.Year.ToString

        With cbo
            .Items.Clear()
            .Items.Add("")
            For r As Integer = startYear - AmountFromYear To startYear + CLng(AmountToYear)
                .Items.Add(r)
            Next
        End With

        Exit Function

Err:
        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function


End Class
