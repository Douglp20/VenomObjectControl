Imports System.IO
Imports System.Drawing

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
    Function ImageFromBytes(ByVal bytes As Byte()) As System.Drawing.Image
        Using ms As New MemoryStream(bytes)
            Return Image.FromStream(ms)
        End Using

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

    Public Function EmailCheck(ByVal Value As String) As Boolean
        'check the email is valid
        Dim ReturnValue As Boolean = True

        On Error GoTo Err

        If String.IsNullOrEmpty(Value) = False Then
            ReturnValue = Value.Contains("@")
        End If

        Return ReturnValue

        Exit Function

Err:
        Dim rtn As String = "The error occur with " + Value + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)
    End Function
    Public Function WWWCheck(ByVal Value As String) As Boolean
        'check the Web site is valid

        Dim ReturnValue As Boolean = True

        On Error GoTo Err

        If String.IsNullOrEmpty(Value) = False Then
            ReturnValue = Value.Contains(".com") Or Value.Contains(".uk")
        End If



        Exit Function

Err:
        Dim rtn As String = "The error occur with " + Value + " within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
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
