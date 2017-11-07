Public Class Form1

    Private attempts_int As Integer = 0 ' number of access attempts
    Private message_str As String = " " 'diplays acess status of users

    'Determine whether code entered is valid or not
    'If valid, open relevant form
    'if invalid , provide error message and track number of access attempts

    Private Sub cmd_enter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_enter.Click

        'Sotres access code entered by user
        Dim acessCode_int As Integer = Val(Txt_SecurityCode.Text)
        Txt_SecurityCode.Clear()

        Select Case acessCode_int

            Case Is < 1000
                message_str = "Assistance required"
                Call DisplayLog()

            Case 1645 To 1689
                message_str = "Technicians"
                Call DisplayLog()
                frmTechnician.ShowDialog()
                attempts_int = 0

            Case 8345
                message_str = "Custodians"
                Call DisplayLog()
                frmCustodian.Show()
                attempts_int = 0
                Me.Hide()

            Case 9998, 1006 To 1008
                message_str = "Scientists"
                Call DisplayLog()
                frmscientist.Show()
                attempts_int = 0
                Me.Hide()

            Case Else ' default if no other case is correct
                If attempts_int < 2 Then
                    message_str = "Access Denied!"
                    Call DisplayLog()
                    MsgBox("Acess Denied. Please try again. Remaining attempts:" & (3 - attempts_int - 1).ToString(), MsgBoxStyle.Critical)
                Else
                    Application.Exit()

                End If
                attempts_int += 1
        End Select
    End Sub

    Private Sub DisplayLog()
        Lst_LogEntry.Items.Insert(0, Date.Now & vbTab & message_str)

    End Sub
    'Handles button clicks for user entering key code
    'one event can be used for all number buttons(0-9)
    'as they all do the same thing basically

    Private Sub numberEntered_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_0.Click, cmd_1.Click, cmd_2.Click, cmd_3.Click, cmd_4.Click, cmd_5.Click, cmd_6.Click, cmd_7.Click, cmd_8.Click, cmd_9.Click

        Dim tmpButton As Button 'creates temporary button object

        tmpButton = sender ' creates reference to last button clicked & assign to temp button

        Txt_SecurityCode.Text &= tmpButton.Text 'add button text to securitycode string


    End Sub
    'clear Security code input textbox
    Private Sub cmd_clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmd_Clear.Click

        Txt_SecurityCode.Clear()


    End Sub
    'rename form on load
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.Text = "Access Control"

    End Sub
End Class
