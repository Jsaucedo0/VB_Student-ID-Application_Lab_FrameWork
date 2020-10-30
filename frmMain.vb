'*******************************************
'*******************************************
'Programmer: Jaime Saucedo
'Course: ITSE 1332.xxxx (VB)
'Program purpose: Student ID Lookup and Pin creation
'GitHub URL: https://github.com/Jsaucedo0/VB_Student-ID-Application_Lab_FrameWork
'*******************************************
'*******************************************

#Region "Compiler_Directives"

'*******************************************
Option Explicit On  'Forces explicit declaration of all variables in a file, or allows implicit declarations of variables
Option Strict On    'Restricts implicit data type conversions to only widening conversions, disallows late binding, and disallows implicit typing
Option Infer Off    'Disables the use of local type inference in declaring variables
#Disable Warning IDE1006    'Disables warnings over class naming convention for controls
'*******************************************

#End Region

Public Class frmMain

    Private bolValidInput As Boolean = False

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        'Exits the Application
        Application.Exit()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load



        Dim strMessage As String =
            "Instructions - " +
            vbNewLine + vbNewLine +
            "Enter your Student ID and then click 'Validate Student ID'" +
            vbNewLine + vbNewLine +
            "If your Student ID is verified, then you can choose your new 7 digit PIN Number " +
            vbNewLine + vbNewLine +
            "If your Student ID is not verified, please re-enter your correct Student ID." +
            vbNewLine + vbNewLine +
            "After you have entered you new PIN, click the 'Verify PIN' button, if correct you are done. " +
            vbNewLine + vbNewLine +
            "Once you have completed selecting your Student ID/PIN combination then you can exit the program or choose to do another Student ID/PIN."

        MessageBox.Show(strMessage, "Student PIN Selection Application", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub btnValidateStudentID_Click(sender As Object, e As EventArgs) Handles btnValidateStudentID.Click

        Try
            Dim bolValidStudentID As Boolean = False
            Dim strValidStudentID() As String =
                {"1234567", "0001112", "4520125", "7895122", "8777541", "8451277", "5555555", "1302850",
                 "8080152", "4562555", "5552012", "5050552", "7825877", "1250255", "1005231", "6545231",
                 "3852085", "7576651", "7881200", "4581002", "7777777”}

            'Loops through array to check if StudentID is valid
            For intCounter As Integer = 0 To strValidStudentID.Length - 1
                If tbxStudentID.Text = strValidStudentID(intCounter) Then
                    bolValidStudentID = True
                End If
            Next

            If tbxStudentID.Text Like "#######" Then 'Valid if 7 digit number

                If bolValidStudentID Then 'Valid StudentID
                    tbxStudentID.Enabled = False
                    lblLookUpResults.Text = "Valid ID"
                    gbxPINSelection.Enabled = True
                    gbxPINSelection.Show()
                    tbxPin.Focus()
                    AcceptButton = btnVerifyPIN

                Else 'Invalid StudentId
                    lblLookUpResults.Text = "Invalid ID"
                    tbxStudentID.Focus()
                    tbxStudentID.SelectAll()
                End If

            Else 'Invalid 7 digit number
                MessageBox.Show("You must only enter numbers and have a length of 7 digits.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tbxStudentID.Focus()
                tbxStudentID.SelectAll()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub BtnClearStudentID_Click(sender As Object, e As EventArgs) Handles BtnClearStudentID.Click
        ClearAll()
    End Sub

    Private Sub btnVerifyPIN_Click(sender As Object, e As EventArgs) Handles btnVerifyPIN.Click

        Try
            If tbxPin.Text Like "#######" Then 'Valid 7 digit number
                Dim strMessage As String =
                "Valid Combination - " +
                 vbNewLine + vbNewLine +
                 "Student ID: " + tbxStudentID.Text + vbNewLine +
                 "Pin: " + tbxPin.Text

                MessageBox.Show(strMessage, "Student PIN Selection Application", MessageBoxButtons.OK, MessageBoxIcon.Information)

            Else 'Invalid 7 digit number
                MessageBox.Show("You must only enter numbers and have a length of 7 digits.", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                tbxPin.Focus()
                tbxPin.SelectAll()
            End If

        Catch ex As Exception
            MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End Try

    End Sub

    Private Sub btnClearPIN_Click(sender As Object, e As EventArgs) Handles btnClearPIN.Click
        tbxPin.Text = ""
        tbxPin.Focus()
    End Sub

    'KeyPress and KeyDown validation for Pin and StudentID textboxes
    Private Sub tbxStudentID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbxStudentID.KeyPress, tbxPin.KeyPress
        'If e.Handled = True don't pass the key on
        'If e.Handled = False do pass the key on
        'Long Form Method:

        If bolValidInput Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Sub tbxStudentID_KeyDown(sender As Object, e As KeyEventArgs) Handles tbxStudentID.KeyDown, tbxPin.KeyDown
        bolValidInput = False
        If My.Computer.Keyboard.ShiftKeyDown Then 'Disallows the use of the shift key
            bolValidInput = False
            Return
        End If

        Select Case e.KeyCode
            Case Keys.D0 To Keys.D9
                bolValidInput = True
            Case Keys.NumPad0 To Keys.NumPad9
                bolValidInput = True
            Case Keys.Back, Keys.Delete
                bolValidInput = True
            Case Keys.Left, Keys.Right
                bolValidInput = True
            Case Else
                bolValidInput = False
        End Select
    End Sub

    Private Sub ClearAll()
        'Resets All fields
        tbxStudentID.Text = ""
        lblLookUpResults.Text = ""
        tbxStudentID.Enabled = True
        tbxPin.Text = ""
        tbxStudentID.Focus()
        AcceptButton = btnValidateStudentID
        gbxPINSelection.Hide()
        gbxPINSelection.Enabled = False
    End Sub

End Class