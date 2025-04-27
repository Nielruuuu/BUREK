Imports System.Data.OleDb

Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\ramj_\OneDrive\Documents\StudentTimeLog.accdb;")
    Dim cmd As OleDbCommand
    Dim da As OleDbDataAdapter
    Dim dt As DataTable

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadLogs()
    End Sub

    Private Sub LoadLogs()
        Try
            con.Open()
            Dim query As String = "SELECT * FROM TimeLogs ORDER BY TimeIn DESC"
            da = New OleDbDataAdapter(query, con)
            dt = New DataTable()
            da.Fill(dt)
            DataGridView1.DataSource = dt
            con.Close()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            con.Close()
        End Try
    End Sub

    Private Sub btnTimeIn_Click(sender As Object, e As EventArgs) Handles btnTimeIn.Click

        If txtStudentID.Text = "" Or txtStudentName.Text = "" Then
            BunifuSnackbar1.Show(Me, "Please fill in Student ID and Name.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Warning)
            Return
        End If


        If Not IsNumeric(txtStudentID.Text) Then
            BunifuSnackbar1.Show(Me, "Student ID must be numbers only.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            txtStudentID.Focus()
            Return
        End If


        If Not System.Text.RegularExpressions.Regex.IsMatch(txtStudentName.Text, "^[a-zA-Z\s]+$") Then
            BunifuSnackbar1.Show(Me, "Student Name must be letters only.", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            txtStudentName.Focus()
            Return
        End If


        Try
            con.Open()
            Dim query As String = "INSERT INTO TimeLogs (StudentID, StudentName, TimeIn) VALUES (?, ?, ?)"
            cmd = New OleDbCommand(query, con)
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = txtStudentID.Text
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = txtStudentName.Text
            cmd.Parameters.Add("?", OleDbType.Date).Value = DateTime.Now
            cmd.ExecuteNonQuery()
            con.Close()

            BunifuSnackbar1.Show(Me, "Time In Recorded Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success)
            LoadLogs()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, "ERROR: " & ex.Message, Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            con.Close()
        End Try
    End Sub


    Private Sub btnTimeOut_Click(sender As Object, e As EventArgs) Handles btnTimeOut.Click
        If txtStudentID.Text = "" Then
            BunifuSnackbar1.Show(Me, "Please enter Student ID ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            Return
        End If

        Try
            con.Open()
            Dim query As String = "UPDATE TimeLogs SET TimeOut = ? WHERE StudentID = ? AND TimeOut IS NULL"
            cmd = New OleDbCommand(query, con)
            cmd.Parameters.Add("?", OleDbType.Date).Value = DateTime.Now
            cmd.Parameters.Add("?", OleDbType.VarChar).Value = txtStudentID.Text
            Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
            con.Close()

            If rowsAffected > 0 Then
                BunifuSnackbar1.Show(Me, "Time Out Recorded Successfully!", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Success)
            Else
                BunifuSnackbar1.Show(Me, "ERROR: ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            End If

            LoadLogs()
        Catch ex As Exception
            BunifuSnackbar1.Show(Me, "ERROR: ", Bunifu.UI.WinForms.BunifuSnackbar.MessageTypes.Error)
            con.Close()
        End Try
    End Sub

End Class
