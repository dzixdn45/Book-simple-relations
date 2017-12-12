Imports System.Data.OleDb
Public Class Login
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String

    Private Sub login_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didinnuryahya.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        command = New OleDbCommand("select * from pemakai where username='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "'", connection)
        read = command.ExecuteReader
        If (read.Read()) Then
            home.Show()
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()

        Else
            MsgBox("username dan password anda salah mohon dipebaiki lagi", MsgBoxStyle.OkOnly, -"login gagal")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub
End Class