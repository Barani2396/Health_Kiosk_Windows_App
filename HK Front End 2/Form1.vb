Imports System.Data.OleDb
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Then
            MsgBox("Plz enter patient name.", MessageBoxButtons.OK)
        Else
            Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
            Dim querry As String = "SELECT [PH No] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
            Dim cmd As New OleDbCommand(querry, con)
            con.Open()
            Try
                TextBox2.Text = cmd.ExecuteScalar().ToString
                con.Close()
            Catch ex As Exception
                MsgBox("Patient name doesn't exists.", MessageBoxButtons.OK)
                TextBox1.Text = ""
            End Try
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        Panel1.Visible = True
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim querry1 As String = "SELECT [Heart Rate] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry2 As String = "SELECT [Blood Sugar] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry3 As String = "SELECT [Body Temperature] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry4 As String = "SELECT [Blood Pressure] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry5 As String = "SELECT [Prev-Heart Rate] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry6 As String = "SELECT [Prev-Blood Sugar] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry7 As String = "SELECT [Prev-Body Temperature] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim querry8 As String = "SELECT [Prev-Blood Pressure] FROM Table1 WHERE [Name]= '" & TextBox1.Text & "';"
        Dim cmd1 As New OleDbCommand(querry1, con)
        Dim cmd2 As New OleDbCommand(querry2, con)
        Dim cmd3 As New OleDbCommand(querry3, con)
        Dim cmd4 As New OleDbCommand(querry4, con)
        Dim cmd5 As New OleDbCommand(querry5, con)
        Dim cmd6 As New OleDbCommand(querry6, con)
        Dim cmd7 As New OleDbCommand(querry7, con)
        Dim cmd8 As New OleDbCommand(querry8, con)
        con.Open()
        Label10.Text = cmd1.ExecuteScalar().ToString
        Label11.Text = cmd2.ExecuteScalar().ToString
        Label12.Text = cmd3.ExecuteScalar().ToString
        Label13.Text = cmd4.ExecuteScalar().ToString
        Label14.Text = cmd5.ExecuteScalar().ToString
        Label15.Text = cmd6.ExecuteScalar().ToString
        Label16.Text = cmd7.ExecuteScalar().ToString
        Label17.Text = cmd8.ExecuteScalar().ToString
        con.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel1.Visible = False
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

End Class

