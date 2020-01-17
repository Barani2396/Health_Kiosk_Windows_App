Imports System.Data
Imports System.Data.OleDb
Imports HK_Front_End_1.Form4
Public Class Form5

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Pass As String
        Dim querry As String = "SELECT Password FROM Table1 WHERE [Name]= '" & Form4.Label7.Text & "';"
        Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim cmd As New OleDbCommand(querry, con)
        con.Open()
        Pass = cmd.ExecuteScalar().ToString
        con.Close()
        If Pass = TextBox1.Text Then
            Form4.Show()
            Form4.Panel1.Show()
            If Form4.Panel1.Visible Then
                Me.Close()
                Form4.TextBox5.Text = ""
                Form4.TextBox6.Text = ""
                Form4.TextBox7.Text = ""
            End If
        Else
            MsgBox("Incorrect Password.", MessageBoxButtons.OK)
            TextBox1.Clear()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form4.Show()
        If Form4.Visible Then
            Me.Close()
        End If
    End Sub

End Class
