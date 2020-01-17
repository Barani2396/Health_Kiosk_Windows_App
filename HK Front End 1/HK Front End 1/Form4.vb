Imports System.Data
Imports System.Data.OleDb
Imports HK_Front_End_1.Form3
Public Class Form4

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label7.Text = Form3.Label8.Text
        Label22.Text = Label7.Text
        Dim querry1 As String = "SELECT Age FROM Table1 WHERE [Name]= '" & Label7.Text & "';"
        Dim querry2 As String = "SELECT Gender FROM Table1 WHERE [Name]= '" & Label7.Text & "';"
        Dim querry3 As String = "SELECT [PH No] FROM Table1 WHERE [Name]= '" & Label7.Text & "';"
        Dim querry4 As String = "SELECT [Email ID] FROM Table1 WHERE [Name]= '" & Label7.Text & "';"
        Dim querry5 As String = "SELECT Address FROM Table1 WHERE [Name]= '" & Label7.Text & "';"
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim cmd1 As New OleDbCommand(querry1, con)
        Dim cmd2 As New OleDbCommand(querry2, con)
        Dim cmd3 As New OleDbCommand(querry3, con)
        Dim cmd4 As New OleDbCommand(querry4, con)
        Dim cmd5 As New OleDbCommand(querry5, con)
        con.Open()
        Try
            Label8.Text = cmd1.ExecuteScalar().ToString
            Label9.Text = cmd2.ExecuteScalar().ToString
            Label10.Text = cmd3.ExecuteScalar().ToString
            Label11.Text = cmd4.ExecuteScalar().ToString
            Label12.Text = cmd5.ExecuteScalar().ToString
            con.Close()
            TextBox1.Text = Label8.Text
            ComboBox1.Text = Label9.Text
            TextBox2.Text = Label10.Text
            TextBox3.Text = Label11.Text
            TextBox4.Text = Label12.Text
        Catch ex As Exception
            MsgBox("Error.", MessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form5.Show()
        If Form5.Visible Then
            Me.Hide()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form3.Show()
        Form3.Load()
        If Form3.Visible Then
            Me.Close()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim Pass As String
        Dim querry1 As String = "SELECT Password FROM Table1 WHERE [Name]= '" & Label22.Text & "';"
        Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
        Dim cmd1 As New OleDbCommand(querry1, con)
        con.Open()
        Pass = cmd1.ExecuteScalar().ToString
        If TextBox1.Text = "" Or ComboBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Plz fill all the info.", MessageBoxButtons.OK)
        Else
            If Pass = TextBox5.Text Then
                If TextBox6.Text = TextBox7.Text Then
                    Try
                        Dim upd As String
                        upd = "UPDATE Table1 SET [Age]='" + TextBox1.Text + "',[Gender]='" + ComboBox1.Text + "',[PH No]='" + TextBox2.Text + "',[Email ID]='" + TextBox3.Text + "',[Address]='" + TextBox4.Text + "',[Password]='" + TextBox6.Text + "' WHERE [Name]='" + Label22.Text + "'"
                        Dim cmd2 As New OleDbCommand(upd, con)
                        cmd2.ExecuteNonQuery()
                        con.Close()
                        MsgBox("Update sucess.", MessageBoxButtons.OK)
                        Panel1.Visible = False
                        Form4_Load(Nothing, Nothing)
                    Catch ex As Exception
                        MsgBox("Update failed.", MessageBoxButtons.OK)
                    End Try
                Else
                    MsgBox("Re-enter the password.", MessageBoxButtons.OK)
                    TextBox6.Clear()
                    TextBox7.Clear()
                End If
            Else
                MsgBox("Enter the old password correctly.", MessageBoxButtons.OK)
                TextBox5.Clear()
                TextBox6.Clear()
                TextBox7.Clear()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Panel1.Visible = False
    End Sub

End Class