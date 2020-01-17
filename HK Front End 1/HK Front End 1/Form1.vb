Imports System.Data
Imports System.Data.OleDb
Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Name As String = ""
        Dim Pass1 As String
        Dim Pass2 As String
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Plz fill all the info.", MessageBoxButtons.OK)
            TextBox1.Clear()
            TextBox2.Clear()
        Else
            Name = TextBox1.Text
            Pass1 = TextBox2.Text
            Dim querry As String = "SELECT Password FROM Table1 WHERE [Name]= '" & Name & "';"
            Dim con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")                                         
            Dim cmd As New OleDbCommand(querry, con)
            con.Open()
            Try
                Pass2 = cmd.ExecuteScalar().ToString
                con.Close()
                If (Pass1 = Pass2) Then
                    MsgBox("Login success.", MessageBoxButtons.OK)
                    Form3.CallF3(1)
                    Form3.Show()
                    If Form3.Visible Then
                        Me.Close()
                    End If
                Else
                    MsgBox("Login failed.", MessageBoxButtons.OK)
                    TextBox1.Clear()
                    TextBox2.Clear()
                End If
            Catch ex As Exception
                MsgBox("Username doesn't exists.", MessageBoxButtons.OK)
                TextBox1.Clear()
                TextBox2.Clear()
            End Try
            
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Form2.Show()
        If Form2.Visible Then
            Me.Close()
        End If
    End Sub

End Class
