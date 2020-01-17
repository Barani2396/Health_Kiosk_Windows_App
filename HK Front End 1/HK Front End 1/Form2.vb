Imports System.Data
Imports System.Data.OleDb
Public Class Form2

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox1.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Plz fill all the info.", MessageBoxButtons.OK)
        Else
            If TextBox6.Text = TextBox7.Text Then
                Try
                    Dim querry As String = "SELECT COUNT(*) AS numRows FROM Table1 WHERE [Name] = '" & TextBox1.Text & "'"
                    Dim queryResult As Integer
                    Dim con As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Barani\Documents\Project\Database\HK-Database.mdb")
                    con.Open()
                    Dim com As New OleDbCommand(querry, con)
                    queryResult = com.ExecuteScalar()
                    con.Close()
                    If queryResult > 0 Then
                        MsgBox("Username already exists.", MessageBoxButtons.OK)
                        TextBox1.Text = ""
                        TextBox2.Text = ""
                        ComboBox1.Text = "Choose Gender"
                        TextBox3.Text = ""
                        TextBox4.Text = ""
                        TextBox5.Text = ""
                        TextBox6.Text = ""
                        TextBox7.Text = ""
                    Else
                        con.Open()
                        Dim ins As String = "INSERT INTO Table1 ([Name], [Age], [Gender], [PH No], [Email ID], [Address], [Password])" + "VALUES ('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox1.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "');"
                        Dim cmd As New OleDbCommand(ins, con)
                        cmd.ExecuteNonQuery()
                        con.Close()
                        MsgBox("Create sucess.", MessageBoxButtons.OK)
                        Form3.CallF3(2)
                        Form3.Show()
                        Form3.Label2.Hide()
                        Form3.Button1.Hide()
                        If Form3.Visible Then
                            Me.Close()
                        End If
                    End If
                Catch ex As Exception
                    MsgBox("Error.", MessageBoxButtons.OK)
                End Try
            Else
                MsgBox("Re-enter the password.", MessageBoxButtons.OK)
                TextBox6.Clear()
                TextBox7.Clear()
            End If
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.Show()
        If Form1.Visible Then
            Me.Close()
        End If
    End Sub

End Class