Imports System.Data.OleDb

Public Class Form1
    Dim conn As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=student_db.accdb;Persist Security Info=False;")
    Dim cmd As OleDbCommand

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtADD.Value = Now
        dtBIRTH.Value = Now
    End Sub

    Private Sub btnSUB_Click(sender As Object, e As EventArgs) Handles btnSUB.Click
        If txtLNAME.Text = "" Or txtFNAME.Text = "" Or txtMNAME.Text = "" Or cboYEAR.SelectedIndex = -1 Or
           txtCN.Text = "" Or txtSTREET.Text = "" Or txtBRGY.Text = "" Or txtMUNI.Text = "" Or
           txtPROV.Text = "" Or txtZIP.Text = "" Or txtGNAME.Text = "" Or txtGFNAME.Text = "" Or
           txtGMNAME.Text = "" Or txtOCC.Text = "" Or txtGNUM.Text = "" Then
            MessageBox.Show("Please fill in all fields before submitting.", "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            Try
                conn.Open()
                Dim query As String = "INSERT INTO students (lname, fname, mname, year_level, course, date_added, birthdate, contact_no, street, brgy, municipality, province, zip_code, g_lname, g_fname, g_mname, g_occupation, g_contact) " &
                                      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
                cmd = New OleDbCommand(query, conn)
                With cmd.Parameters
                    .AddWithValue("@lname", txtLNAME.Text)
                    .AddWithValue("@fname", txtFNAME.Text)
                    .AddWithValue("@mname", txtMNAME.Text)
                    .AddWithValue("@year", cboYEAR.Text)
                    .AddWithValue("@course", cboCOURSE.Text)
                    .AddWithValue("@add", dtADD.Value)
                    .AddWithValue("@birth", dtBIRTH.Value)
                    .AddWithValue("@cn", txtCN.Text)
                    .AddWithValue("@street", txtSTREET.Text)
                    .AddWithValue("@brgy", txtBRGY.Text)
                    .AddWithValue("@muni", txtMUNI.Text)
                    .AddWithValue("@prov", txtPROV.Text)
                    .AddWithValue("@zip", txtZIP.Text)
                    .AddWithValue("@glname", txtGNAME.Text)
                    .AddWithValue("@gfname", txtGFNAME.Text)
                    .AddWithValue("@gmname", txtGMNAME.Text)
                    .AddWithValue("@occ", txtOCC.Text)
                    .AddWithValue("@gnum", txtGNUM.Text)
                End With
                cmd.ExecuteNonQuery()
                MessageBox.Show("Student record saved successfully!")
                ClearFields()
            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            Finally
                conn.Close()
            End Try
        End If
    End Sub

    Private Sub btnDEL_Click(sender As Object, e As EventArgs) Handles btnDEL.Click
        Dim lname = InputBox("Enter last name of student to delete:")
        Try
            conn.Open()
            cmd = New OleDbCommand("DELETE FROM students WHERE lname = ?", conn)
            cmd.Parameters.AddWithValue("@lname", lname)
            Dim rows = cmd.ExecuteNonQuery()
            MessageBox.Show(If(rows > 0, "Student record deleted.", "No record found."))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub btnUP_Click(sender As Object, e As EventArgs) Handles btnUP.Click
        Dim lname = InputBox("Enter last name of student to update:")
        Try
            conn.Open()
            Dim query As String = "UPDATE students SET fname=?, mname=?, year_level=?, birthdate=?, contact_no=? WHERE lname=?"
            cmd = New OleDbCommand(query, conn)
            With cmd.Parameters
                .AddWithValue("@fname", txtFNAME.Text)
                .AddWithValue("@mname", txtMNAME.Text)
                .AddWithValue("@year", cboYEAR.Text)
                .AddWithValue("@birth", dtBIRTH.Value)
                .AddWithValue("@cn", txtCN.Text)
                .AddWithValue("@lname", lname)
            End With
            Dim rows = cmd.ExecuteNonQuery()
            MessageBox.Show(If(rows > 0, "Student record updated.", "No record found."))
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        Finally
            conn.Close()
        End Try
    End Sub

    Private Sub btnCL_Click(sender As Object, e As EventArgs) Handles btnCL.Click
        ClearFields()
    End Sub

    Private Sub btnEX_Click(sender As Object, e As EventArgs) Handles btnEX.Click
        If MessageBox.Show("Are you sure you want to exit?", "Confirm", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub

    Private Sub ClearFields()
        txtLNAME.Clear()
        txtFNAME.Clear()
        txtMNAME.Clear()
        cboYEAR.SelectedIndex = -1
        cboCOURSE.SelectedIndex = -1
        dtADD.Value = Now
        dtBIRTH.Value = Now
        txtCN.Clear()
        txtSTREET.Clear()
        txtBRGY.Clear()
        txtMUNI.Clear()
        txtPROV.Clear()
        txtZIP.Clear()
        txtGNAME.Clear()
        txtGFNAME.Clear()
        txtGMNAME.Clear()
        txtOCC.Clear()
        txtGNUM.Clear()
    End Sub
End Class
