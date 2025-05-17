Imports System.Data.OleDb

Public Class Form1
    ' Connection string should be readonly to prevent accidental changes
    Private ReadOnly connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=student_db.accdb;Persist Security Info=False;"

    ' Create connections when needed rather than keeping a class-level connection
    Private Function CreateConnection() As OleDbConnection
        Return New OleDbConnection(connectionString)
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        dtADD.Value = Now
        dtBIRTH.Value = Now

        ' Setup filter combo
        filterbycombo.SelectedIndex = 0 ' Default to Last Name

        ' Setup and load the ListView data
        LoadStudentData()
    End Sub

    Private Sub btnSUB_Click(sender As Object, e As EventArgs) Handles btnSUB.Click
        ' Validate all required fields
        If txtLNAME.Text = "" Then
            MsgBox("Enter last name", vbInformation, "Missing")
            txtLNAME.Focus()
            Exit Sub
        ElseIf txtFNAME.Text = "" Then
            MsgBox("Enter first name", vbInformation, "Missing")
            txtFNAME.Focus()
            Exit Sub
        ElseIf txtMNAME.Text = "" Then
            MsgBox("Enter middle name", vbInformation, "Missing")
            txtMNAME.Focus()
            Exit Sub
        ElseIf cboYEAR.SelectedIndex = -1 Then
            MsgBox("Select year level", vbInformation, "Missing")
            cboYEAR.Focus()
            Exit Sub
        ElseIf txtCN.Text = "" Then
            MsgBox("Enter contact number", vbInformation, "Missing")
            txtCN.Focus()
            Exit Sub
        ElseIf txtSTREET.Text = "" Then
            MsgBox("Enter street", vbInformation, "Missing")
            txtSTREET.Focus()
            Exit Sub
        ElseIf txtBRGY.Text = "" Then
            MsgBox("Enter barangay", vbInformation, "Missing")
            txtBRGY.Focus()
            Exit Sub
        ElseIf txtMUNI.Text = "" Then
            MsgBox("Enter municipality", vbInformation, "Missing")
            txtMUNI.Focus()
            Exit Sub
        ElseIf txtPROV.Text = "" Then
            MsgBox("Enter province", vbInformation, "Missing")
            txtPROV.Focus()
            Exit Sub
        ElseIf txtZIP.Text = "" Then
            MsgBox("Enter zip code", vbInformation, "Missing")
            txtZIP.Focus()
            Exit Sub
        ElseIf txtGNAME.Text = "" Then
            MsgBox("Enter guardian last name", vbInformation, "Missing")
            txtGNAME.Focus()
            Exit Sub
        ElseIf txtGFNAME.Text = "" Then
            MsgBox("Enter guardian first name", vbInformation, "Missing")
            txtGFNAME.Focus()
            Exit Sub
        ElseIf txtGMNAME.Text = "" Then
            MsgBox("Enter guardian middle name", vbInformation, "Missing")
            txtGMNAME.Focus()
            Exit Sub
        ElseIf txtOCC.Text = "" Then
            MsgBox("Enter guardian occupation", vbInformation, "Missing")
            txtOCC.Focus()
            Exit Sub
        ElseIf txtGNUM.Text = "" Then
            MsgBox("Enter guardian contact number", vbInformation, "Missing")
            txtGNUM.Focus()
            Exit Sub
        ElseIf txtEMAIL.Text = "" Then
            MsgBox("Enter email", vbInformation, "Missing")
            txtEMAIL.Focus()
            Exit Sub
        ElseIf txtHS.Text = "" Then
            MsgBox("Enter high school", vbInformation, "Missing")
            txtHS.Focus()
            Exit Sub
        ElseIf txtCOLL.Text = "" Then
            MsgBox("Enter college", vbInformation, "Missing")
            txtCOLL.Focus()
            Exit Sub
        ElseIf txtYG.Text = "" Then
            MsgBox("Enter year graduated", vbInformation, "Missing")
            txtYG.Focus()
            Exit Sub
        ElseIf txtSY.Text = "" Then
            MsgBox("Enter school year", vbInformation, "Missing")
            txtSY.Focus()
            Exit Sub
        End If

        Using conn As OleDbConnection = CreateConnection()
            Using cmd As New OleDbCommand()
                Try
                    conn.Open()
                    Dim query As String = "INSERT INTO students " &
                        "(lname, fname, mname, year_level, course, date_added, birthdate, contact_no, street, brgy, municipality, province, zip_code, " &
                        "g_lname, g_fname, g_mname, g_occupation, g_contact, email, highschool, college, yeargraduate, school_year) " &
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

                    cmd.Connection = conn
                    cmd.CommandText = query

                    ' Convert all text inputs to uppercase before inserting except email which remains as is
                    txtLNAME.Text = txtLNAME.Text.ToUpper()
                    txtFNAME.Text = txtFNAME.Text.ToUpper()
                    txtMNAME.Text = txtMNAME.Text.ToUpper()
                    txtCN.Text = txtCN.Text.ToUpper()
                    txtSTREET.Text = txtSTREET.Text.ToUpper()
                    txtBRGY.Text = txtBRGY.Text.ToUpper()
                    txtMUNI.Text = txtMUNI.Text.ToUpper()
                    txtPROV.Text = txtPROV.Text.ToUpper()
                    txtZIP.Text = txtZIP.Text.ToUpper()
                    txtGNAME.Text = txtGNAME.Text.ToUpper()
                    txtGFNAME.Text = txtGFNAME.Text.ToUpper()
                    txtGMNAME.Text = txtGMNAME.Text.ToUpper()
                    txtOCC.Text = txtOCC.Text.ToUpper()
                    txtGNUM.Text = txtGNUM.Text.ToUpper()

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
                        .AddWithValue("@email", txtEMAIL.Text)             ' email case maintained
                        .AddWithValue("@highschool", txtHS.Text.ToUpper())
                        .AddWithValue("@college", txtCOLL.Text.ToUpper())
                        .AddWithValue("@yeargraduate", txtYG.Text.ToUpper())
                        .AddWithValue("@school_year", txtSY.Text.ToUpper())
                    End With

                    cmd.ExecuteNonQuery()
                    MsgBox("Student record saved successfully!", vbInformation, "Success")
                    ClearFields()
                    LoadStudentData()
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message, vbCritical, "Error")
                End Try
            End Using
        End Using
    End Sub

    Private Sub btnDEL_Click(sender As Object, e As EventArgs) Handles btnDEL.Click
        If txtLNAME.Text = "" Then
            MsgBox("Please select a student to delete", vbInformation, "Missing")
            Exit Sub
        End If

        If MsgBox("Are you sure you want to delete this student record?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then
            Exit Sub
        End If

        Using conn As OleDbConnection = CreateConnection()
            Using cmd As New OleDbCommand("DELETE FROM students WHERE lname = ? AND fname = ?", conn)
                Try
                    conn.Open()
                    cmd.Parameters.AddWithValue("@lname", txtLNAME.Text)
                    cmd.Parameters.AddWithValue("@fname", txtFNAME.Text)

                    Dim rows = cmd.ExecuteNonQuery()
                    If rows > 0 Then
                        MsgBox("Student record deleted successfully!", vbInformation, "Success")
                        ClearFields()
                        LoadStudentData()
                    Else
                        MsgBox("No matching student record found.", vbInformation, "Not Found")
                    End If
                Catch ex As Exception
                    MsgBox("Error: " & ex.Message, vbCritical, "Error")
                End Try
            End Using
        End Using
    End Sub

    Private Sub btnUP_Click(sender As Object, e As EventArgs) Handles btnUP.Click
        ' Validation checks
        If txtLNAME.Text = "" Then
            MsgBox("Enter last name", vbInformation, "Missing")
            txtLNAME.Focus()
            Exit Sub
        ElseIf txtFNAME.Text = "" Then
            MsgBox("Enter first name", vbInformation, "Missing")
            txtFNAME.Focus()
            Exit Sub
        ElseIf txtMNAME.Text = "" Then
            MsgBox("Enter middle name", vbInformation, "Missing")
            txtMNAME.Focus()
            Exit Sub
        ElseIf cboYEAR.Text = "" Then
            MsgBox("Select year level", vbInformation, "Missing")
            cboYEAR.Focus()
            Exit Sub
        ElseIf txtCN.Text = "" Then
            MsgBox("Enter contact number", vbInformation, "Missing")
            txtCN.Focus()
            Exit Sub
        ElseIf txtEMAIL.Text = "" Then
            MsgBox("Enter email", vbInformation, "Missing")
            txtEMAIL.Focus()
            Exit Sub
        ElseIf txtHS.Text = "" Then
            MsgBox("Enter high school", vbInformation, "Missing")
            txtHS.Focus()
            Exit Sub
        ElseIf txtCOLL.Text = "" Then
            MsgBox("Enter college", vbInformation, "Missing")
            txtCOLL.Focus()
            Exit Sub
        ElseIf txtYG.Text = "" Then
            MsgBox("Enter year graduated", vbInformation, "Missing")
            txtYG.Focus()
            Exit Sub
        ElseIf txtSY.Text = "" Then
            MsgBox("Enter school year", vbInformation, "Missing")
            txtSY.Focus()
            Exit Sub
        End If

        If MsgBox("Do you really want to update this student record?", vbQuestion + vbYesNo, "Update") = vbNo Then
            Exit Sub
        End If

        Using conn As OleDbConnection = CreateConnection()
            Using cmd As New OleDbCommand()
                Try
                    conn.Open()
                    Dim query As String = "UPDATE students SET fname=?, mname=?, year_level=?, birthdate=?, contact_no=?, email=?, highschool=?, college=?, yeargraduate=?, school_year=? WHERE lname=?"

                    cmd.Connection = conn
                    cmd.CommandText = query

                    ' Convert all inputs except email to uppercase before updating
                    txtLNAME.Text = txtLNAME.Text.ToUpper()
                    txtFNAME.Text = txtFNAME.Text.ToUpper()
                    txtMNAME.Text = txtMNAME.Text.ToUpper()
                    txtCN.Text = txtCN.Text.ToUpper()
                    txtSTREET.Text = txtSTREET.Text.ToUpper()
                    txtBRGY.Text = txtBRGY.Text.ToUpper()
                    txtMUNI.Text = txtMUNI.Text.ToUpper()
                    txtPROV.Text = txtPROV.Text.ToUpper()
                    txtZIP.Text = txtZIP.Text.ToUpper()
                    txtGNAME.Text = txtGNAME.Text.ToUpper()
                    txtGFNAME.Text = txtGFNAME.Text.ToUpper()
                    txtGMNAME.Text = txtGMNAME.Text.ToUpper()
                    txtOCC.Text = txtOCC.Text.ToUpper()
                    txtGNUM.Text = txtGNUM.Text.ToUpper()

                    With cmd.Parameters
                        .AddWithValue("@fname", txtFNAME.Text)
                        .AddWithValue("@mname", txtMNAME.Text)
                        .AddWithValue("@year", cboYEAR.Text)
                        .AddWithValue("@birth", dtBIRTH.Value)
                        .AddWithValue("@cn", txtCN.Text)
                        .AddWithValue("@email", txtEMAIL.Text) ' email case maintained
                        .AddWithValue("@highschool", txtHS.Text.ToUpper())
                        .AddWithValue("@college", txtCOLL.Text.ToUpper())
                        .AddWithValue("@yeargraduate", txtYG.Text.ToUpper())
                        .AddWithValue("@school_year", txtSY.Text.ToUpper())
                        .AddWithValue("@lname", txtLNAME.Text)
                    End With

                    Dim rows = cmd.ExecuteNonQuery()
                    MessageBox.Show(If(rows > 0, "Student record was updated successfully.", "No record found."), "Update", MessageBoxButtons.OK, MessageBoxIcon.Information)

                    LoadStudentData()
                    ClearFields()
                Catch ex As Exception
                    MessageBox.Show("Error: " & ex.Message)
                End Try
            End Using
        End Using
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
        txtEMAIL.Clear()
        txtHS.Clear()
        txtCOLL.Clear()
        txtYG.Clear()
        txtSY.Clear()
    End Sub

    Public Sub LoadStudentData()
        ListView1.Items.Clear()

        If ListView1.Columns.Count = 0 Then
            ListView1.View = View.Details
            ListView1.FullRowSelect = True
            ListView1.GridLines = True

            ListView1.Columns.Add("Last Name", 100)
            ListView1.Columns.Add("First Name", 100)
            ListView1.Columns.Add("Middle Name", 100)
            ListView1.Columns.Add("Year Level", 80)
            ListView1.Columns.Add("Course", 100)
            ListView1.Columns.Add("Date Added", 100)
            ListView1.Columns.Add("Birthdate", 100)
            ListView1.Columns.Add("Contact No", 100)
            ListView1.Columns.Add("Street", 100)
            ListView1.Columns.Add("Barangay", 100)
            ListView1.Columns.Add("Municipality", 100)
            ListView1.Columns.Add("Province", 100)
            ListView1.Columns.Add("ZIP Code", 80)
            ListView1.Columns.Add("Guardian LName", 100)
            ListView1.Columns.Add("Guardian FName", 100)
            ListView1.Columns.Add("Guardian MName", 100)
            ListView1.Columns.Add("Occupation", 100)
            ListView1.Columns.Add("Guardian Contact", 100)
            ListView1.Columns.Add("Email", 150)
            ListView1.Columns.Add("High School", 150)
            ListView1.Columns.Add("College", 150)
            ListView1.Columns.Add("Year Graduated", 100)
            ListView1.Columns.Add("School Year", 100)
        End If

        Using conn As OleDbConnection = CreateConnection()
            Using cmd As New OleDbCommand("SELECT * FROM students ORDER BY lname ASC", conn)
                Try
                    conn.Open()
                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        While reader.Read()
                            Dim item As New ListViewItem(reader("lname").ToString())
                            item.SubItems.Add(reader("fname").ToString())
                            item.SubItems.Add(reader("mname").ToString())
                            item.SubItems.Add(reader("year_level").ToString())
                            item.SubItems.Add(reader("course").ToString())
                            item.SubItems.Add(reader("date_added").ToString())
                            item.SubItems.Add(reader("birthdate").ToString())
                            item.SubItems.Add(reader("contact_no").ToString())
                            item.SubItems.Add(reader("street").ToString())
                            item.SubItems.Add(reader("brgy").ToString())
                            item.SubItems.Add(reader("municipality").ToString())
                            item.SubItems.Add(reader("province").ToString())
                            item.SubItems.Add(reader("zip_code").ToString())
                            item.SubItems.Add(reader("g_lname").ToString())
                            item.SubItems.Add(reader("g_fname").ToString())
                            item.SubItems.Add(reader("g_mname").ToString())
                            item.SubItems.Add(reader("g_occupation").ToString())
                            item.SubItems.Add(reader("g_contact").ToString())
                            item.SubItems.Add(reader("email").ToString())
                            item.SubItems.Add(reader("highschool").ToString())
                            item.SubItems.Add(reader("college").ToString())
                            item.SubItems.Add(reader("yeargraduate").ToString())
                            item.SubItems.Add(reader("school_year").ToString())

                            If Not reader.IsDBNull(reader.GetOrdinal("ID")) Then
                                item.Tag = reader("ID")
                            End If

                            ListView1.Items.Add(item)
                        End While
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error loading student data: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems.Count > 0 Then
            Dim selectedItem As ListViewItem = ListView1.SelectedItems(0)

            txtLNAME.Text = selectedItem.SubItems(0).Text
            txtFNAME.Text = selectedItem.SubItems(1).Text
            txtMNAME.Text = selectedItem.SubItems(2).Text
            cboYEAR.Text = selectedItem.SubItems(3).Text
            cboCOURSE.Text = selectedItem.SubItems(4).Text

            Dim addDate As DateTime
            Dim birthDate As DateTime

            If DateTime.TryParse(selectedItem.SubItems(5).Text, addDate) Then
                dtADD.Value = addDate
            End If

            If DateTime.TryParse(selectedItem.SubItems(6).Text, birthDate) Then
                dtBIRTH.Value = birthDate
            End If

            txtCN.Text = selectedItem.SubItems(7).Text
            txtSTREET.Text = selectedItem.SubItems(8).Text
            txtBRGY.Text = selectedItem.SubItems(9).Text
            txtMUNI.Text = selectedItem.SubItems(10).Text
            txtPROV.Text = selectedItem.SubItems(11).Text
            txtZIP.Text = selectedItem.SubItems(12).Text
            txtGNAME.Text = selectedItem.SubItems(13).Text
            txtGFNAME.Text = selectedItem.SubItems(14).Text
            txtGMNAME.Text = selectedItem.SubItems(15).Text
            txtOCC.Text = selectedItem.SubItems(16).Text
            txtGNUM.Text = selectedItem.SubItems(17).Text
            txtEMAIL.Text = selectedItem.SubItems(18).Text
            txtHS.Text = selectedItem.SubItems(19).Text
            txtCOLL.Text = selectedItem.SubItems(20).Text
            txtYG.Text = selectedItem.SubItems(21).Text
            txtSY.Text = selectedItem.SubItems(22).Text

            txtFNAME.Focus()
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles searchTxtbox.TextChanged
        Using conn As OleDbConnection = CreateConnection()
            Using cmd As New OleDbCommand()
                Try
                    conn.Open()
                    Dim query As String = ""

                    If filterbycombo.SelectedIndex = 0 Then
                        query = "SELECT * FROM students WHERE lname LIKE ? ORDER BY lname ASC"
                    ElseIf filterbycombo.SelectedIndex = 1 Then
                        query = "SELECT * FROM students WHERE course LIKE ? ORDER BY lname ASC"
                    ElseIf filterbycombo.SelectedIndex = 2 Then
                        query = "SELECT * FROM students WHERE year_level LIKE ? ORDER BY lname ASC"
                    ElseIf filterbycombo.SelectedIndex = 3 Then
                        query = "SELECT * FROM students WHERE municipality LIKE ? ORDER BY lname ASC"
                    Else
                        Return
                    End If

                    cmd.Connection = conn
                    cmd.CommandText = query
                    cmd.Parameters.AddWithValue("@search", "%" & searchTxtbox.Text & "%")

                    ListView1.Items.Clear()

                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        If Not reader.HasRows Then Return

                        While reader.Read()
                            Dim item As New ListViewItem(reader("lname").ToString())
                            item.SubItems.Add(reader("fname").ToString())
                            item.SubItems.Add(reader("mname").ToString())
                            item.SubItems.Add(reader("year_level").ToString())
                            item.SubItems.Add(reader("course").ToString())
                            item.SubItems.Add(reader("date_added").ToString())
                            item.SubItems.Add(reader("birthdate").ToString())
                            item.SubItems.Add(reader("contact_no").ToString())
                            item.SubItems.Add(reader("street").ToString())
                            item.SubItems.Add(reader("brgy").ToString())
                            item.SubItems.Add(reader("municipality").ToString())
                            item.SubItems.Add(reader("province").ToString())
                            item.SubItems.Add(reader("zip_code").ToString())
                            item.SubItems.Add(reader("g_lname").ToString())
                            item.SubItems.Add(reader("g_fname").ToString())
                            item.SubItems.Add(reader("g_mname").ToString())
                            item.SubItems.Add(reader("g_occupation").ToString())
                            item.SubItems.Add(reader("g_contact").ToString())
                            item.SubItems.Add(reader("email").ToString())
                            item.SubItems.Add(reader("highschool").ToString())
                            item.SubItems.Add(reader("college").ToString())
                            item.SubItems.Add(reader("yeargraduate").ToString())
                            item.SubItems.Add(reader("school_year").ToString())

                            If Not reader.IsDBNull(reader.GetOrdinal("ID")) Then
                                item.Tag = reader("ID")
                            End If

                            ListView1.Items.Add(item)
                        End While
                    End Using
                Catch ex As Exception
                    MessageBox.Show("Error searching student data: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End Using
        End Using
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles filterbycombo.SelectedIndexChanged
        If searchTxtbox.Text.Trim() <> "" Then
            TextBox1_TextChanged(sender, e)
        End If
    End Sub

    ' Unused event handlers can remain empty or be removed if not needed
    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
    End Sub

    Private Sub Label29_Click(sender As Object, e As EventArgs) Handles Label29.Click
    End Sub

    Private Sub Label30_Click(sender As Object, e As EventArgs) Handles Label30.Click
    End Sub

    Private Sub txtEMAIL_TextChanged(sender As Object, e As EventArgs) Handles txtEMAIL.TextChanged
    End Sub

    Private Sub sexcombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles sexcombo.SelectedIndexChanged
    End Sub

    Private Sub txtHS_TextChanged(sender As Object, e As EventArgs) Handles txtHS.TextChanged
    End Sub

    Private Sub txtCOLL_TextChanged(sender As Object, e As EventArgs) Handles txtCOLL.TextChanged
    End Sub

    Private Sub txtYG_TextChanged(sender As Object, e As EventArgs) Handles txtYG.TextChanged
    End Sub

    Private Sub txtSY_TextChanged(sender As Object, e As EventArgs) Handles txtSY.TextChanged
    End Sub

    Private Sub Label24_Click(sender As Object, e As EventArgs) Handles Label24.Click
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class