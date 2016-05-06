Public Class Form1
    Dim SQL As New SQLControl
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        If txtQuery.Text <> "" Then 'keeps it from crashing if button pressed and no text in box
            If SQL.HasConnection = True Then 'checks connection
                SQL.RunQuery(txtQuery.Text)

                If SQL.SQLDataset.Tables.Count > 0 Then 'makes sure data is in table 
                    dgvData.DataSource = SQL.SQLDataset.Tables(0)

                End If
            End If
        End If

    End Sub

    Private Sub btnCreateUser_Click(sender As Object, e As EventArgs) Handles btnCreateUser.Click

        'Query for user to make sure it doesn't already exist
        SQL.RunQuery("SELECT * FROM Users WHERE Users.username = '" & txtUsername.Text & "' ")
        'check row count for username
        If SQL.SQLDataset.Tables(0).Rows.Count > 0 Then
            MsgBox("User already exists")
            Exit Sub
        Else
            CreateUser()
            txtUsername.Clear()
            txtPassword.Clear()
            txtEmailAddress.Clear()
            chkActive.Checked = True
            chkAdmin.Checked = False
        End If
    End Sub
    Public Sub CreateUser()
        If Len(txtUsername.Text) >= 5 And Len(txtPassword.Text) >= 6 Then

            'get the active/admin checkbox values
            Dim intActive As Integer
            Dim intAdmin As Integer
            If chkActive.Checked = True Then intActive = 1 Else intActive = 0
            If chkAdmin.Checked = True Then intAdmin = 1 Else intAdmin = 0

            'Add user to database
            SQL.AddUser(txtUsername.Text, txtPassword.Text, txtEmailAddress.Text, intActive, intAdmin)

        Else
            MsgBox("Username or Password is too short!")
            Exit Sub
        End If
    End Sub

End Class
