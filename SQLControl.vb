Imports System.Data.Sql
Imports System.Data.SqlClient



Public Class SQLControl
    ' Creates a new SQL Connection - requires at least 4 parameters (server pointer, database pointer, credentials"
    Public SQLCon As New SqlConnection With {.ConnectionString = "Server=DESKTOP-VFK5HHV\SQLEXPRESSJEFF; Database=Users;User=sa;Pwd=password; "}
    Public SQLCmd As SqlCommand 'Queries database
    Public SQLDA As SqlDataAdapter
    Public SQLDataset As DataSet

    ''Tests that the Connection can be opened. Returns an exception if it can't
    Public Function HasConnection() As Boolean
        Try
            SQLCon.Open()

            SQLCon.Close()
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    Public Sub RunQuery(Query As String)
        Try
            SQLCon.Open()
            SQLCmd = New SqlCommand(Query, SQLCon)

            ' Load SQL Records for datagrid
            'first set data adapter to reference SQL commands
            SQLDA = New SqlDataAdapter(SQLCmd)   'creates new data adapter ref: query and connection
            SQLDataset = New DataSet 'new instance of data set
            SQLDA.Fill(SQLDataset) 'fill data set

            SQLCon.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
            'Tries closing connection if it misses the close command above.
            If SQLCon.State = ConnectionState.Open Then
                SQLCon.Close()
            End If
        End Try
    End Sub

    Public Sub AddUser(Username As String, Password As String, Email As String, IsActive As Integer, IsAdmin As Integer)
        Try
            Dim strInsert As String = "INSERT INTO Users (Username, Password, Email, Active, Admin)" &
                                      "VALUES (" &
                                      "'" & Username & "'," &
                                      "'" & Password & "'," &
                                      "'" & Email & "'," &
                                      "'" & IsActive & "'," &
                                      "'" & IsAdmin & "') "
            SQLCon.Open()
            SQLCmd = New SqlCommand(strInsert, SQLCon)

            SQLCmd.ExecuteNonQuery()

            SQLCon.Close()

            MsgBox("User successfuly added")
        Catch ex As Exception

        End Try
    End Sub



End Class
