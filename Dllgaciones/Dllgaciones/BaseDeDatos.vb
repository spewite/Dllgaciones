Imports System.Data.SqlClient
Imports System.Data
Imports System.IO

Public Class BaseDeDatos

    Shared Function ConsultaBBDD(NombreInstanciaBBDD As String, NombreBBDD As String, UserBBDD As String, PasswordBBDD As String, SentenciaSQL As String) As DataTable
        Dim connectionString As String = $"Data Source={NombreInstanciaBBDD}; Initial Catalog={NombreBBDD}; User ID={UserBBDD}; Password={PasswordBBDD}"
        Dim dataTable As New DataTable

        Try
            Using conexion As New SqlConnection(connectionString)
                conexion.Open()

                Using command As New SqlCommand(SentenciaSQL, conexion)
                    Using reader As SqlDataReader = command.ExecuteReader()
                        dataTable.Load(reader)
                    End Using
                End Using
            End Using

            Return dataTable
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error en la base de datos")
            Return New DataTable()
        End Try
    End Function

    Shared Function InsertBBDD(NombreInstanciaBBDD As String, nombreBBDD As String, UserBBDD As String, PasswordBBDD As String, SentenciaSQL As String) As Integer
        Dim connectionString As String = $"Initial Catalog={nombreBBDD}; Data Source={NombreInstanciaBBDD}; User ID={UserBBDD}; Password={PasswordBBDD}"
        Dim filasInsertadas As Integer = 0

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using cmd As New SqlCommand(SentenciaSQL, connection)

                    ' Ejecutar la consulta
                    filasInsertadas = cmd.ExecuteNonQuery()
                End Using
                connection.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical + vbOKOnly, "Error en la BBDD")
        End Try

        Return filasInsertadas
    End Function
    Shared Function UpdateBBDD(NombreInstanciaBBDD As String, nombreBBDD As String, UserBBDD As String, PasswordBBDD As String, SentenciaSQL As String) As Integer
        Dim connectionString As String = $"Initial Catalog={nombreBBDD}; Data Source={NombreInstanciaBBDD}; User ID={UserBBDD}; Password={PasswordBBDD}"
        Dim filasActualizadas As Integer = 0

        Try
            Using connection As New SqlConnection(connectionString)
                connection.Open()
                Using cmd As New SqlCommand(SentenciaSQL, connection)

                    ' Ejecutar la consulta
                    filasActualizadas = cmd.ExecuteNonQuery()
                End Using
                connection.Close()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical + vbOKOnly, "Error en la BBDD")
        End Try

        Return filasActualizadas
    End Function
End Class
