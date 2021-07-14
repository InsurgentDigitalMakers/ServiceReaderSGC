
Imports System.IO
Imports System.Data.SqlClient

Public Class DataManager

    'Private Const cadenaConnLocal = "Data Source=DESKTOP-M4JHF48\SQLEXPRESS;Initial Catalog=SistemaCalidad;User ID=sa;Password=becky"
    'Private Const cadenaConnProd = "Data Source=192.168.1.129;Initial Catalog=SistemaCalidad;User ID=Admin;Password=123456"
    Private Const isProduccion As Integer = 0

    Private Const FileConexion As String = "c:\configReader\conexion.txt"
    Private Const KEYCRIPTO = "Novocast"
    Private Const FechaInicial = "01/01/"
    Private Const FormatoFecha = "MM/dd/yyyy"

    Public Function getCadenaConexion() As String
        Dim conexion As String = ""
        Dim objReader As New StreamReader(FileConexion)
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not (sLine Is Nothing) Then
                conexion = Desencripta(sLine)
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return conexion
    End Function

    Public Function Desencripta(ByVal cryptedText As String)
        Dim crypto As New Simple3Des(KEYCRIPTO)
        Return crypto.DecryptData(cryptedText)
    End Function

    Public Sub ManageInsert(ByVal JsonResult As String, ByVal tableName As String, ByVal spCreaTemp As String, ByVal spDropTemp As String, ByVal spUpdateMain As String)
        Dim reader As New ServiceReader
        'creo tabla temporal
        ExecuteStoredProcedure(spCreaTemp)
        Dim table As DataTable = reader.TransformaJSONToDatatable(JsonResult)
        InsertData(table, tableName)
        'actualizo de mi tabla temporal a la principal
        ExecuteStoredProcedure(spUpdateMain)
        'borro tabla temporal
        ExecuteStoredProcedure(spDropTemp)

    End Sub

    Public Sub InsertData(ByVal table As DataTable, ByVal tableName As String)
        If table.Rows.Count > 0 Then
            'insert to db
            Using connection As New SqlConnection(getCadenaConexion)

                connection.Open()
                Using transaction As SqlTransaction = connection.BeginTransaction()
                    Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(connection, SqlBulkCopyOptions.Default, transaction)

                        Try
                            bulkCopy.DestinationTableName = tableName
                            bulkCopy.WriteToServer(table)
                            transaction.Commit()
                        Catch ex As Exception
                            transaction.Rollback()
                            connection.Close()
                            Throw New Exception(ex.Message)
                        End Try
                    End Using
                End Using
            End Using
        End If
    End Sub

    Public Function getHora(ByVal EjecucionId As Integer) As TimeSpan
        Dim hora As TimeSpan
        Dim sql = String.Format("SELECT HoraInicio FROM ServicioReaderTime WHERE EjecucionId={0}", EjecucionId)
        Dim dt = obtenerDataQuery(sql)
        For Each row As DataRow In dt.Rows
            hora = CType(row("HoraInicio"), TimeSpan)
        Next
        Return hora
    End Function
    Public Function getHorasEjecucion() As List(Of ReaderTime)
        Dim sql As String = "SELECT * FROM ServicioReaderTime"
        Dim dt = obtenerDataQuery(sql)
        Return dtToListTime(dt)
    End Function

    Private Function dtToListTime(ByVal tabla As DataTable) As List(Of ReaderTime)
        Dim lista As New List(Of ReaderTime)
        For Each row As DataRow In tabla.Rows
            Dim horaEjecucion As New ReaderTime
            With horaEjecucion
                .EjecucionId = CInt(row("EjecucionId"))
                .HoraInicio = CType(row("HoraInicio"), TimeSpan)
            End With
            lista.Add(horaEjecucion)
        Next
        Return lista
    End Function

    ''' <summary>
    ''' Esta Función regresa falso si no existe un registro en log con la fecha y hora especificadas 
    ''' y verdadero si si existe
    ''' </summary>
    Public Function validaExecutionTime(ByVal Fecha As Date, ByVal Hora As TimeSpan)

        Dim sql As String = "SELECT LogReaderId FROM LogReader WHERE Convert(DATE,Fecha)=Convert(DATE,@Fecha) AND Hora=@Hora"
        Dim Dt As DataTable
        Try
            Using cnSql As New SqlConnection(getCadenaConexion)

                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnSql.State = ConnectionState.Open Then
                    cnSql.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = sql
                    .Connection = cnSql
                    .Parameters.AddWithValue("@Fecha", Fecha)
                    .Parameters.AddWithValue("@Hora", Hora)
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                Da.Fill(Dt)
                If cnSql.State = ConnectionState.Open Then
                    cnSql.Close()
                End If

            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

        'si esta el registro todo esta bien
        If Dt.Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function getLastDateRead(ByVal Tabla As String) As String
        Dim query As String = String.Format("SELECT TOP 1 Fecha FROM LogReader WHERE NombreTabla='{0}' ORDER By Fecha DESC, Hora DESC  ", Tabla)
        Dim dt As DataTable = obtenerDataQuery(query)
        If dt.Rows.Count > 0 Then
            Dim Fecha As String
            Dim FechaBD As Date
            FechaBD = CType(dt.Rows(0)("Fecha"), Date)
            FechaBD = FechaBD.AddDays(-1)
            Fecha = FechaBD.ToString(FormatoFecha)
            Return Fecha
        Else
            Dim year As Integer = Date.Now.Year
            Return FechaInicial + CStr(year)
        End If
    End Function

    Public Sub InsertLog(ByVal fecha As Date, hora As TimeSpan, ByVal NombreTabla As String)
        Try

            Dim sql As String = "INSERT INTO LogReader(Fecha,Hora,NombreTabla)values(@Fecha,@Hora,@NombreTabla)"
            Dim ConnString = getCadenaConexion()

            Using conn As New SqlConnection(ConnString)
                Using comm As New SqlCommand(sql)
                    comm.CommandType = CommandType.Text
                    comm.Connection = conn
                    conn.Open()
                    comm.Parameters.AddWithValue("@Fecha", fecha)
                    comm.Parameters.AddWithValue("@Hora", hora)
                    comm.Parameters.AddWithValue("@NombreTabla", NombreTabla)
                    comm.ExecuteNonQuery()
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Function obtenerDataQuery(ByVal query As String) As DataTable

        Try
            Using cnSql As New SqlConnection(getCadenaConexion)
                Dim Dt As DataTable
                Dim Da As New SqlDataAdapter
                Dim Cmd As New SqlCommand
                If Not cnSql.State = ConnectionState.Open Then
                    cnSql.Open()
                End If
                With Cmd
                    .CommandType = CommandType.Text
                    .CommandText = query
                    .Connection = cnSql
                End With
                Da.SelectCommand = Cmd
                Dt = New DataTable
                Da.Fill(Dt)
                If cnSql.State = ConnectionState.Open Then
                    cnSql.Close()
                End If
                Return Dt
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Sub ExecuteStoredProcedure(ByVal SpName As String)
        Try

            Dim ConnString = getCadenaConexion()
            Using conn As New SqlConnection(ConnString)
                Using comm As New SqlCommand(SpName)
                    comm.CommandType = CommandType.StoredProcedure
                    comm.Connection = conn
                    conn.Open()
                    comm.ExecuteNonQuery()
                End Using
                conn.Close()
            End Using
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub
End Class