Imports System.IO
Imports ServiceLibrary

Public Class ServiceSGCReader
    Private WithEvents myTimer As System.Timers.Timer

    Dim reader As New ServiceReader
    Dim dataAccess As New DataManager
    Private Const Method As String = "Metodo"
    Private Const Table As String = "TablaBD"
    Private Const SPCreaTablaTemp As String = "SPCreaTemp"
    Private Const SPDropTablaTemp As String = "SPDropTemp"
    Private Const SPUpdateTableMain As String = "SPUpdateMain"
    Private Const KEYCRIPTO = "Novocast"
    Private Const FileProduccion = "c:\configReader\config.txt"
    Private Const FileRechazo = "c:\configReader\configRechazo.txt"
    Private Const FileError = "c:\configReader\logError.txt"
    Private Const FileProdOddissey = "c:\configReader\configProduction.txt"
    Private Const FileConfig = "c:\configReader\configService.txt"
    Private Const FileFrecuency = "ReadFrecuency"
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Agregue el código aquí para iniciar el servicio. Este método debería poner
        ' en movimiento los elementos para que el servicio pueda funcionar.
        Me.myTimer = New System.Timers.Timer()

        Me.myTimer.Enabled = True
        'ejecute cada X minutos ahora leído del archivo
        Dim timeMinutes As Integer = CInt(reader.ReadConfigParam(FileFrecuency, FileConfig))
        Me.myTimer.Interval = 1000 * 60 * timeMinutes

        Me.myTimer.Start()
    End Sub

    Protected Overrides Sub OnStop()
        ' Agregue el código aquí para realizar cualquier anulación necesaria para detener el servicio.
    End Sub

    Protected Sub myTimer_Elapsed(ByVal sender As Object, e As EventArgs) Handles myTimer.Elapsed
        Try
            Dim esHora = validaTiempo()

            If esHora.result Then
                ReadApiProduction(esHora)
                ReadApiRechazo(esHora)
                ReadApiProductionOddissey(esHora)
            End If
        Catch ex As Exception
            WriteErrorLog(ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ReadApiProduction(ByVal esHora As Validacion)
        Try
            Dim listaParam As New List(Of Parametro)
            'llena lista de parametros
            listaParam = reader.ReadParametersJSON(FileProduccion)
            Dim JSONParam = reader.ContentParametro(View.ProduccionMoldeo2, listaParam, FileProduccion)
            Dim jsonResultado = reader.ReadApi(View.ProduccionMoldeo2, JSONParam, reader.ReadConfigParam(Method, FileProduccion), FileProduccion)

            Dim NombreTabla As String = reader.ReadConfigParam(Table, FileProduccion)
            Dim spCreateTemp As String = reader.ReadConfigParam(SPCreaTablaTemp, FileProduccion)
            Dim spDropTemp As String = reader.ReadConfigParam(SPDropTablaTemp, FileProduccion)
            Dim spUpdateMain As String = reader.ReadConfigParam(SPUpdateTableMain, FileProduccion)
            dataAccess.ManageInsert(jsonResultado, NombreTabla, spCreateTemp, spDropTemp, spUpdateMain)
            Dim fecha As Date = Date.Now
            Dim hora As TimeSpan = dataAccess.getHora(esHora.EjecucionId)
            dataAccess.InsertLog(fecha, hora, NombreTabla)
        Catch ex As Exception
            WriteErrorLog(ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ReadApiRechazo(ByVal esHora As Validacion)
        Try
            Dim listaParam As New List(Of Parametro)
            listaParam = reader.ReadParametersJSON(FileRechazo)
            Dim JSONParam = reader.ContentParametro(View.Rechazo, listaParam, FileRechazo)
            Dim jsonResultado = reader.ReadApi(View.Rechazo, JSONParam, reader.ReadConfigParam(Method, FileProduccion), FileProduccion)

            Dim NombreTabla As String = reader.ReadConfigParam(Table, FileRechazo)
            Dim spCreateTemp As String = reader.ReadConfigParam(SPCreaTablaTemp, FileRechazo)
            Dim spDropTemp As String = reader.ReadConfigParam(SPDropTablaTemp, FileRechazo)
            Dim spUpdateMain As String = reader.ReadConfigParam(SPUpdateTableMain, FileRechazo)
            dataAccess.ManageInsert(jsonResultado, NombreTabla, spCreateTemp, spDropTemp, spUpdateMain)
            Dim fecha As Date = Date.Now
            Dim hora As TimeSpan = dataAccess.getHora(esHora.EjecucionId)
            dataAccess.InsertLog(fecha, hora, NombreTabla)
        Catch ex As Exception
            WriteErrorLog(ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Sub ReadApiProductionOddissey(ByVal esHora As Validacion)
        Try
            Dim listaParam As New List(Of Parametro)
            listaParam = reader.ReadParametersJSON(FileProdOddissey)
            Dim JSONParam = reader.ContentParametro(View.Produccion, listaParam, FileProdOddissey)
            Dim jsonResultado = reader.ReadApi(View.Produccion, JSONParam, reader.ReadConfigParam(Method, FileProdOddissey), FileProdOddissey)

            Dim NombreTabla As String = reader.ReadConfigParam(Table, FileProdOddissey)
            Dim spCreateTemp As String = reader.ReadConfigParam(SPCreaTablaTemp, FileProdOddissey)
            Dim spDropTemp As String = reader.ReadConfigParam(SPDropTablaTemp, FileProdOddissey)
            Dim spUpdateMain As String = reader.ReadConfigParam(SPUpdateTableMain, FileProdOddissey)
            dataAccess.ManageInsert(jsonResultado, NombreTabla, spCreateTemp, spDropTemp, spUpdateMain)
            Dim fecha As Date = Date.Now
            Dim hora As TimeSpan = dataAccess.getHora(esHora.EjecucionId)
            dataAccess.InsertLog(fecha, hora, NombreTabla)
        Catch ex As Exception
            WriteErrorLog(ex.Message, ex.StackTrace)
        End Try
    End Sub

    Private Function validaTiempo() As Validacion
        'obtengo la hora actual
        Dim hora As DateTime = DateTime.Now
        Dim fecha As Date = Date.Now
        Dim validacion As New Validacion
        'obtengo los registros
        Dim horas = dataAccess.getHorasEjecucion()
        For Each horaE As ReaderTime In horas
            'si mi hora es superior a un registro de ejecución
            If hora.TimeOfDay >= horaE.HoraInicio Then
                'valido si existe el registro en el log
                If dataAccess.validaExecutionTime(fecha, horaE.HoraInicio) = False Then
                    'si no existe
                    validacion.result = True
                    validacion.EjecucionId = horaE.EjecucionId
                    Return validacion
                End If
            End If
        Next
        ' no es necesario insertar nada
        validacion.result = False
        Return validacion
    End Function

    Private Sub WriteErrorLog(ByVal mensaje As String, ByVal trace As String)
        Dim ruta As String = FileError
        Dim escritor As StreamWriter

        escritor = New StreamWriter(FileError)
        Dim fecha As DateTime = DateTime.Now
        escritor.WriteLine(fecha)
        escritor.WriteLine(mensaje)
        escritor.WriteLine(trace)

        escritor.Flush()
        escritor.Close()
    End Sub
End Class
