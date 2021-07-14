Imports System.IO
Imports System.Net
Imports System.Runtime.CompilerServices
Imports System.Text
Imports Newtonsoft.Json
Imports System
Imports System.Collections

Public Class ServiceReader


    'Private Const DIR_API As String = "http://192.6.10.239/Odyssey/OODAPI/FetchData/DataView"
    'Private Const COMPANY As String = "OEC"

    Private Const DIR_API As String = "DirApi"
    Private Const X_API_KEY As String = "XApiKey"
    Private Const COMPANY As String = "Company"
    Private Const FILEPARAM As String = "Parameter"
    Private Const SeparadorFile As String = "|"
    Private Const ParamDate As String = "MAINDB.ProductionHistory.Date"
    Private Const Table As String = "TablaBD"

    Public Function ContentParametro(ByVal view As String, ByVal listaParam As List(Of Parametro), ByVal ConfigFile As String)
        Dim content As String
        content = ConstruyeHeader(view, ConfigFile)
        content += ParameterList(listaParam)
        content += "}"
        Return content
    End Function

    Public Function ConstruyeHeader(ByVal view As String, ByVal ConfigFile As String) As String
        Dim JSON As String = "{"
        JSON += """CompanyID"":"
        JSON += String.Format("""{0}"",", ReadConfigParam(COMPANY, ConfigFile))
        JSON += """DataView"":"
        JSON += String.Format("""{0}"",", view)
        Return JSON
    End Function

    Public Function ParameterList(ByVal listaParam As List(Of Parametro)) As String
        Dim JSON As String = """ParameterList"": ["
        Dim cantidad As Integer = 0
        For Each param As Parametro In listaParam
            JSON += agregaParametro(param)
            cantidad += 1
            If cantidad = listaParam.Count Then
                JSON += "]"
            Else
                JSON += ","
            End If
        Next
        Return JSON
    End Function

    Public Function agregaParametro(ByVal param As Parametro)
        Dim paramJson As String = "{"
        paramJson += String.Format("""FieldName"": ""{0}"",", param.FieldName)
        paramJson += String.Format("""Operator"": ""{0}"",", param.Operador)
        paramJson += """ParameterValue"" : """ + param.ParameterValue + """}"
        Return paramJson
    End Function
    Public Function ReadApi(ByVal View As String, ByVal Parameters As String, ByVal Method As String, ByVal ConfigFile As String) As String
        Dim dataMan As New DataManager
        Dim url = ReadConfigParam(DIR_API, ConfigFile)
        Dim response As String
        Dim request As HttpWebRequest = WebRequest.Create(url)
        request.Method = Method
        request.Headers.Add("x-api-key", dataMan.Desencripta(ReadConfigParam(X_API_KEY, ConfigFile)))
        request.ContentType = "application/json"
        request.Accept = "application/json"
        Dim data = Encoding.UTF8.GetBytes(Parameters)

        Using requestStream = request.GetRequestStream
            requestStream.Write(data, 0, data.Length)
            requestStream.Close()

            Using responseStream = request.GetResponse.GetResponseStream
                Using reader As New StreamReader(responseStream)
                    response = reader.ReadToEnd()
                End Using
            End Using
        End Using
        Return response
    End Function

    Public Function TransformaJSONToDatatable(ByVal jsonResponse As String) As DataTable

        Dim jsonObj = JsonConvert.DeserializeObject(jsonResponse)
        'especificamente la propiedad el JSON que tiene el resultado
        Dim resultData = jsonObj("DataSetOut")
        Dim resultados = JsonConvert.SerializeObject(resultData).ToString

        If resultados = "null" Then
            Return New DataTable
        Else
            Dim myData As DataSet = JsonConvert.DeserializeObject(Of DataSet)(resultados)
            Dim myTable As DataTable = myData.Tables(0)
            Return myTable
        End If

    End Function

    Public Function ReadConfigParam(ByVal Parameter As String, ByVal ConfigFile As String) As String
        Dim value As String = ""
        Dim objReader As New StreamReader(ConfigFile)
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                Dim components = sLine.Split(SeparadorFile)
                Dim key = components(0)
                If key = Parameter Then
                    Return components(1)
                End If
            End If
        Loop Until sLine Is Nothing
        objReader.Close()
        Return value
    End Function

    'Lee parametros del archivo de configuración
    Public Function ReadParametersJSON(ByVal ConfigFile As String) As List(Of Parametro)
        Dim objReader As New StreamReader(ConfigFile)
        Dim lista As New List(Of Parametro)
        Dim Tabla As String = ReadConfigParam(Table, ConfigFile)
        lista.Add(getFixedParam(Tabla)) 'agrego parametro fecha
        Dim sLine As String = ""
        Do
            sLine = objReader.ReadLine()
            If Not sLine Is Nothing Then
                Dim components = sLine.Split(SeparadorFile)
                Dim key = components(0)
                If key = FILEPARAM Then
                    Dim value = components(1)
                    Dim argsParam = value.Split(",")
                    Dim parameter As New Parametro
                    With parameter
                        .FieldName = argsParam(0)
                        .Operador = argsParam(1)
                        .ParameterValue = argsParam(2)
                    End With
                    lista.Add(parameter)
                End If
            End If
        Loop Until sLine Is Nothing

        objReader.Close()
        Return lista
    End Function

    'la fecha cambia dinamicamente por eso la calculo
    Public Function getFixedParam(ByVal Tabla As String) As Parametro
        Dim param As New Parametro
        param.FieldName = ParamDate
        param.Operador = Operador.GreaterEQ
        param.ParameterValue = getDateToRead(Tabla)
        Return param
    End Function

    'Obtiene la fecha para la consulta
    Public Function getDateToRead(ByVal Tabla As String) As String
        Dim dataMan As New DataManager
        Dim Fecha As String
        Fecha = dataMan.getLastDateRead(Tabla)
        Return Fecha
    End Function
End Class
