Imports System.Data.SqlClient

Public Class clsConecta
    Private oConex As SqlConnection
    Private gsNomArchivoIni As String = "Sific_Conexion.ini"
    Private gsNomArchivoDes As String = "SIFIC_DES.ini"

    'CREDITOMOVIL
    'Private gsCadConex As String = "Initial Catalog=DBNEGOCIO;Data Source=10.19.150.57;User=tibox;password=tibox*2015"
    'SBS
    'Private gsCadConex2 As String = "Initial Catalog=SCORE2;Data Source=SVRRIESGOS02;User=dbaccess;password=tibox*2017"


    '***************************************
    'PRODUCCION
    '***************************************
    '***************************************
    'NUBE DESARROLLO
    Public gsCadConex As String = "Initial Catalog=DBNEGOCIO;Data Source=20.102.119.245;User=dbaccess;password=sql*1018;Connection Timeout=0"
    '***************************************
    '***************************************

    'SBS EXTERNO
    'Private gsCadConex2 As String = "Initial Catalog=SCORE2;Data Source=181.65.238.214;User=dbaccess;password=tibox*2017"
    Private gsCadConex2 As String = ""

    'TIBOX Interna
    'Private gsCadConex As String = "Initial Catalog=DBTIBOX;Data Source=190.187.249.243;User=tibox;password=tibox*2015"

    'SVC Interna
    'Private gsCadConex As String = "Initial Catalog=DBSCV;Data Source=10.19.150.57;User=tibox;password=tibox*2015"


    'TIBOX
    'Private gsCadConex As String = "Initial Catalog=DBTIBOX;Data Source=181.65.238.212;User=tibox;password=tibox*2015"


    'NegocioSMS
    'Private gsCadConex As String = "Integrated Security=SSPI;Initial Catalog=DBNEGOCIO;Data Source=SQLCLUS01"

    'Externa
    'Americatel
    'Private gsCadConex As String = "Initial Catalog=DBTIBOX;Data Source=190.187.249.2;User=tibox;password=tibox*2015"


    Public Function ConexionActual() As SqlConnection
        Return oConex
    End Function

    Public Function AbrirTransaccion() As SqlTransaction
        AbrirTransaccion = oConex.BeginTransaction()

    End Function

    Public Sub AbrirConexionSBS()
        Dim C As SqlCommand

        'oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBCore;Data Source=PROYECTO01\SQLEXPRESS")
        'oConex = New SqlConnection("Initial Catalog=DBNEGOCIO;Data Source=10.19.37.241;User=tibox;password=tibox*2015")

        oConex = New SqlConnection(gsCadConex2)
        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
    End Sub

    Public Sub AbrirConexionCreditoMovil()
        Dim C As SqlCommand

        'oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBCore;Data Source=PROYECTO01\SQLEXPRESS")
        oConex = New SqlConnection("Initial Catalog=DBNEGOCIO;Data Source=10.19.37.241;User=tibox;password=tibox*2015")

        oConex = New SqlConnection(gsCadConex)
        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
    End Sub



    Public Sub AbrirConexionTIBOX()
        Dim C As SqlCommand

        'oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBCore;Data Source=PROYECTO01\SQLEXPRESS")
        oConex = New SqlConnection("Initial Catalog=DBTIBOX;Data Source=10.19.37.241;User=tibox;password=tibox*2015")

        oConex = New SqlConnection(gsCadConex)
        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
    End Sub

    Public Sub AbrirConexion()
        Dim C As SqlCommand

        'oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBCore;Data Source=PROYECTO01\SQLEXPRESS")
        'oConex = New SqlConnection("Initial Catalog=DBCORE;Data Source=SRVLAB01;User=sa;password=1234")
        oConex = New SqlConnection(gsCadConex)
        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
    End Sub


    Public Function DameConexion() As SqlConnection
        Dim C As SqlCommand
        'Dim lsSrv As String = ObtenerServidorBaseDeDatos()
        'Dim oEnc As ClsEncript

        'oEnc = New ClsEncript

        'oEnc.DesencriptaArchivo(gsNomArchivoIni, gsNomArchivoDes)
        'gsCadConex = oEnc.DevuelveCadenaConexion(gsNomArchivoDes)

        Try
            My.Computer.FileSystem.DeleteFile(gsNomArchivoDes)
        Catch

        End Try

        oConex = New SqlConnection(gsCadConex)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Return oConex
    End Function


    Public Sub CierraConexion()
        oConex.Close()
        oConex = Nothing

    End Sub


End Class
