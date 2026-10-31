Imports System.Data.SqlClient
Public Class clsConecta
    Private oConex As SqlConnection
    '***************************************
    'NUBE PRODUCCION
    'Public gsCadConex As String = "Initial Catalog=DBTK;Data Source=20.0.0.31;User=sa;password=tibox*2015;Connection Timeout=0"
    Public gsCadConex As String = "Initial Catalog=DBNEGOCIO;Data Source=20.102.119.245;User=dbaccess;password=sql*1018;Connection Timeout=0"
    '***************************************
    Public Function ConexionActual() As SqlConnection
        Return oConex
    End Function

    Public Function AbrirTransaccion() As SqlTransaction
        AbrirTransaccion = oConex.BeginTransaction()

    End Function
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


    Public Sub CierraConexion()
        oConex.Close()
        oConex = Nothing

    End Sub

End Class
