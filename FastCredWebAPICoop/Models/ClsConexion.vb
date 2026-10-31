Imports System.Data.SqlClient
Public Class ClsConexion

    Public oConex As SqlConnection

    Public Sub Abrirconexion()

        'Dim sCadConexion As String = "Initial Catalog=DBNEGOCIO;Data Source=190.187.249.243;User=sa;password=tibox*2015;Connection Timeout=0"
        Dim sCadConexion As String = "Initial Catalog=DBNEGOCIO;Data Source=10.19.150.57;User=sa;password=tibox*2015;Connection Timeout=0"

        oConex = New SqlConnection

        oConex.ConnectionString = sCadConexion
        oConex.Open()

    End Sub

    Public Sub CerrarConexion()
        oConex.Close()
    End Sub

End Class
