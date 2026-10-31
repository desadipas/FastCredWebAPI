Imports System.Data.SqlClient
Public Class inNegocio
    Public Property sToken As String
    Public Property nSolicitud As Integer
    Public Property sDNI As String
    Public Property dFechaNac As Date

    Public Function ValidaLogin() As outNegocio

        Dim R As New outNegocio
        Dim D(2) As String
        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim dr As SqlClient.SqlDataReader
        Dim nResp As Integer
        Dim sRespuesta As String

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_ApiCli_ValidaLogin", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@psDNI", SqlDbType.VarChar, 50).Value = sDNI
            oComm.Parameters.Add("@pdFechaNac", SqlDbType.Date).Value = dFechaNac
            oComm.Parameters.Add("@pnResp", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@psResp", SqlDbType.VarChar, 50).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nResp = oComm.Parameters("@pnResp").Value
            sRespuesta = oComm.Parameters("@psResp").Value

            oConex.Close()

            D(0) = Str(nResp).Trim
            D(1) = sRespuesta

        Catch ex As Exception
            D(0) = "0"
            D(1) = "ERROR: " + ex.Message

        End Try

        R.nResultado = CInt(D(0))
        R.sRespuesta = D(1)

        Return R

    End Function


End Class
