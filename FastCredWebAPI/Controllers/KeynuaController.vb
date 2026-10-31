Imports System.Data.SqlClient
Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description

Namespace Controllers
    Public Class KeynuaController
        Inherits ApiController

        <ResponseType(GetType(String))>
        Function PostKeynua(ByVal CC As clsKeynua) As IHttpActionResult

            Dim R As String = "KEYNUA"

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            Dim oConex As SqlConnection
            Dim C As SqlCommand
            Dim oComm As SqlClient.SqlCommand
            Dim D() As String
            Dim sCadConexion As String
            Dim oConect As New clsConecta

            sCadConexion = oConect.gsCadConex


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            oComm = New SqlClient.SqlCommand("RegistraKeynuaEvento", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@cComentario", SqlDbType.VarChar, 1500).Value = CC.payload.ToString
            oComm.ExecuteNonQuery()

            oConex.Close()



            Return Ok(R)
            'Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function


    End Class
End Namespace