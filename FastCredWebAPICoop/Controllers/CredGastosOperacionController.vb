Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description

Namespace Controllers
    Public Class CredGastosOperacionController
        Inherits ApiController

        <ResponseType(GetType(inCredGastosOperacion))>
        Function PostCredGastosOperacion(ByVal CC As inCredGastosOperacion) As IHttpActionResult
            Dim R As outCredGastosOperacion

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outCredGastosOperacion()

            If CC.nSolicitud = 1 Then

            ElseIf CC.nSolicitud = 2 Then

            ElseIf CC.nSolicitud = 3 Then

            End If


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function



    End Class
End Namespace