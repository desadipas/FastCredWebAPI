Imports System.Net
Imports System.Web.Http
Imports System.Web.Http.Description

Namespace Controllers
    Public Class NegocioController
        Inherits ApiController

        <ResponseType(GetType(inNegocio))>
        Function PostNegocio(ByVal CC As inNegocio) As IHttpActionResult

            Dim R As outNegocio

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outNegocio()

            Select Case CC.nSolicitud
                Case 0 'nSolicitud = 0 
                    R = CC.ValidaLogin()

            End Select


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function



    End Class
End Namespace