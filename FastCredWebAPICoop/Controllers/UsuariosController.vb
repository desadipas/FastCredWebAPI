Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description


Namespace Controllers
    Public Class UsuariosController
        Inherits ApiController

        <ResponseType(GetType(inUsuarios))>
        Function PostCatalogoCodigos(ByVal CC As inUsuarios) As IHttpActionResult
            Dim R As outUsuarios

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outUsuarios()

            If CC.nSolicitud = 1 Then
                R = CC.RecuperaSolicitud()
            ElseIf CC.nSolicitud = 2 Then
                R = CC.CambiarClave()
            ElseIf CC.nSolicitud = 3 Then
                R = CC.ForzarCambioClave()
            End If


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function



    End Class
End Namespace