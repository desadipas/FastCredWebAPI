Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description

Namespace Controllers
    Public Class ProspectosController
        Inherits ApiController

        <ResponseType(GetType(iProspecto))>
        Function GetProspectos() As IHttpActionResult
            Dim C As New iProspecto

            'C.oProspecto.cCodUsu = "NSILVA"

            Return Ok(C)

        End Function

        ' POST: api/Sesiones
        <ResponseType(GetType(iProspecto))>
        Function PostProspectos(ByVal prospecto As iProspecto) As IHttpActionResult
            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            Dim prospectoresp As New iProspectoResp

            prospectoresp = prospecto.RegistraProspecto()

            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, prospectoresp)

        End Function


    End Class
End Namespace