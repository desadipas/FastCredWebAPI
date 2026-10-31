Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description
Imports FastCredWebAPI.FastCredWebAPI

Namespace Controllers
    Public Class CatalogoCodigosController
        Inherits ApiController

        Private db As DBNEGOCIOEntities

        'Function GetCatalogoCodigos(ByVal sToken As String) As IQueryable(Of CatalogoCodigos)

        <ResponseType(GetType(inCatalogoCodigos))>
        Function PostCatalogoCodigos(ByVal CC As inCatalogoCodigos) As IHttpActionResult
            Dim R As outCatalogoCodigos

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outCatalogoCodigos()

            R = CC.RecuperaSolicitud()

            'Validar Token

            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function


    End Class

End Namespace