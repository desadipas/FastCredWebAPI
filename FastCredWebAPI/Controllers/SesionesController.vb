Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description
Imports FastCredWebAPI.FastCredWebAPI
Namespace Controllers
    Public Class SesionesController
        Inherits ApiController

        Private db As New DBNEGOCIOEntities
        Function GetSesiones() As IQueryable(Of Usuarios)

            'Dim S As New Seguridad
            'S.GenerarToken()

            Return db.Usuarios

        End Function

        'POST: api/Sesiones
        <ResponseType(GetType(iSesion))>
        Function PostSesiones(ByVal sesion As iSesion) As IHttpActionResult
            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            Dim sesionresp As New SesionResp

            'sesionresp.dFecVenc = CDate("29/05/2019")
            'sesionresp.nResultado = 1
            'sesionresp.sToken = sesion.cClave
            'sesionresp.sRespuesta = sesion.cCodUsu

            'Logueo
            If sesion.nTipoSolicitud = 1 Then
                'Genero token y guardo
                sesionresp = sesion.ValidaLogin()
            End If

            'Cambio de Clave
            If sesion.nTipoSolicitud = 2 Then
                sesionresp = sesion.CambiarClave()
            End If


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, sesionresp)

        End Function



    End Class
End Namespace