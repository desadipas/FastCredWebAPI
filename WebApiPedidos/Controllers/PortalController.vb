Imports System.Net
Imports System.Web.Http
Imports System.Web.Http.Description

Namespace Controllers
    Public Class PortalController
        Inherits ApiController


        <ResponseType(GetType(inPortal))>
        Function PostPortal(ByVal CC As inPortal) As IHttpActionResult

            Dim R As outPortal

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outPortal()

            Select Case CC.nSolicitud
                Case 0 'nSolicitud = 0 
                    R = CC.RegistraZona()
                Case 1 'nSolicitud = 1 
                    R = CC.RecuperaZonas()
                Case 2 'nSolicitud = 2 
                    R = CC.EliminaZona()
                Case 3 'nSolicitud = 3 
                    R = CC.RegistraUbicacion()
                Case 4 'nSolicitud = 4 
                    R = CC.RecuperaUbicaciones()
                Case 5 'nSolicitud = 5 
                    R = CC.RecuperaUbicacionesDeUsuario()
                Case 6 'nSolicitud = 6 
                    R = CC.RecuperaTodosLosUsuarios()
                Case 7 'nSolicitud = 7 
                    R = CC.RecuperaUsuariosDeZona()
                Case 8
                    R = CC.AsignaZona()
                Case 9
                    R = CC.DesAsignaZona()
                Case 10
                    R = CC.ValidaClave()
                Case 11
                    R = CC.RegistraUsuario()
                Case 12
                    R = CC.ValidaRol()
                Case 13
                    R = CC.RecuperaAgencias()
                Case 14
                    R = CC.RecuperaComercios()
                Case 15
                    R = CC.RecuperaPagoParaAnulacion()
                Case 16
                    R = CC.RegistraAnulaPago()
                Case 17
                    R = CC.RecuperaEntregas()
                Case 18
                    R = CC.RecuperaEntregaACambiar()
                Case 19
                    R = CC.RegistraCambioEntrega()
            End Select


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function





    End Class
End Namespace