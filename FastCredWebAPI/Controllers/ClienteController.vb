Imports System.Net
Imports System.Web.Http
Imports System.Web.Http.Description

Namespace Controllers
    Public Class ClienteController
        Inherits ApiController


        <ResponseType(GetType(inNegocioCliente))>
        Function PostNegocio(ByVal CC As inNegocioCliente) As IHttpActionResult

            Dim R As outNegocioCliente

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outNegocioCliente()

            Select Case CC.nSolicitud
                Case 1
                    R = CC.ValidaClaveCliente()
                Case 2
                    R = CC.RegistraSMS()
                Case 3
                    R = CC.RecuperaCatalogo()
                Case 4
                    R = CC.RecuperaZonas()
                Case 5
                    R = CC.RegistraSMSXDNI()
                Case 6
                    R = CC.RegistraUsuarioExistente()
                Case 7
                    R = CC.RegistraUsuarioNuevo()
                Case 8
                    R = CC.RecuperaCreditos()
                Case 9
                    R = CC.RecuperaPagosCliente()
                Case 10
                    R = CC.RecuperaCuotasCliente()
                Case 11
                    R = CC.RecuperaPrecalifiNew()
                Case 12
                    R = CC.RegistraSolicitudNueva()
                Case 13
                    R = CC.ValidaZona()
                Case 14
                    R = CC.RegistraSolicitudNuevaRapida()

                Case 99
                    R = CC.PuntoDeChequeo()

            End Select

            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)



        End Function



    End Class
End Namespace