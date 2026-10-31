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
                    R = CC.RegistraComercio()
                Case 1 'nSolicitud = 1 -> Recupera Busqueda de Comercios
                    R = CC.RecuperaBusquedaComercios()
                Case 2 'nSolicitud = 2 -> Recupera Productos
                    R = CC.RecuperaProductos()
                Case 3 'nSolicitud = 3 -> Registra Orden de Pedidos
                    R = CC.RegistraOrden()
                Case 4 'nSolicitud = 4 -> Recupera Ordenes
                    R = CC.RecuperaOrdenes()
                Case 5 'nSolicitud = 5 -> Registrar Cobranza de Pedido
                    R = CC.RegistraCobroPedido()
                Case 6 'nSolicitud = 6 -> Rcupera Cobros
                    R = CC.RecuperaCobros()
                Case 7 'nSolicitud = 7 -> Rcupera Cobro de Comercio
                    R = CC.RecuperaCobroComercio()
                Case 8 'nSolicitud = 8 -> Registra entrega Pedido
                    R = CC.RegistraEntregaPedido()
                Case 9 'nSolicitud = 9 -> Valida Clave
                    R = CC.ValidaClave()
                Case 10 'nSolicitud = 10
                    R = CC.RegistraUbicacion()
                Case 11 'nSolicitud = 11
                    R = CC.ValidaZona()
                Case 12 'nSolicitud = 12
                    R = CC.RegistraNogestion()
                Case 13
                    R = CC.RecuperaCatalogo()
                Case 14
                    R = CC.ValidaSiExisteOrden()
                Case 15
                    R = CC.EliminaOrden()
                Case 16
                    R = CC.RecuperaBusquedaSoloComercio()
                Case 17
                    R = CC.RecuperaPagoParaAnulacion()
                Case 18
                    R = CC.RegistraAnulaPago()
                Case 19
                    R = CC.ValidaRol()

            End Select


            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function



    End Class
End Namespace