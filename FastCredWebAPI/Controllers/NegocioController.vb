Imports System.Net
Imports System.Web.Http

Imports System.Web.Http.Description
Imports FastCredWebAPI.FastCredWebAPI

Namespace Controllers
    Public Class NegocioController
        Inherits ApiController

        Private db As DBNEGOCIOEntities

        <ResponseType(GetType(inNegocio))>
        Function PostNegocio(ByVal CC As inNegocio) As IHttpActionResult

            Dim R As outNegocio

            If Not ModelState.IsValid Then
                Return BadRequest(ModelState)
            End If

            R = New outNegocio()

            Select Case CC.nSolicitud
                Case 0 'nSolicitud = 0 -> Recupera Todas las politicas de negocio
                    R = CC.RecuperaPoliticas()
                Case 1 'nSolicitud = 1 -> Recupera Creditos a Aprobar
                    R = CC.RecuperaSolicitud()
                Case 2 'nSolicitud = 2 -> Recupera persona
                    R = CC.RecuperaSolicitud()
                Case 3 'nSolicitud = 3 -> Registra persona
                    R = CC.RegistraPersona()
                Case 4 'nSolicitud = 4 -> Actualiza Datos Control
                    R = CC.RegistraDatosControl()
                Case 5 'nSolicitud = 5 -> recupera ingresos y egresos
                    R = CC.RecuperaIngresosEgresos()
                Case 6 'nSolicitud = 6 -> recupera listado de grupos
                    R = CC.RecuperaListaGruposComunales()
                Case 7 'nSolicitud = 7 -> recupera Detalle de Grupo
                    R = CC.RecuperaDetalleGrupoComunal()
                Case 8 'nSolicitud = 8 -> registra grupo comunal
                    R = CC.RegistraGrupoComunal()
                Case 9 'nSolicitud = 9 -> registra socio
                    R = CC.RegistraGrupoComunalSocio()
                Case 10 'nSolicitud = 10 -> buscar persona para socio de grupo comunal
                    R = CC.BuscarGrupoComunalSocio()
                Case 11 'nSolicitud = 11 -> Recupera datos de Socio
                    R = CC.RecuperaDatosSocio()
                Case 12 'nSolicitud = 12 -> Recupera datos de Grupo
                    R = CC.RecuperaDatosGrupoComunal()
                Case 13 'nSolicitud = 13 -> Elimina Socio
                    R = CC.EliminaSocioGrupoComunal()
                Case 14 'nSolicitud = 14 -> recupera listado de grupos para aprobar
                    R = CC.RecuperaListaGruposComunalesParaAprobar()
                Case 15 'nSolicitud = 15 -> recupera listado de prestamos de grupo para aprobar
                    R = CC.RecuperaDetalleGrupoComunalParaAprobar()
                Case 16 'nSolicitud = 16 -> recupera listado de prestamos de grupo para Desembolsar
                    R = CC.RecuperaDetalleGrupoComunalParaDesembolsar()
                Case 17 'nSolicitud = 17 -> recupera listado de pagos de grupo 
                    R = CC.RecuperaListaGruposComunalesParaPagar()
                Case 18 'nSolicitud = 18 -> recupera Record  de Cliente
                    R = CC.RecuperaRecordCliente()
                Case 19 'nSolicitud = 19 -> recupera lista pagos
                    R = CC.RecuperaListaPagos()
                Case 20 'Logueo e Importacion de Tablas Base
                    R = CC.RecuperaTablasIniciales()
                Case 21 'Registra GPS
                    R = CC.RegistraGPS()
                Case 22 'Valida Existencia de Usuario
                    R = CC.ValidaExisteUsuario()
                Case 23 'Valida Existencia de Usuario
                    R = CC.RecuperaFotos()
                Case 24 'Transaccion de Pago
                    R = CC.RegistrarPago()
                Case 25 'Importar datos basicos
                    R = CC.RecuperaDatosBasicos()
                Case 26 'Importar datos persona
                    R = CC.RecuperaPersona()
                Case 27 'Regsitra Solicitud de Pyme
                    R = CC.RegistraSolcitudPyme()
                Case 28 'Regsitra Solicitud de Consumo
                    R = CC.RegistraSolcitudConsumo()
                Case 29 'Regsitra Solicitud de Empenio
                    R = CC.RegistraSolcitudEmpenio()
                Case 30 'Regsitra Verificacion
                    R = CC.RegistraVerificacion()
                Case 31 'Regsitra Desembolso
                    R = CC.Movil2_RegistraDesembolsoPorProducto()
                Case 32 'Cierre Operaciones
                    R = CC.CierreOperaciones()
                Case 33 'Cierre Operaciones
                    R = CC.RegistraProspeccion()
                Case 34 'Cierre Operaciones
                    R = CC.EliminaInactivo()
                Case 35 'Recupera Lista Para Asignar
                    R = CC.RecuperaListaXAsignar()
                Case 36 'Registra Asignacion de Asesor y Desembolsador
                    R = CC.RegistraAsignacionAsesor()
                Case 37 'RecuperaLista Por Abonar
                    R = CC.RecuperaListaXAbonar()
                Case 38 'Registra Lista Por Abonar
                    R = CC.RegistraAbonosDesembolsos()
                Case 39 'Registra Lista Por Abonar
                    R = CC.RecuperaListaDesembolsadores()
                Case 40 'Recupera Datos de Credito
                    R = CC.RecuperaDatosDeCredito()
                Case 41 'Recupera cuota en linea
                    R = CC.RecuperaCuotaONLINE()
                Case 42 'Recupera cuota en linea
                    R = CC.RecuperaListaAgencias()
                Case 43 'Recupera cuota en linea Categorias Micro
                    R = CC.RecuperaCuotaONLINE_Categorias()
                Case 44
                    R = CC.ActualizaIMEI()
                Case 45
                    R = CC.ValidaUsuario()
                Case 46
                    R = CC.RegistraVoucher()
                Case 47
                    R = CC.RecuperaResumenCobranza()
                Case 48
                    'R = CC.RecuperaBonosXCobranza()
                    R = CC.RecuperaBonosXCobranzaNueva()
                Case 49
                    R = CC.GuardarFotoFirmaContrato()
                Case 50
                    R = CC.RecuperaFotoFirmaContrato()
                Case 51
                    R = CC.RecuperaDatosEnvioSMS()
                Case 52
                    R = CC.RegistraEnvioSMS()
                Case 53
                    R = CC.RecuperaDatosEnvioSMS_Cliente()
                Case 54
                    R = CC.RegistraEnvioSMSCliente()
                Case 55
                    R = CC.RegistraMarcadoAsistencia()
                Case 56
                    R = CC.RegistraExcepcion()
                Case 57
                    R = CC.RegistraCambioLinea()
                Case 58
                    R = CC.RecuperaDatosCambioLinea()
                Case 59
                    R = CC.RecuperaDatosBandejaAprobacion()
                Case 60
                    R = CC.RecuperaDatosAprobacion()
                Case 61
                    R = CC.MontoCuotaPorProd()
                Case 62
                    R = CC.RegistraAprobacion()
                Case 63
                    R = CC.RegistraRechazoCredito()
                Case 64
                    R = CC.RecuperaTamañoNegocio()
                Case 65
                    R = CC.RecuperaFotosVerificacion()
                Case 66
                    R = CC.RegistraMotivoNOPago()
                Case 67
                    R = CC.RecuperaBonosXColocacion()
                Case 68
                    R = CC.PrefiltroXProd()
                Case 69
                    R = CC.RecuperaBandejaVerificacion()
                Case 70
                    R = CC.ValidaRol()
                Case 71
                    R = CC.RecuperaImagenesNegocio()
                Case 72
                    R = CC.RecuperaEvaluacion()
                Case 73
                    R = CC.RegistraVerificacionNueva()
                Case 74
                    R = CC.RecuperaDesembolsosDiaListaXProducto()
                Case 75
                    R = CC.RecuperaPagosListaItem()
                Case 76
                    R = CC.RecuperaPagoClientePorProducto()
                Case 77
                    R = CC.RegistraMotivoNODesembolso()
                Case 78
                    R = CC.RecuperaBandejaAprobacion()
                Case 79
                    R = CC.RecuperaListaCreditosPorDesembolsar()
                Case 80
                    R = CC.RecuperaDatosBonoSupvervisor()
                Case 81
                    R = CC.RecuperaPrecalifiNew()
                Case 82
                    R = CC.RegistraEnvioAVerificacion()
                Case 83
                    R = CC.RecuperaBandejaVerificacionParaActualizar()
                Case 84
                    R = CC.RecuperaDataPrecalifNueva
                Case 85
                    R = CC.RecuperaInactivos()
                Case 86
                    R = CC.RegistraSolcitudPymeExcepcion
                Case 87
                    R = CC.RecuperaDireccionesActuales()
                Case 88
                    R = CC.ActualzaDireccionesActuales()
                Case 89
                    R = CC.ActualzaImeiAnulaCierre()
                Case 90
                    R = CC.RecuperaCelularesActuales()
                Case 91
                    R = CC.ActualzaCelularCliente()
                Case 92
                    R = CC.RecuperaDatosCambioTasa()
                Case 93
                    R = CC.RegistraaDatosCambioTasa()
                Case 94
                    R = CC.CambioClave()
                Case 95
                    R = CC.RecuperaScore()
                Case 96
                    R = CC.RegistraEntregaEfectivo()
                Case 97
                    R = CC.CargaUsuarios()
                Case 98
                    R = CC.RecuperaEfectivoDisponible()
                Case 99
                    R = CC.RecuperaProyeccionAmortizacion()
                Case 100
                    R = CC.RegistrarPagoGrupalLote()
                Case 101
                    R = CC.RecuperaConsultaSolicitudes()
                Case 102
                    R = CC.RegistrarTasacion()
                Case 103
                    R = CC.RecuperaRespuestaTasacion()
                Case 104
                    R = CC.RegistrarRespuestaTasacion()
                Case 105
                    R = CC.RegistrarExperian()
                Case 106
                    R = CC.RecuperarExperian()
                Case 107
                    R = CC.RecuperaConsultaBonos()
                Case 108
                    R = CC.RecuperaConsultaClienteTasaPreferencial()
                Case 109
                    R = CC.RegistrarSolicitudTasaPreferencial()
                Case 110
                    R = CC.RecuperaConsultaClienteTasaPreferencialaAprobar()
                Case 111
                    R = CC.RegistrarAprobacionTasaPreferencial()
                Case 112
                    R = CC.RecuperaGruposAprobacion()
                Case 113
                    R = CC.RecuperaBandejaAprobacionGrupo()
                Case 114
                    R = CC.ValidaRegistroVoucher()
                Case 115
                    R = CC.RecuperaGruposDesembolso()
                Case 116
                    R = CC.RecuperaCreditosXAsignarAsesorGrupal()
                Case 117
                    R = CC.RecuperaCreditosXAbonarADesembGrupal()
                Case 118
                    R = CC.RecuperaConsultaSolicitudesxCred()
                Case 127
                    R = CC.ActualizarVersionAppUsuario()
                Case 130
                    R = CC.ConsultaDocumentoExperian()
                Case 131
                    R = CC.RegistrarHuellaDigital()
                Case 132
                    R = CC.ConsultaDocumentoGOLDENCASH()
                Case 133
                    R = CC.RecuperaCuotaSimuladorONLINE_Categorias()
                Case 134
                    R = CC.RecuperaDatosOperacionesEspecialesAdm()
                Case 135
                    R = CC.EjecutaOperacionesEspecialesAdmAmpliacion()
                Case 136
                    R = CC.EjecutaOperacionesEspecialesAdmCambioProd()
                Case 137
                    R = CC.EjecutaOperacionesEspecialesAdmTasa()
                Case 138
                    R = CC.LinkKeynua()
                Case 139
                    R = CC.RecuperaCreditosOpeEspec()
                Case 140
                    R = CC.RecuperaDatosAprobacionIndividual()
                Case 141
                    R = CC.RecuperaAbonosPorAprobarPorAgencia()
                Case 142
                    R = CC.RecuperaCanceladosHoy()
                Case 143
                    R = CC.RecuperaExperianData()

            End Select

            'Validar Token

            Return CreatedAtRoute("DefaultApi", New With {.id = 1}, R)

        End Function



    End Class
End Namespace