Imports System.Data
Imports System.Data.Entity
Imports System.Data.Entity.Infrastructure

Imports System
Imports System.Collections.Generic

Public Class outNegocio

    Public Property nResultado As Integer
    Public Property sRespuesta As String
    Public Property nTasaComp As Double
    Public Property cDireccionDomicilio As String
    Public Property cDireccionNegocio As String

    Public Property cMovil1 As String

    Public Property ListaPrestamoPorAprobar As IQueryable(Of vPrestamoPorAprobar)

    Public Property ValoresPantalla As IQueryable(Of SistFormularios)
    Public Property ReglasdeNegocio As IQueryable(Of sistNegocioReglas)
    Public Property ReglasdeNegocioParam As IQueryable(Of sistNegocioReglasParam)

    Public Property ReglasdeNegocioDependencias As IQueryable(Of sistNegocioReglasDependencias)

    Public Property ListaFormulariosRoles As IQueryable(Of sistFormulariosRoles)

    Public Property CuadreCajaUsuarios As IQueryable(Of CuadreCajaUsuario)

    Public Property ListaAgencias As IQueryable(Of Agencias)

    Public Property ListaGrupos As IQueryable(Of clsGrupoComunal)
    Public Property oGrupoComunal As clsGrupoComunal
    Public Property oSocio As clsGrupoComunalSocio

    Public Property oRecord As clsPosicionCliente

    Public Property ListaPagosCredito As IQueryable(Of dPagoCredito)

    Public Property oUsuario As dUsuario
    Public Property oUsuarioroles As IQueryable(Of dUsuarioroles)
    Public Property oCatalogocodigos As IQueryable(Of dCatalogocodigos)
    Public Property nExisteUsuario As Integer
    Public Property oFotosCredito As IQueryable(Of EvaluacionFotos)

    Public Property TranPago As dTransacPago

    'Public Property oPersonas As IQueryable(Of dPersona)
    Public Property oPersonas As dPersona()
    Public Property oNegocios As IQueryable(Of dNegocio)

    Public Property oZonas As IQueryable(Of dZonas)

    Public Property oTarifarioAgencia As IQueryable(Of TarifarioAgencia)

    Public Property oTarifarioGrupal As IQueryable(Of TarifarioGrupal)

    Public Property oTarifarioMicro As IQueryable(Of TarifarioMicro)
    Public Property oTarifarioMicroSemanal As IQueryable(Of TarifarioMicroSemanal)
    Public Property oTarifarioMicroJuniorSemanal As IQueryable(Of TarifarioMicroJuniorSemanal)

    Public Property oListaClientesPorVerificar As IQueryable(Of vListaClientesPorVerificar)
    Public Property oListaClientesPorVerificarFotos As IQueryable(Of vListaClientesPorVerificarFotos)

    Public Property oListaClientesPorDesembolsar As IQueryable(Of vListaCreditosPorDesembolsar)
    Public Property oListaClientesInactivos As IQueryable(Of ClientesInactivos)
    Public Property oListaArticuloFamilia As IQueryable(Of ArticuloFamilia)
    Public Property oListaArticuloMarcas As IQueryable(Of ArticuloMarcas)
    Public Property oListaArticuloModelo As IQueryable(Of ArticuloModelo)
    Public Property ListaCreditosXAsignar As dAsignaAsesor()

    Public Property ListaAbonos As dAbonoDesemb()
    Public Property ListaDesembolsadoresAbonar As String()

    Public Property oResumenCobranza As clsResumenCobranza
    Public Property oBonosXCobranza As clsBonoXCobranza()

    Public Property oFirmaContratoData As clsfirmacontratodata
    Public Property oSMS As clsSMS
    Public Property oSMSCliente As clsSMSCliente
    Public Property oDatosActLinea As clsDatosActLinea
    Public Property oCreditosAAprobar As clsCreditosAAprobar()
    Public Property oDatosAprobacion As clsDatosAprobacion
    Public Property oCatalogocodigos_2 As clsCatalogocodigos()
    Public Property oFotosVerificacion As String()
    Public Property oAmpliaciones As clsItemAmpliacion()
    Public Property oProyecamortizacion As clsproyecamortizacion()
    Public Property oBonosXColocacion As clsBonoXColocacion()
    Public Property oBandejaVerificacion As clsVerificacionCliente()
    Public Property oFotosEvaluacion As String()
    Public Property oEvalNegocio As clsEvalNegocio
    Public Property oDesembolsos As clsDesembolsos()
    Public Property oDesembolsoItem As clsDesembolsoItem
    Public Property oClientePorProducto As clsClientePorProducto
    Public Property oAprobaciones As clsAprobacionNueva()
    Public Property oUsuariosEquipo As clsUsuariosEquipo()
    Public Property oBandejaDesembolsoNuevo As clsDesembolsoNuevo()
    Public Property oDatosBonoSupvervisor As clsDatosBonoSupvervisor
    Public Property oPreCalif As clsPrecalif

    Public Property nPorcComision As Double
    Public Property nIncreMora As Double
    Public Property nMoraAnt As Double
    Public Property nMorahoy As Double

    Public Property nSaldoFinMes As Double
    Public Property nSaldoHoy As Double
    Public Property nSaldoVencidoFinMes As Double
    Public Property nSaldoVencidoHoy As Double

    Public Property cCliente As String
    Public Property nCodCred As Long

    Public Property oScore As clsScore

    Public Property oUsuarios As String()

    Public Property nDisponible As Double
    Public Property oProductos As clsProductos()
    Public Property oPagoGrupalCabecera As clsPagogrupalcabecera()
    Public Property oPagoGrupalCabeceraDetalle As clsPagogrupalcabeceraDetalle()
    Public Property oConsultaSolicitudes As clsConsultaSolicitudes()
    Public Property oRespuestasTasacion As clsTasacionCliente()
    Public Property oConsultaBonos As clsBono()
    Public Property oTasaPreferencial As clsTasaPreferencial()
    Public Property oAsigAbonAsesorGrupal As clsAsigAbonAsesorGrupal()
    Public Property oVarnegocio As IQueryable(Of eVarnegocio)
    Public Property experianBasic As ResponseExperianBasic
    Public Property equifaxInformacion As ResponseEquifaxBasic

    Public Property oNegociosRubros As IQueryable(Of Rubro)
    Public Property oDeudasEntidades As IQueryable(Of dEntidad)

    Public Property oCredKeynua As clsCredKeynua()
    Public Property oBandejaCancelados As clsCanceladosHoy()
    Public Property oBandejaExperian As clsExperian()
End Class
