Option Explicit On
<Serializable()> Public Class clsNuevoFlujoCaja
    Public oVentasActuales As clsDataset
    Public oMesTipo As clsDataset ''Mes Alto,Medio,Bajo
    Public oCampañas As clsDataset 'Meses del Año
    Public oSupuestos As clsDataset
    Public oPeriodo As clsDataset
    Public oVentasActualesDetalle As clsDataset
    Public oMesesAltos As clsDataset
    Public oPorcentajeMaximo As clsDataset
    Public oPeriodo1 As clsDataset
    Public oPeriodo2 As clsDataset
    Public oActivoFijo As clsDataset
    Public DatosProceso As clsDataset
    Public oPagosNegocio As clsDataset
    Public oIngresoCapital As clsDataset
    Public oIngresoActivo As clsDataset
    Public oIngresoMixto As clsDataset

    Public oEgresoCapital As clsDataset
    Public oEgresoActivo As clsDataset
    Public oEgresoMixto As clsDataset

    Public oCredAgricola As clsDataset

    Public oCredAnt As clsDataset

    Public oDataCredito As clsDataset

    'Nuevo flujo vvillafranca 19-07-2012
    Public oProyVentas As clsDataset
    Public oMesesSupuestos As clsDataset
    Public oIngresosExtra As clsDataset

    Public oDatosEntradaFC As clsDataset
    Public oListaProyEntradaFC As clsDataset
    Public oFechaRegFC As clsDataset
    Public oListaCampaniaFC As clsDataset
    Public oListaIngreExtraFC As clsDataset
    Public oListaOtrosCreditos As clsDataset



    Public Sub New()
        oVentasActuales = New clsDataset
        oMesTipo = New clsDataset ''Mes Alto,Medio,Bajo
        oCampañas = New clsDataset ''Mes del Año
        oSupuestos = New clsDataset
        oPeriodo = New clsDataset
        oVentasActualesDetalle = New clsDataset
        oMesesAltos = New clsDataset
        oPorcentajeMaximo = New clsDataset
        oPeriodo1 = New clsDataset
        oActivoFijo = New clsDataset
        DatosProceso = New clsDataset
        oPagosNegocio = New clsDataset
        oIngresoCapital = New clsDataset
        oIngresoActivo = New clsDataset
        oIngresoMixto = New clsDataset
        oDataCredito = New clsDataset
        oEgresoCapital = New clsDataset
        oEgresoActivo = New clsDataset
        oEgresoMixto = New clsDataset
        oCredAgricola = New clsDataset
        oPeriodo2 = New clsDataset
        oCredAnt = New clsDataset
        oProyVentas = New clsDataset
        oMesesSupuestos = New clsDataset
        oIngresosExtra = New clsDataset

        oDatosEntradaFC = New clsDataset
        oListaProyEntradaFC = New clsDataset
        oFechaRegFC = New clsDataset
        oListaCampaniaFC = New clsDataset
        oListaIngreExtraFC = New clsDataset
        oListaOtrosCreditos = New clsDataset

    End Sub
End Class

