
Option Explicit On

<Serializable()> Public Class clsCredCliExclusivos
    Public oTipoAsignacion As ComTipos.clsDataset
    Public oCampanias As ComTipos.clsDataset
    Public oNumEntidades As ComTipos.clsDataset
    Public oCalifSBS As ComTipos.clsDataset
    Public oTipOrganizacion As ComTipos.clsDataset
    Public oGeneros As ComTipos.clsDataset
    Public oEstadoCivil As ComTipos.clsDataset
    Public oActPrimaria As ComTipos.clsDataset
    Public oActEconomica As ComTipos.clsDataset
    Public oMeses As ComTipos.clsDataset
    Public oDatos As ComTipos.clsDataset

    Public Sub New()
        oTipoAsignacion = New ComTipos.clsDataset
        oCampanias = New ComTipos.clsDataset
        oNumEntidades = New ComTipos.clsDataset
        oCalifSBS = New ComTipos.clsDataset
        oTipOrganizacion = New ComTipos.clsDataset
        oGeneros = New ComTipos.clsDataset
        oEstadoCivil = New ComTipos.clsDataset
        oActPrimaria = New ComTipos.clsDataset
        oActEconomica = New ComTipos.clsDataset
        oDatos = New ComTipos.clsDataset
        oMeses = New ComTipos.clsDataset

    End Sub
End Class

