<Serializable()> Public Class ClsCreConsulta

    Public Data As ComTipos.clsDataset
    Public CuotaPag As ComTipos.clsDataset
    Public CuotaPend As ComTipos.clsDataset
    Public Kardex As ComTipos.clsDataset
    Public Garantias As ComTipos.clsDataset
    Public DataDeudCal As ComTipos.clsDataset
    Public DataDeudaFecha As ComTipos.clsDataset
    Public DataRelaPers As ComTipos.clsDataset
    Public DatosCalendario As ComTipos.clsDataset
    Public CalendDesembolsos As ComTipos.clsDataset


    Public Sub New()
        Data = New ComTipos.clsDataset
        CuotaPag = New ComTipos.clsDataset
        CuotaPend = New ComTipos.clsDataset
        Kardex = New ComTipos.clsDataset
        Garantias = New ComTipos.clsDataset
        DataDeudCal = New ComTipos.clsDataset
        DataDeudaFecha = New ComTipos.clsDataset
        DataRelaPers = New ComTipos.clsDataset
        DatosCalendario = New ComTipos.clsDataset
        CalendDesembolsos = New ComTipos.clsDataset

    End Sub

End Class
