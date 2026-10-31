<Serializable()> Public Class clsRegistroTeleCredito
    Public CboBancos As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset
    Public CboAgencias As ComTipos.clsDataset
    Public oValidacion As ComTipos.clsDataset
    Public oExport_Cab As ComTipos.clsDataset
    Public oExport_Det As ComTipos.clsDataset

    Public Sub New()
        CboBancos = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
        CboAgencias = New ComTipos.clsDataset
        oValidacion = New ComTipos.clsDataset
        oExport_Cab = New ComTipos.clsDataset
        oExport_Det = New ComTipos.clsDataset
    End Sub
End Class
