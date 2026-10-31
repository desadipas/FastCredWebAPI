<Serializable()> Public Class ClsWRegistroCredConsumo
    Public Datos As ComTipos.clsDataset
    Public CboCalificacion As ComTipos.clsDataset
    Public Deudas As ComTipos.clsDataset
    Public CboClasificacion As ComTipos.clsDataset
    Public cboTipoGrati As ComTipos.clsDataset 'ARCV 10-09-2010

    Public Sub New()
        Datos = New ComTipos.clsDataset
        CboCalificacion = New ComTipos.clsDataset
        Deudas = New ComTipos.clsDataset
        CboClasificacion = New ComTipos.clsDataset
        cboTipoGrati = New ComTipos.clsDataset 'ARCV 10-09-2010
    End Sub
End Class
