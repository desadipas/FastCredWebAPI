<Serializable()> Public Class ClsWRegistroCredDiario
    Public Datos As ComTipos.clsDataset
    Public Deudas As ComTipos.clsDataset
    Public CboClasificacion As ComTipos.clsDataset

    Public Sub New()
        Datos = New ComTipos.clsDataset
        Deudas = New ComTipos.clsDataset
        CboClasificacion = New ComTipos.clsDataset
    End Sub
End Class
