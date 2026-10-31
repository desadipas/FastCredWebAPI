
<Serializable()> Public Class ClsCredPlanilla
    Public cboAnalista As ComTipos.clsDataset
    Public cboAgencia As ComTipos.clsDataset
    Public cboRecaudador As ComTipos.clsDataset
    Public cboRuta As ComTipos.clsDataset
    Public cboProducto As ComTipos.clsDataset
    Public cboSubProd As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset

    Public Sub New()
        cboAnalista = New ComTipos.clsDataset
        cboAgencia = New ComTipos.clsDataset
        cboRecaudador = New ComTipos.clsDataset
        cboProducto = New ComTipos.clsDataset
        cboRuta = New ComTipos.clsDataset
        cboSubProd = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
    End Sub


End Class
