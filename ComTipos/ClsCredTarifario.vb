Option Explicit On

<Serializable()> Public Class ClsCredTarifario
    Public cboCampania As ComTipos.clsDataset
    Public cboAgencia As ComTipos.clsDataset
    Public cboMoneda As ComTipos.clsDataset
    Public cboProducto As ComTipos.clsDataset
    Public cboSubProd As ComTipos.clsDataset
    Public GridLineas As ComTipos.clsDataset
    Public Datos As ComTipos.clsDataset

    Public Sub New()
        cboCampania = New ComTipos.clsDataset
        cboAgencia = New ComTipos.clsDataset
        cboMoneda = New ComTipos.clsDataset
        cboProducto = New ComTipos.clsDataset
        cboSubProd = New ComTipos.clsDataset
        GridLineas = New ComTipos.clsDataset
        Datos = New ComTipos.clsDataset
    End Sub
End Class
