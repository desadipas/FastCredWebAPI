<Serializable()> Public Class ClsCredVisita
    Public cboDestino As ComTipos.clsDataset
    Public CboConVisita As ComTipos.clsDataset
    Public CboAteCli As ComTipos.clsDataset

    Public Sub New()
        cboDestino = New ComTipos.clsDataset
        CboConVisita = New ComTipos.clsDataset
        CboAteCli = New ComTipos.clsDataset
    End Sub
End Class
