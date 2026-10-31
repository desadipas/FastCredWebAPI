<Serializable()> Public Class ClsParamGestor

    Public cboGestores As ComTipos.clsDataset
    Public cboAgencias As ComTipos.clsDataset
    Public GrdAnalistasSINGestor As ComTipos.clsDataset
    Public GrdAnalistasConGestor As ComTipos.clsDataset

    Public Sub New()
        cboGestores = New ComTipos.clsDataset
        GrdAnalistasSINGestor = New ComTipos.clsDataset
        GrdAnalistasConGestor = New ComTipos.clsDataset
        cboAgencias = New ComTipos.clsDataset
    End Sub

End Class
