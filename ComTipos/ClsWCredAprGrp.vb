<Serializable()> Public Class ClsWCredAprGrp

    Public oIntegrantes As ComTipos.clsDataset
    Public oDatos As ComTipos.clsDataset

    Public Sub New()
        oIntegrantes = New ComTipos.clsDataset
        oDatos = New ComTipos.clsDataset
    End Sub
End Class
