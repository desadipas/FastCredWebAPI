<Serializable()> Public Class ClsCapAhoCtaDatos

    Public oDatosCta As ComTipos.clsDataset
    Public oCtaPers As ComTipos.clsDataset
    Public oHistBloq As ComTipos.clsDataset


    Public Sub New()
        oDatosCta = New ComTipos.clsDataset
        oCtaPers = New ComTipos.clsDataset
        oHistBloq = New ComTipos.clsDataset

    End Sub

End Class
