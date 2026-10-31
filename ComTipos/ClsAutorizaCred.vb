Option Explicit On

<Serializable()> Public Class ClsAutorizaCred

    Public oDatos As ComTipos.clsDataset
    Public oGridAutor As ComTipos.clsDataset

    Public Sub New()
        oDatos = New ComTipos.clsDataset
        oGridAutor = New ComTipos.clsDataset
    End Sub

End Class
