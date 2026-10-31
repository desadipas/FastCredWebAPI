<Serializable()> Public Class ClsImpCredSol

    Public oDatos As ComTipos.clsDataset
    Public oPersonas As ComTipos.clsDataset
    Public oCalDesem As ComTipos.clsDataset

    Public Sub New()
        oDatos = New ComTipos.clsDataset
        oPersonas = New ComTipos.clsDataset
        oCalDesem = New ComTipos.clsDataset
    End Sub
End Class
