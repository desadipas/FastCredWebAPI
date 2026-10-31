Option Explicit On

<Serializable()> Public Class ClsAutonomia
    Public oNiveles As ComTipos.clsDataset
    Public oNivelesUsu As ComTipos.clsDataset
    Public oRangos As ComTipos.clsDataset
    Public oRangosAutori As ComTipos.clsDataset
    Public oAgencias As ComTipos.clsDataset
    Public oProd As ComTipos.clsDataset
    Public oSubProd As ComTipos.clsDataset
    Public oMoneda As ComTipos.clsDataset

    Public Sub New()
        oNiveles = New ComTipos.clsDataset
        oNivelesUsu = New ComTipos.clsDataset
        oRangos = New ComTipos.clsDataset
        oRangosAutori = New ComTipos.clsDataset
        oAgencias = New ComTipos.clsDataset
        oProd = New ComTipos.clsDataset
        oSubProd = New ComTipos.clsDataset
        oMoneda = New ComTipos.clsDataset
    End Sub
End Class
