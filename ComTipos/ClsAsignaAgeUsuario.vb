<Serializable()> Public Class ClsAsignaAgeUsuario
    Public oGrid As clsDataset
    Public oAgencias As clsDataset
    Public oCboAutonomias As clsDataset
    Public oCboTipoFiltroTamiz As clsDataset

    Public Sub New()
        oGrid = New clsDataset
        oAgencias = New clsDataset
        oCboAutonomias = New clsDataset
        oCboTipoFiltroTamiz = New clsDataset
    End Sub
End Class
