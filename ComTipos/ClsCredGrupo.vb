Option Explicit On

<Serializable()> Public Class ClsCredGrupo
    Public cboTipoGrupo As ComTipos.clsDataset
    Public cboCargosEnGrupo As ComTipos.clsDataset
    Public DatosGrupo As ComTipos.clsDataset
    Public DatosInteg As ComTipos.clsDataset

    Public Sub New()
        cboTipoGrupo = New ComTipos.clsDataset
        cboCargosEnGrupo = New ComTipos.clsDataset
        DatosGrupo = New ComTipos.clsDataset
        DatosInteg = New ComTipos.clsDataset

    End Sub
End Class
