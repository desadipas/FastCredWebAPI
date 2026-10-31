<Serializable()> Public Class ClsPersDireccion

    Public oTipoDir1 As ComTipos.clsDataset
    Public oTipoDir2 As ComTipos.clsDataset
    Public oTipoDir3 As ComTipos.clsDataset
    Public oTipoDir4 As ComTipos.clsDataset
    Public oDatosDir As ComTipos.clsDataset
    Public oLocal As ComTipos.clsDataset



    Public Sub New()
        oTipoDir1 = New ComTipos.clsDataset
        oTipoDir2 = New ComTipos.clsDataset
        oTipoDir3 = New ComTipos.clsDataset
        oTipoDir4 = New ComTipos.clsDataset
        oDatosDir = New ComTipos.clsDataset
        oLocal = New ComTipos.clsDataset
    End Sub
End Class
