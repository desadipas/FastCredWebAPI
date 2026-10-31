
<Serializable()> Public Class ClsWCargaPersona
    Public Datos As clsDataset
    Public cboTipDocs As clsDataset
    Public cboLocalidad As clsDataset
    Public cboDpto As clsDataset
    Public cboProv As clsDataset
    Public cboDistr As clsDataset
    Public cboCIIU As clsDataset
    Public cboTipoPersona As clsDataset
    Public cboEstadoCiv As clsDataset
    Public cboSexo As clsDataset


    Public Sub New()
        Datos = New clsDataset
        cboTipDocs = New clsDataset
        cboLocalidad = New clsDataset
        cboDpto = New clsDataset
        cboProv = New clsDataset
        cboDistr = New clsDataset
        cboCIIU = New clsDataset
        cboTipoPersona = New clsDataset
        cboEstadoCiv = New clsDataset
        cboSexo = New clsDataset
    End Sub

End Class
