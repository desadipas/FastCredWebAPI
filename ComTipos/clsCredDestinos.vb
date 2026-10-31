Option Explicit On
<Serializable()> Public Class clsCredDestinos

    Public oAgencias As clsDataset
    Public oGrid As clsDataset
    Public oGest As clsDataset
    Public oAsesores As clsDataset
    Public oTipoCredito As clsDataset
    Public oDataGrilla As clsDataset
    Public oPlazo As clsDataset
    Public oDias As clsDataset
    Public oEquivale As clsDataset
    Public oExcel As clsDataset
    Public oRecuperaExcel As clsDataset
    Public oValidaExcel As clsDataset

    Public Sub New()
        oGrid = New clsDataset
        oGest = New clsDataset
        oAgencias = New clsDataset
        oAsesores = New clsDataset
        oTipoCredito = New clsDataset
        oDataGrilla = New clsDataset
        oDias = New clsDataset
        oEquivale = New clsDataset
        oExcel = New clsDataset
        oRecuperaExcel = New clsDataset
        oValidaExcel = New clsDataset
        oPlazo = New clsDataset
    End Sub

End Class
