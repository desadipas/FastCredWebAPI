Option Explicit On
<Serializable()> Public Class ClsCredCondicion
    Public oagencias As clsDataset
    Public ogrid As clsDataset
    Public ogest As clsDataset
    Public oasesores As clsDataset
    Public otipocredito As clsDataset
    Public odatagrilla As clsDataset
    Public oplazo As clsDataset
    Public odias As clsDataset
    Public oequivale As clsDataset
    Public oexcel As clsDataset
    Public orecuperaexcel As clsDataset
    Public ovalidaexcel As clsDataset
    Public Sub New()
        ogrid = New clsDataset
        ogest = New clsDataset
        oagencias = New clsDataset
        oasesores = New clsDataset
        otipocredito = New clsDataset
        odatagrilla = New clsDataset
        odias = New clsDataset
        oequivale = New clsDataset
        oexcel = New clsDataset
        orecuperaexcel = New clsDataset
        ovalidaexcel = New clsDataset
        oplazo = New clsDataset
    End Sub
End Class
