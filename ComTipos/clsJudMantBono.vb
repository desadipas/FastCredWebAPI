Option Explicit On
<Serializable()> Public Class clsJudMantBono
    Public otarifarios As clsDataset
    Public ocabecera As clsDataset
    Public odetalle As clsDataset
    Public onumerotarifarios As clsDataset
    Public olista As clsDataset
    Public Sub New()
        otarifarios = New clsDataset
        ocabecera = New clsDataset
        odetalle = New clsDataset
        onumerotarifarios = New clsDataset
        olista = New clsDataset
    End Sub
End Class
