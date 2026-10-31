Option Explicit On
<Serializable()> Public Class clsJudBonoGestor
    Public oagencias As clsDataset
    Public ogestores As clsDataset
    Public otarifario As clsDataset
    Public otarifarioactual As clsDataset
    Public onumerotarifarios As clsDataset

    Public Sub New()
        ogestores = New clsDataset
        otarifario = New clsDataset
        otarifarioactual = New clsDataset
        onumerotarifarios = New clsDataset
        oagencias = New clsDataset
    End Sub
End Class
