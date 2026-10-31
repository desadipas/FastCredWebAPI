Option Explicit On

<Serializable()> Public Class clsJudAsignaTarifarioGestor
    Public ogestores As clsDataset
    Public otarifario As clsDataset
    Public otarifarioactual As clsDataset
    Public Sub New()
        ogestores = New clsDataset
        otarifario = New clsDataset
        otarifarioactual = New clsDataset
    End Sub
End Class
