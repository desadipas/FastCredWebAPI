Option Explicit On
<Serializable()> Public Class clsJudRegVizualizar
    Public odatos As clsDataset
    Public odiralt As clsDataset
    Public otelalt As clsDataset
    Public ocbotelefono As clsDataset
    Public ocbodpto As clsDataset
    Public oregistro As clsDataset
    Public Sub New()
        odatos = New clsDataset
        odiralt = New clsDataset
        otelalt = New clsDataset
        ocbotelefono = New clsDataset
        ocbodpto = New clsDataset
        oregistro = New clsDataset
    End Sub

End Class
