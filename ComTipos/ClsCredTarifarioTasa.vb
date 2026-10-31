Option Explicit On

<Serializable()> Public Class ClsCredTarifarioTasa
    Public oCboAutonomias As clsDataset
    Public oGridDatos As clsDataset

    Public Sub New()
        oCboAutonomias = New clsDataset
        oGridDatos = New clsDataset
    End Sub

End Class
