Option Explicit On

<Serializable()> Public Class clsInstituciones
    Public oGrid As clsDataset

    Public Sub New()
        oGrid = New clsDataset
    End Sub
End Class
