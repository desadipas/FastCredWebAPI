Public Class InterconnectResponse
    Public Property Codigo As Integer?
    Public Property Nombre As String
    Public Property fechaHora As String
    Public Property zonaHoraria As String
    Public Property TieneError As Boolean
    Public Property DetallesError As InterconnectDetalleError
    Public Property Data As InterconnectData
End Class
