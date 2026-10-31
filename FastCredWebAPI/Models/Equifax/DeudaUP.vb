Public Class DeudaUP
    Public Property id_DeudasUltimoPeriodo_Deudas As Integer 'identificador
    Public Property Entidad As String
    Public Property SistemaFinanciero As String
    Public Property CalificacionMesActual As String
    Public Property MontoSubTotalMesActual As Double?
    Public Property CalificacionMesAnterior As String
    Public Property MontoSubTotalMesAnterior As Double?
    Public Property CalificacionAnioAnterior As String
    Public Property MontoSubTotalAnioAnterior As Double?
    Public Property Productos As Producto() 'List(Of Producto)
End Class
