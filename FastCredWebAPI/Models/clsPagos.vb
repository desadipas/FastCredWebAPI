Imports FastCredWebAPI.FastCredWebAPI

Public Class clsPagos

    Private db As New DBNEGOCIOEntities
    Public Property nCodAge As Nullable(Of Integer)
    Public Property nCodCred As Nullable(Of Integer)
    Public Property nDiasAtraso As Nullable(Of Integer)
    Public Property nMontoPago As Nullable(Of Double)
    Public Property nCapital As Nullable(Of Double)
    Public Property nInteres As Nullable(Of Double)
    Public Property nIgv As Nullable(Of Double)
    Public Property nGasto As Nullable(Of Double)
    Public Property dFecha As Nullable(Of Date)
    Public Property cUser As String
    Public Property cNomOpe As String




End Class
