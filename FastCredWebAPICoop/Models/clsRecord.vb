Imports System.Data.Entity
Imports FastCredWebAPI.FastCredWebAPI

Public Class clsRecord

    Private db As New DBNEGOCIOEntities

    Public Property cDoc As String
    Public Property nCodAge As Integer
    Public Property ncodcred As Integer
    Public Property nCuotas As Integer
    Public Property nDiasAtraso As Integer
    Public Property cEstado As String
    Public Property cAsesor As String
    Public Property nNumPagos As Nullable(Of Integer)
    Public Property nNumPagosConAtraso As Nullable(Of Integer)
    Public Property nDiasAtrasoPromedio As Integer
    Public Property nPrestamo As Decimal
    Public Property nMontoCuota As Nullable(Of Decimal)

    Public Property Pagos As IQueryable(Of clsPagos)
    Public Property Calendario As IQueryable(Of clsCalendario)

    Public Sub CargaDatos()

        Dim vPagos = From P In db.CredKardex
                     Join O In db.Operaciones On O.nCodOpe Equals P.nCodOpe
                     Where P.nAgeCred = nCodAge And P.nCodCred = ncodcred And (P.nCodOpe = 2 Or P.nCodOpe = 9)
                     Select New clsPagos With {.cNomOpe = O.cCodOpe, .cUser = P.cUser, .dFecha = P.dfecha, .nCapital = P.nCapital, .nInteres = P.nIntComp, .nCodAge = P.nAgeCred, .nCodCred = P.nCodCred, .nDiasAtraso = P.nDiasAtraso, .nGasto = P.nGasto, .nIgv = P.nIGV, .nMontoPago = P.nMonto}

        Pagos = vPagos

        Dim vCal = From C In db.CredCalendario
                   Where C.nCodAge = nCodAge And C.nCodCred = ncodcred
                   Select New clsCalendario With {.cEstado = If(C.nEstado = 0, "PENDIENTE", "CANCELADO"), .dFechaVenc = C.dFecVenc, .dFechaPago = C.dFecPago, .nCodAge = C.nCodAge, .nCodCred = C.nCodCred, .nDiasAtraso = DbFunctions.DiffDays(C.dFecVenc, C.dFecPago), .nMontoPago = C.nCapital + C.nIntComp + C.nIgv, .nMontoPagoPendiente = (C.nCapital - C.nCapPag) + (C.nIntComp - C.nIntPag) + (C.nIgv - C.nIgvPag), .nNroCuota = C.nNroCuota}

        Calendario = vCal



    End Sub

End Class
