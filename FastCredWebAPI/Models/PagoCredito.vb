Imports FastCredWebAPI.FastCredWebAPI

Public Class dPagoCredito

    Private db As New DBNEGOCIOEntities

    Public Property cDoc As String
    Public Property cNomcliente As String
    Public Property dFechaVenc As Date
    Public Property nPrestamo As Double
    Public Property nCodAge As Integer
    Public Property nCodCred As Integer
    Public Property nMonto As Double
    Public Property cMovil As String
    Public Property cZonaPersona As String
    Public Property cDireccionDom As String
    Public Property cZonaNegocio As String
    Public Property cDirecNeg As String
    Public Property nNroProxCuota As String
    Public Property nDeudaAtrasada As Double
    Public Property nDiasAtraso As Double
    Public Property UltFecCob As Date
    Public Property cMovil2 As String
    Public Property cReferencia As String
    Public Property cRazonSocial As String
    Public Property nMontoCancelaCuota As Double
    Public Property nMontoCancelaCredito As Double
    Public Property sAmpliacion As String
    Public Property nLatitud As Double
    Public Property nLongitud As Double
    Public Property nPagoMinio As Double
    Public Property bPagadoHoy As Boolean
    Public Property nNumPagosDelDia As Integer
    Public Property cCodUsuCob As String
    Public Property nProd As Integer
    Public Property nCiclos As Integer
    Public Property cCategoria As String

    Public Property nCuotasTotales As Integer

    Public Property nCuotasPagadas As Integer

    Public Property nCuotasPendientes As Integer
    Public Property cCodUsuAnalista As String
    Public Property nOrdenRuta As Integer
    Public Property nPagopromedio As Double
    Public Property nIDGrupo As Integer
    Public Property Calendario As IQueryable(Of dCalendarioPago)

    Public Sub CargaCalendario()


        Dim vCalendario = From C In db.CredCalendario
                          Where C.nCodAge = nCodAge And C.nCodCred = nCodCred
                          Select New dCalendarioPago With {.nCapital = C.nCapital, .nCapPag = C.nCapPag,
                            .dFechaPago = C.dFecPago, .dFechaVenc = C.dFecVenc, .nCodAge = C.nCodAge,
                            .nCodCred = C.nCodCred, .nEstado = C.nEstado, .nIgv = C.nIgv, .nIgvPag = C.nIgvPag,
                            .nInteres = C.nIntComp, .nIntPag = C.nIntPag, .nNroCuota = C.nNroCuota
                            }

        Calendario = vCalendario



    End Sub

End Class
