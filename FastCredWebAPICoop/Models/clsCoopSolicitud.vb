Public Class clsCoopSolicitud

    Public Property cDoc As String
    Public Property nIDCred As Integer
    Public Property nIDPreSol As Integer
    Public Property nIDEval As Integer
    Public Property nOperacion As Integer
    Public Property nMonto As Double
    Public Property nMoneda As Integer
    Public Property nPeriodo As Integer
    Public Property nNumCuo As Integer
    Public Property nDiaPago As Integer
    Public Property cCodAsesor As String
    Public Property cCodPromotor As String
    Public Property nEstado As Integer
    Public Property nDestino As Integer
    Public Property nFuente As Integer
    Public Property nAdeudado As Integer
    Public Property dFecDesem As Date
    Public Property dFecPrimCuota As Date
    Public Property nTipoTasa As Integer
    Public Property nConvenio As Integer
    Public Property nProducto As Integer
    Public Property nModCredito As Integer

    Public Property nTasaIni As Double
    Public Property nTasaFin As Double
    Public Property nCodTarif As Integer
    Public Property nModDesemb As Integer
    Public Property nTasa As Double
    Public Property nTasaMor As Double
    Public Property cCtaVinc As String
    Public Property nTipoEval As Integer
    Public Property nMotivoOper As Integer
    Public Property cComentario As String
    Public Property nVerifDomNeg As Integer
    Public Property cComVerifDomNeg As String
    Public Property nActividadInterna As Integer
    Public Property nSubProd As Integer

    Public Property oCoopPersona As clsCoopPersona
    Public Property oEvaluacion As clscoopevalpyme
    Public Property oCoopevalunidfam As clsCoopevalunidfam
    Public Property oCoopevaldiario As clsCoopevaldiario
    Public Property oCoopevalgrupal As clsCoopevalgrupal

End Class
