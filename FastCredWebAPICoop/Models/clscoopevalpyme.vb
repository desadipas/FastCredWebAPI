Public Class clscoopevalpyme

    Public Property nIDEval As Long
    Public Property nCodPers As Long
    Public Property cDoc As String
    Public Property dFecRegistro As Date
    Public Property dFecVenc As Date
    Public Property nTipCam As Double
    Public Property dFecEEFF As Date
    Public Property nTipoPropiedad As Integer
    Public Property nTiempoUsoLoc As Integer
    Public Property nTiempoExpNeg As Integer
    Public Property nCalificacionExterna As Integer
    Public Property nHorasDedic As Integer
    Public Property nHorasAtencion As Integer
    Public Property nNumColab As Integer

    Public Property nTipoFuente As Integer
    Public Property cNomNegocio As String
    Public Property cDescripcion As String
    Public Property cDireccion As String
    Public Property cTelefono As String
    Public Property cCodZona As String
    Public Property nLatitud As Double
    Public Property nLongitud As Double

    Public Property oEvalDetBalance As clsEvalDetBalance()
    Public Property oEvalsemovientes As clsEvalsemovientes()
    Public Property oEvalDeudas As clsEvalDeudas()
    Public Property oevalotrosneg As clsevalotrosneg()
    Public Property oEvalIngEgrEstac As clsEvalIngEgrEstac()
    Public Property oEvalgiros As clsevalgiros()
    Public Property oEvalAgricola As clsEvalAgricola()

End Class
