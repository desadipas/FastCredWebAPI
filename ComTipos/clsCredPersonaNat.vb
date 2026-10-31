<Serializable()> Public Class clsCredPersonaNat

    Public nCodAge As Integer
    Public nCodPers As Integer
    Public nDirTipo1 As Integer
    Public nDirTipo2 As Integer
    Public nDirTipo3 As Integer
    Public nDirTipo4 As Integer
    Public cDirValor1 As String
    Public cDirValor2 As String
    Public cDirValor3 As String
    Public cDirValor4 As String
    Public cDirReferencia As String
    Public cZona As String
    Public cTelefono As String
    Public nEstado As Integer
    Public nTipoPersona As Integer
    Public cCIIU As String
    Public cApePat As String
    Public cApeMat As String
    Public cNombres As String
    Public cApeCas As String
    Public cSexo As String
    Public dFecNac As DateTime
    Public cMovil As String
    Public cDNI As String
    Public nEstadoCivil As Integer
    Public nNacionalidad As Integer
    Public nRelacIns As Integer
    Public cRpm_RPC As String
    Public cMail As String
    Public bParentesco As Integer
    Public bImpactoSocial As Integer

    Public Documentos() As clsTipoCombo
    Public Parentesco() As clsTipoRelacionPer
    Public ImpactoSocial As DataTable

End Class
