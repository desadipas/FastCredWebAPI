Option Explicit On

<Serializable()> Public Class ClsEval_UniFam
    Public oDisCli As clsDataset
    Public oExpNeg As clsDataset
    Public oRieZonNeg As clsDataset
    Public oRefCom As clsDataset
    Public oComp As clsDataset
    Public oCtasCobAtr As clsDataset
    Public oGraInst As clsDataset
    Public oEdad As clsDataset
    Public oRieZonDom As clsDataset
    Public oSolFam As clsDataset
    Public oAntCred As clsDataset
    Public oDatos As clsDataset
    Public oDependientes As clsDataset

    Public Sub New()
        oDisCli = New clsDataset
        oExpNeg = New clsDataset
        oRieZonNeg = New clsDataset
        oRefCom = New clsDataset
        oComp = New clsDataset
        oCtasCobAtr = New clsDataset
        oGraInst = New clsDataset
        oEdad = New clsDataset
        oRieZonDom = New clsDataset
        oSolFam = New clsDataset
        oAntCred = New clsDataset
        oDatos = New clsDataset
        oDependientes = New clsDataset

    End Sub

End Class
