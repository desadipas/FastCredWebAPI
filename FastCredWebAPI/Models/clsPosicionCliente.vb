Imports System.Data.Entity
Imports FastCredWebAPI.FastCredWebAPI

Public Class clsPosicionCliente

    Private db As New DBNEGOCIOEntities

    Public Property cCliente As String
    Public Property cDirecDom As String
    Public Property cDirecNeg As String
    Public Property cMovil1 As String
    Public Property cCentralNegativa As String
    Public Property nLatNeg As Decimal
    Public Property nLongNeg As Decimal
    Public Property nLatDom As Decimal
    Public Property nLongDom As Decimal
    Public Property RCC As IQueryable(Of RCC_CLiente)
    Public Property Record As IQueryable(Of clsRecord)

    Public Sub CargaDatos(ByVal psDoc As String)
        Dim ArrRec() As clsRecord

        Dim vDatosCli = From P In db.Persona
                        Where P.cDoc = psDoc
                        Select P



        For Each P In vDatosCli
            cCliente = P.cNombres + " " + P.cApellidos
            nLatDom = P.nLatitud
            nLongDom = P.nLongitud
            cDirecDom = P.cDireccionDom + "," + P.cReferenciaDom
            cMovil1 = P.cMovil1
        Next

        Dim vDatosCentralNeg = From P In db.PersCentralNegativa
                               Where P.cDoc = psDoc
                               Select P

        cCentralNegativa = "NO SE ENCUENTRA EN CENTRAL NEGATIVA"

        For Each P In vDatosCentralNeg
            cCentralNegativa = "EN CENTRAL NEGATIVA : " + P.cComentario
        Next

        Dim nValorCodEvalMax As Integer

        Dim vEvaluacionMax = From EV In db.Evaluacion
                             Join P In db.Persona On P.nCodPers Equals EV.nCodPers
                             Where P.cDoc.Trim = psDoc.Trim
                             Group EV By EV.nCodEval Into MyGroup = Group
                             Select MyGroup



        For Each E In vEvaluacionMax
            nValorCodEvalMax = E.Max().nCodEval
            Exit For
        Next

        Dim vEvaluacion = From E In db.Evaluacion
                          Where E.nCodEval = nValorCodEvalMax
                          Select E


        For Each E In vEvaluacion
            nLatNeg = E.nLatNeg
            nLongNeg = E.nLongNeg
            cDirecNeg = E.cDirecNeg + "," + E.cReferenciaNeg
        Next


        'Dim vRCC = From R In db.RCC_CLiente
        '           Where R.Cod_Doc_ID.Trim = psDoc.Trim
        '           Select R
        'RCC = vRCC


        Dim vRercord = From R In db.vRecordCliente_Consulta
                       Where R.cDoc = psDoc
                       Order By R.ncodcred Descending
                       Select New clsRecord With {.cDoc = R.cDoc, .cEstado = R.cEstado, .nCodAge = R.nCodAge, .ncodcred = R.ncodcred, .nCuotas = R.nCuotas, .nDiasAtraso = R.nDiasAtraso, .nDiasAtrasoPromedio = R.nDiasAtrasoPromedio, .nMontoCuota = R.nMontoCuota, .nNumPagos = R.nNumPagos, .nNumPagosConAtraso = R.nNumPagosConAtraso, .nPrestamo = R.nPrestamo, .cAsesor = R.cCodUsuCob, .nTasacomp = R.nTasacomp, .dFecVig = R.dFecVig, .cNomGrupo = R.cNomGrupo, .cNomProd = R.cNomProd}

        ArrRec = vRercord.ToArray

        For i = 0 To ArrRec.Count - 1
            ArrRec(i).CargaDatos()
        Next

        Record = ArrRec.AsQueryable


    End Sub

End Class
