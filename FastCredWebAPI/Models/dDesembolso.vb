Imports System.Data.SqlClient

Public Class dDesembolso

    Public psCodage As String
    Public psCodcred As String
    Public psCodUsu As String
    Public Structure Cuota
        Public Fecha As Date
        Public NumCuota As Integer
        Public Capital As Double
        Public Interes As Double
        Public Igv As Double
        Public InteresNeto As Double
        Public Gasto As Double
    End Structure

    Public Function CuotaPeriodoFijo(ByVal pTasaComp As Double, ByVal pTasaComision As Double,
                ByVal pnTasaSeguro As Double,
                ByVal pnCuotas As Double, ByVal vMonto As Double, ByVal pnPlazo As Integer) As Double

        Dim Pot1 As Double
        Dim nTasaTmp As Double
        Dim nTasaComiTmp As Double

        nTasaTmp = Format((((1 + (pTasaComp / 100)) ^ (pnPlazo / 30)) - 1), "#,0.0000000")

        If pnPlazo = 7 Then
            'Semanal debe venir sin igv luego de llevarlo al factor le sumo el igv
            nTasaTmp = nTasaTmp * 1.18
        End If

        nTasaComiTmp = Format((((1 + (pTasaComision / 100)) ^ (pnPlazo / 30)) - 1), "#,0.0000000")

        If pnPlazo = 14 Then
            nTasaTmp = Format((((1 + (pTasaComp / 100)) ^ (pnPlazo / 28)) - 1), "#,0.0000000")
            nTasaComiTmp = Format((((1 + (pTasaComision / 100)) ^ (pnPlazo / 28)) - 1), "#,0.0000000")
        End If


        nTasaTmp = nTasaTmp + nTasaComiTmp
        nTasaTmp = Format(nTasaTmp * (1 + (pnTasaSeguro / 100)), "#,0.0000000")


        'Obtengo la cuota de pago
        Pot1 = (1 + nTasaTmp) ^ pnCuotas
        CuotaPeriodoFijo = ((Pot1 * nTasaTmp) / (Pot1 - 1)) * vMonto
        CuotaPeriodoFijo = CDbl(Format(CuotaPeriodoFijo, "#0.00"))


    End Function

    Public Function MontoCuotaPorProdSinIGV(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String, ByVal psTipoTarifario As String, ByVal psProd As String, ByVal pnPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String
        Dim pnTasaNeta As Double
        Dim nCuotaNeta As Double
        Dim nGastoFijo As Double
        Dim nTasaFactorDiario As Double
        Dim MontoCuota As Double

        If psProd <> "4" Then

            pnTasaNeta = 2.9

            'If psTipoTarifario = "1" Then
            '    pnTasa = "16.30"
            '    pnTasaNeta = 2.98
            'ElseIf psTipoTarifario = "3" Then
            '    pnTasa = "17.50"
            '    pnTasaNeta = 2.98
            'Else
            '    pnTasa = "23.00"
            '    pnTasaNeta = 2.98
            'End If

            'Diario
            If pnTipoPlazo = 1 Then
                'llevar la tasa a factor diario
                nTasaFactorDiario = (((1 + CDbl(pnTasa) / 100.0) ^ (1.0 / 20.0)) - 1) * 100.0
                MontoCuota = CuotaPeriodoFijo(nTasaFactorDiario, 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
                ValorCuota = MontoCuota

            End If

            'Mensual
            If pnTipoPlazo = 4 Then
                ValorCuota = CuotaPeriodoFijo(CDbl(pnTasa), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            End If

        Else
            'pnTasa = "16.00"
            'pnTasaNeta = 2.98

            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
            ValorCuota = CuotaPeriodoFijo(CDbl(pnTasa), 0, 0, 1, CDbl(pnPrestamo), CInt(pnPlazo))

        End If

        D(0) = Format(ValorCuota, "#0.00")
        MontoCuotaPorProdSinIGV = D

    End Function


    Public Function MontoCuotaPorProd(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String, ByVal psTipoTarifario As String, ByVal psProd As String, ByVal pnPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String
        Dim pnTasaNeta As Double
        Dim nCuotaNeta As Double
        Dim nGastoFijo As Double
        Dim nTasaFactorDiario As Double
        Dim MontoCuota As Double


        If psProd.Trim = "3" Then
            pnPlazo = 1

            pnTasaNeta = 2.98

            'Diario
            If pnTipoPlazo = 1 Then

                ''llevar la tasa a factor diario
                'nTasaFactorDiario = (((1 + CDbl(pnTasa) / 100.0) ^ (1.0 / 20.0)) - 1) * 100.0
                'MontoCuota = CuotaPeriodoFijo((CDbl(nTasaFactorDiario) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
                'ValorCuota = MontoCuota

                'Codigo Anterior
                'mdipas 2024
                'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas) / 30, CDbl(pnPrestamo), 30) / 30
                'nCuotaNeta = CuotaPeriodoFijo((pnTasaNeta * 1.18), 0, 0, CDbl(pnCuotas) / 30, CDbl(pnPrestamo), 30) / 30
                ValorCuota = CuotaPeriodoFijo((CDbl(Val(pnTasa)) * 1.18), 0, 0, CDbl(Val(pnCuotas)) / 30, CDbl(Val(pnPrestamo)), 30) / 30
                nCuotaNeta = CuotaPeriodoFijo((pnTasaNeta * 1.18), 0, 0, CDbl(Val(pnCuotas)) / 30, CDbl(Val(pnPrestamo)), 30) / 30
                'end
                nGastoFijo = ValorCuota - nCuotaNeta
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                ValorCuota = nCuotaNeta + nGastoFijo + Format((nGastoFijo * 0.18), "#0.00")

            End If

            'Mensual
            If pnTipoPlazo = 4 Then
                ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            End If

        ElseIf psProd.Trim = "5" Then
            pnPlazo = 30
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "9" Then
            pnPlazo = 28
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "11" Or psProd.Trim = "12" Then
            pnPlazo = 14
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "6" Or psProd.Trim = "7" Or psProd.Trim = "8" Then
            pnPlazo = 7
            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa)), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), CInt(pnPlazo))
            ValorCuota = ValorCuota
        ElseIf psProd.Trim = "10" Then
            pnPlazo = 30
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))
            'ValorCuota = ValorCuota / 20
        Else
            'pnTasa = "16.00"
            'pnTasaNeta = 2.98
            pnPlazo = 30
            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, 1, CDbl(pnPrestamo), CInt(pnPlazo))

        End If
        'mdipas 2024
        D(0) = Str(Format(ValorCuota, "##,##0.00")) 'Format(ValorCuota, "#0.00")
        MontoCuotaPorProd = D

    End Function

    Public Function MontoCuotaPorProd_flat(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String, ByVal psTipoTarifario As String, ByVal psProd As String, ByVal pnPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String
        Dim pnTasaNeta As Double
        Dim nCuotaNeta As Double
        Dim nGastoFijo As Double
        Dim nTasaFactorDiario As Double
        Dim MontoCuota As Double

        If psProd.Trim = "3" Then
            pnPlazo = 1

            pnTasaNeta = 2.98

            'Diario
            If pnTipoPlazo = 1 Then

                ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas / 20) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas / 20) * pnPrestamo) / pnCuotas) * 0.18)
                nCuotaNeta = (pnPrestamo / pnCuotas) + (((pnTasaNeta / 100 * pnCuotas / 20) * pnPrestamo) / pnCuotas) + ((((pnTasaNeta / 100 * pnCuotas / 20) * pnPrestamo) / pnCuotas) * 0.18)

                nGastoFijo = ValorCuota - nCuotaNeta
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                ValorCuota = nCuotaNeta + nGastoFijo + Format((nGastoFijo * 0.18), "#0.00")

            End If

            'Mensual
            If pnTipoPlazo = 4 Then
                ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) * 0.18)
            End If

        ElseIf psProd.Trim = "5" Then
            pnPlazo = 30
            ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) * 0.18)

        ElseIf psProd.Trim = "9" Then
            pnPlazo = 28
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, pnCuotas, CDbl(pnPrestamo), CInt(pnPlazo))

        ElseIf psProd.Trim = "6" Or psProd.Trim = "7" Or psProd.Trim = "8" Then
            pnPlazo = 7

            ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas / 4) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas / 4) * pnPrestamo) / pnCuotas) * 0.18)
            ValorCuota = ValorCuota
        ElseIf psProd.Trim = "10" Then
            pnPlazo = 30
            ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas) * pnPrestamo) / pnCuotas) * 0.18)
            'ValorCuota = ValorCuota / 20
        ElseIf psProd.Trim = "11" Or psProd.Trim = "12" Then
            ValorCuota = (pnPrestamo / pnCuotas) + (((pnTasa / 100 * pnCuotas / 2) * pnPrestamo) / pnCuotas) + ((((pnTasa / 100 * pnCuotas / 2) * pnPrestamo) / pnCuotas) * 0.18)
            ValorCuota = ValorCuota

        Else
            'pnTasa = "16.00"
            'pnTasaNeta = 2.98
            pnPlazo = 30
            'ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, 1, CDbl(pnPrestamo), CInt(pnPlazo))

        End If
        'mdipas 2024
        D(0) = Str(Format(ValorCuota, "##,##0.00")) 'Format(ValorCuota, "#0.00")
        MontoCuotaPorProd_flat = D

    End Function

    Public Function MontoCuota2(ByVal pnPrestamo As String, ByVal pnCuotas As String, ByVal pnTasa As String, ByVal pnTipoPlazo As String) As String()
        Dim ValorCuota As Double
        Dim D(1) As String


        'Diario
        If pnTipoPlazo = 1 Then
            'MontoCuota = (((((1 + (pnTasa * 1.18) / 100) ^ (pnCuotas * 1.0 / 30.0)) - 1) * pnPrestamo) + pnPrestamo) / pnCuotas
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas) / 30, CDbl(pnPrestamo), 30) / 30
        End If

        'Mensual
        If pnTipoPlazo = 4 Then
            ValorCuota = CuotaPeriodoFijo((CDbl(pnTasa) * 1.18), 0, 0, CDbl(pnCuotas), CDbl(pnPrestamo), 30)
        End If

        D(0) = Format(ValorCuota, "#,0.00")
        MontoCuota2 = D

    End Function



    Public Function RegistraCalendarioPorProducto(ByVal psCodage As String, ByVal psCodCred As String, ByRef pnResp As Integer) As String

        Dim nMontoCuota As Double
        Dim nMontoCuotaNeto As Double
        Dim nGastoFijo As Double
        Dim nGastoFijoIgv As Double

        Dim nPrestamo As Double
        Dim nNumCuotas As Integer
        Dim nTasa As Double
        Dim nTasaFactorDiario As Double
        Dim nInteresEsperado As Double
        Dim nInteresNeto As Double

        Dim nTipoPlazo As Integer
        Dim nTasaCompNeto As Double
        Dim nInteresTempo As Double
        Dim nInteresNetoTempo As Double
        Dim nTasaCompNetoTempo As Double
        Dim nTasaCompTempo As Double

        Dim nSaldoTempo As Double
        Dim InteresNetoIgv As Double
        Dim nProd As Integer
        Dim nTipoTarifario As Integer

        Dim D() As String

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim MatCal() As Cuota
        Dim MatFeriados() As Date

        Dim dFechaTempo As Date
        Dim bFeriado As Boolean

        Dim nNumCuotasTempo As Integer
        Dim dFechaPrimPago As Date
        Dim nDiasTrans As Integer
        Dim nCampSinInt As Integer
        Dim oConect As New clsConecta
        Dim sCadConexion As String

        sCadConexion = oConect.gsCadConex


        Try


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing


            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosCalendarioPorProducto_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = CInt(psCodCred)
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTasa", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTipoPlazo", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTasaCompNeto", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nProd", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTipoTarifario", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pdFecPriPago", SqlDbType.Date).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCampSinInt", SqlDbType.Int).Direction = ParameterDirection.Output


            oComm.ExecuteNonQuery()

            nPrestamo = oComm.Parameters("@nPrestamo").Value
            nNumCuotas = oComm.Parameters("@nNumCuotas").Value
            nTasa = oComm.Parameters("@nTasa").Value
            nTipoPlazo = oComm.Parameters("@nTipoPlazo").Value
            nTasaCompNeto = oComm.Parameters("@nTasaCompNeto").Value
            nProd = oComm.Parameters("@nProd").Value
            nTipoTarifario = oComm.Parameters("@nTipoTarifario").Value
            dFechaPrimPago = oComm.Parameters("@pdFecPriPago").Value
            nCampSinInt = oComm.Parameters("@nCampSinInt").Value

            oConex.Close()

            '*********************************************
            'CARGA FERIADOS
            '*********************************************
            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                i = i + 1
                ReDim Preserve MatFeriados(i)
                MatFeriados(i - 1) = dr.Item("dFecha")

            End While

            dr.Close()
            dr = Nothing

            '****************************************************************************

            If nProd = 10 Then

                nDiasTrans = 30
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf nProd = 4 Then


                nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), dFechaPrimPago)
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)



            ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then
                nDiasTrans = 7
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str((nTasa)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str((nTasaCompNeto)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf nProd = 5 Then

                nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), dFechaPrimPago)
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, "30")
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, "30")
                nMontoCuotaNeto = CDbl(D(0))

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            Else


                D = MontoCuotaPorProdSinIGV(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasa).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nInteresEsperado = (CDbl(D(0)) * nNumCuotas) - nPrestamo

                D = MontoCuotaPorProdSinIGV(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(nTasaCompNeto).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                nInteresNeto = (CDbl(D(0)) * nNumCuotas) - nPrestamo

                'D = MontoCuota2(nPrestamo, nNumCuotas, nTasa, nTipoPlazo)
                'nMontoCuota = CDbl(D(0))
                'D = MontoCuota2(nPrestamo, nNumCuotas, nTasaCompNeto, nTipoPlazo)
                'nMontoCuotaNeto = CDbl(D(0))


                D = MontoCuotaPorProd(Str(nPrestamo), Str(nNumCuotas), Str(nTasa), "1", "1", Str(nProd), "1")
                nMontoCuota = CDbl(D(0))
                D = MontoCuotaPorProd(Str(nPrestamo), Str(nNumCuotas), Str(nTasaCompNeto), "1", "1", Str(nProd), "1")
                nMontoCuotaNeto = CDbl(D(0))


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

                'D = MontoCuotaPorProdSinIGV(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(Math.Round(nTasaCompNeto * 1.18, 2)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                'nMontoCuotaNeto = CDbl(D(0))


            End If

            '****************************************************************************

            'Si producto Diario Pasar tasas de calculo a factor diario
            If nProd = 3 Then
                nTasaCompNetoTempo = Math.Round((((1 + nTasaCompNeto / 100.0) ^ (1 / 20.0)) - 1) * 100.0, 4)
                nTasaCompTempo = Math.Round((((1 + nTasa / 100.0) ^ (1 / 20.0)) - 1) * 100.0, 4)
            Else
                nTasaCompNetoTempo = nTasaCompNeto
                nTasaCompTempo = nTasa
            End If

            Dim bCorrecto As Boolean
            Dim nDiaPago As Integer
            Dim nMesPago As Integer
            Dim bFechaEnc As Boolean
            Dim nDiasFeriado As Integer
            nDiasFeriado = 0

            bCorrecto = False


            Do While Not bCorrecto


                    ReDim MatCal(nNumCuotas)
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                If nProd = 6 Or nProd = 7 Or nProd = 8 Or nProd = 10 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                Else
                    dFechaTempo = dFechaTempo.AddDays(1)

                End If

                '*******************************************************************
                'Si es campaña sin interes correr 1 dia si es diario
                ' o correr 7 dias si es semanal o mensual
                '*******************************************************************
                If nCampSinInt = 1 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                    If nProd = 6 Or nProd = 7 Or nProd = 8 Or nProd = 10 Then
                        dFechaTempo = dFechaTempo.AddDays(7)
                    Else
                        dFechaTempo = dFechaTempo.AddDays(6)
                    End If

                End If

                nSaldoTempo = nPrestamo
                    nNumCuotasTempo = nNumCuotas
                    nDiaPago = Day(dFechaPrimPago)

                    For i = 0 To nNumCuotas - 1

                    If nProd = 4 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaPrimPago
                    ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then

                        'Restar los feriados movidos para no cambiar de dia de fecha
                        If nDiasFeriado > 0 And nDiasFeriado < 7 Then
                            dFechaTempo = dFechaTempo.AddDays(7 - nDiasFeriado)
                        Else
                            dFechaTempo = dFechaTempo.AddDays(7)
                        End If
                        'Volver a cero variable que acumula feriados
                        nDiasFeriado = 0

                    ElseIf nProd = 5 Then
                        nDiasFeriado = 0
                        If i = 0 Then
                            dFechaTempo = dFechaPrimPago
                        Else
                            dFechaTempo = DateAdd(DateInterval.Day, 1, MatCal(i - 1).Fecha)
                            'dFechaTempo = CDate("01/" + Microsoft.VisualBasic.Right("00" + Str(dFechaTempo.Month), 2) + "/" + Str(dFechaTempo.Year).Trim)
                            nMesPago = dFechaTempo.Month
                            bFechaEnc = False
                            Do While Not bFechaEnc
                                If dFechaTempo.Day = nDiaPago Or DateDiff(DateInterval.Month, MatCal(i - 1).Fecha, dFechaTempo) >= 2 Then
                                    bFechaEnc = True
                                    Exit Do
                                End If
                                dFechaTempo = dFechaTempo.AddDays(1)
                            Loop
                        End If
                    ElseIf nProd = 10 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(30)
                    Else
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(1)
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    'SI es feriado incrementa en 1
                    bFeriado = False
                        For j = 0 To MatFeriados.Count - 1
                            If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                                bFeriado = True
                                Exit For
                            End If
                        Next
                        If bFeriado Then
                            dFechaTempo = dFechaTempo.AddDays(1)
                            nDiasFeriado = nDiasFeriado + 1
                            bFeriado = False
                        End If

                        'SI es feriado incrementa en 1
                        bFeriado = False
                        For j = 0 To MatFeriados.Count - 1
                            If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                                bFeriado = True
                                Exit For
                            End If
                        Next
                        If bFeriado Then
                            dFechaTempo = dFechaTempo.AddDays(1)
                            nDiasFeriado = nDiasFeriado + 1
                            bFeriado = False
                        End If

                        'SI es feriado incrementa en 1
                        bFeriado = False
                        For j = 0 To MatFeriados.Count - 1
                            If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                                bFeriado = True
                                Exit For
                            End If
                        Next
                        If bFeriado Then
                            dFechaTempo = dFechaTempo.AddDays(1)
                            nDiasFeriado = nDiasFeriado + 1
                            bFeriado = False
                        End If

                        If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                            dFechaTempo = dFechaTempo.AddDays(1)
                            nDiasFeriado = nDiasFeriado + 1
                        End If

                        If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                            dFechaTempo = dFechaTempo.AddDays(1)
                            nDiasFeriado = nDiasFeriado + 1
                        End If

                        MatCal(i).Fecha = CDate(Format(dFechaTempo, "dd/MM/yyyy"))
                        MatCal(i).NumCuota = i + 1

                    If nProd = 4 Then
                        If i = 0 Then
                            nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        Else
                            nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        End If

                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")

                    ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then
                        '**** Antes estaba asi y calculaba los dias transcurridos
                        'If i = 0 Then
                        '    nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        'Else
                        '    nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        'End If

                        '**** Se cambio para que sea a plazo fijo
                        nDiasTrans = 7

                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")

                    ElseIf nProd = 5 Or nProd = 10 Then

                        If i = 0 Then
                            nDiasTrans = DateDiff(DateInterval.Day, CDate(Format(Now(), "dd/MM/yyyy")), MatCal(i).Fecha)
                        Else

                            nDiasTrans = DateDiff(DateInterval.Day, MatCal(i - 1).Fecha, MatCal(i).Fecha)
                        End If


                        nInteresTempo = Format(CDbl(Format((((1 + (nTasaCompTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format(CDbl(Format((((1 + (nTasaCompNetoTempo / 100)) ^ (nDiasTrans / 30)) - 1), "#,0.0000000")) * nSaldoTempo, "#0.00")

                    Else
                        nInteresTempo = Format((nTasaCompTempo / 100.0) * nSaldoTempo, "#0.00")
                        nInteresNetoTempo = Format((nTasaCompNetoTempo / 100.0) * nSaldoTempo, "#0.00")
                    End If



                    MatCal(i).Interes = nInteresTempo
                        MatCal(i).InteresNeto = nInteresNetoTempo
                        nGastoFijo = MatCal(i).Interes - MatCal(i).InteresNeto
                        MatCal(i).Gasto = nGastoFijo
                        MatCal(i).Igv = CDbl(Format(MatCal(i).Interes * 0.18, "#0.00"))

                        If i = (nNumCuotas - 1) Then
                            MatCal(i).Capital = Format(nSaldoTempo, "#0.00")
                        Else
                            MatCal(i).Capital = Format(nMontoCuota - MatCal(i).Interes - MatCal(i).Igv, "#0.00")
                        End If
                        If MatCal(i).Capital < 0 Then
                            MatCal(i).Capital = 0
                        End If

                        nSaldoTempo = nSaldoTempo - MatCal(i).Capital

                    Next i


                    If nNumCuotas > 1 Then
                        If nProd = 3 Then

                            i = nNumCuotas - 1
                            If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                                bCorrecto = False
                                nTasaCompTempo = nTasaCompTempo - 0.0001
                            ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                                nTasaCompTempo = nTasaCompTempo + 0.0001
                            Else
                                bCorrecto = True
                            End If

                        Else   ' Para Producto Mensual
                            'i = nNumCuotas - 1
                            'If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                            '    bCorrecto = False
                            '    nMontoCuota = nMontoCuota + 0.01
                            'ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                            '    nMontoCuota = nMontoCuota - 0.01
                            'Else
                            '    bCorrecto = True
                            'End If

                            bCorrecto = True

                        End If

                    Else
                        bCorrecto = True
                    End If

                Loop


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            If nProd = 10 Then
                nNumCuotas = nNumCuotas * 20
            End If

            For i = 0 To nNumCuotas - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraCuota", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = CInt(psCodCred)
                oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
                oComm.Parameters.Add("@nNroCuota", SqlDbType.Int).Value = MatCal(i).NumCuota
                oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = MatCal(i).Fecha
                oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = MatCal(i).Capital
                oComm.Parameters.Add("@nIntComp", SqlDbType.Money).Value = MatCal(i).Interes
                oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = MatCal(i).Igv
                oComm.Parameters.Add("@nIntComOrig", SqlDbType.Money).Value = MatCal(i).InteresNeto
                oComm.Parameters.Add("@nComi", SqlDbType.Money).Value = MatCal(i).Gasto

                oComm.ExecuteNonQuery()

            Next

            oConex.Close()

            pnResp = 1
            RegistraCalendarioPorProducto = "Registro con Exito"

        Catch ex As Exception
            pnResp = 0
            RegistraCalendarioPorProducto = ex.Message

        End Try



    End Function


    Public Function RegistraCalendarioPorProducto_FLAT(ByVal psCodage As String, ByVal psCodCred As String, ByRef pnResp As Integer) As String

        Dim nMontoCuota As Double
        Dim nMontoCuotaNeto As Double
        Dim nGastoFijo As Double
        Dim nGastoFijoIgv As Double

        Dim nPrestamo As Double
        Dim nNumCuotas As Integer
        Dim nTasa As Double
        Dim nTasaFactorDiario As Double
        Dim nInteresEsperado As Double
        Dim nInteresNeto As Double

        Dim nTipoPlazo As Integer
        Dim nTasaCompNeto As Double
        Dim nInteresTempo As Double
        Dim nInteresNetoTempo As Double
        Dim nTasaCompNetoTempo As Double
        Dim nTasaCompTempo As Double

        Dim nSaldoTempo As Double
        Dim InteresNetoIgv As Double
        Dim nProd As Integer
        Dim nTipoTarifario As Integer

        Dim D() As String

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer

        Dim MatCal() As Cuota
        Dim MatFeriados() As Date

        Dim dFechaTempo As Date
        Dim bFeriado As Boolean

        Dim nNumCuotasTempo As Integer
        Dim dFechaPrimPago As Date
        Dim nDiasTrans As Integer
        Dim nCampSinInt As Integer
        Dim oConect As New clsConecta
        Dim sCadConexion As String

        sCadConexion = oConect.gsCadConex


        Try


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing


            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosCalendarioPorProducto_2", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = CInt(psCodCred)
            oComm.Parameters.Add("@nPrestamo", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nNumCuotas", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTasa", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTipoPlazo", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTasaCompNeto", SqlDbType.Money).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nProd", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nTipoTarifario", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@pdFecPriPago", SqlDbType.Date).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nCampSinInt", SqlDbType.Int).Direction = ParameterDirection.Output


            oComm.ExecuteNonQuery()

            nPrestamo = oComm.Parameters("@nPrestamo").Value
            nNumCuotas = oComm.Parameters("@nNumCuotas").Value
            nTasa = oComm.Parameters("@nTasa").Value
            nTipoPlazo = oComm.Parameters("@nTipoPlazo").Value
            nTasaCompNeto = oComm.Parameters("@nTasaCompNeto").Value
            nProd = oComm.Parameters("@nProd").Value
            nTipoTarifario = oComm.Parameters("@nTipoTarifario").Value
            dFechaPrimPago = oComm.Parameters("@pdFecPriPago").Value
            nCampSinInt = oComm.Parameters("@nCampSinInt").Value

            oConex.Close()

            '*********************************************
            'CARGA FERIADOS
            '*********************************************
            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            oComm = New SqlClient.SqlCommand("PA_RecuperaFeriados", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure

            oComm.ExecuteNonQuery()
            dr = oComm.ExecuteReader(CommandBehavior.SingleResult)
            i = 0

            While dr.Read
                i = i + 1
                ReDim Preserve MatFeriados(i)
                MatFeriados(i - 1) = dr.Item("dFecha")

            End While

            dr.Close()
            dr = Nothing

            '****************************************************************************

            If nProd = 10 Or nProd = 4 Or nProd = 5 Then

                nDiasTrans = 30
                nMontoCuota = (nPrestamo / nNumCuotas) + ((nTasa * 1.0 / 100.0) * nNumCuotas * nPrestamo) / nNumCuotas + ((nTasa * 1.0 / 100.0) * nNumCuotas * nPrestamo) * 0.18 / nNumCuotas
                nMontoCuotaNeto = (nPrestamo / nNumCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * nNumCuotas * nPrestamo) / nNumCuotas + ((nTasaCompNeto * 1.0 / 100.0) * nNumCuotas * nPrestamo) * 0.18 / nNumCuotas

                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)


            ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then

                nDiasTrans = 7
                nMontoCuota = (nPrestamo / nNumCuotas) + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) / nNumCuotas + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) * 0.18 / nNumCuotas
                nMontoCuotaNeto = (nPrestamo / nNumCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) / nNumCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) * 0.18 / nNumCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            ElseIf nProd = 11 Or nProd = 12 Then

                nDiasTrans = 15
                nMontoCuota = (nPrestamo / nNumCuotas) + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) / nNumCuotas + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) * 0.18 / nNumCuotas
                nMontoCuotaNeto = (nPrestamo / nNumCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) / nNumCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) * 0.18 / nNumCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

            Else


                nDiasTrans = 1
                nMontoCuota = (nPrestamo / nNumCuotas) + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) / nNumCuotas + ((nTasa * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) * 0.18 / nNumCuotas
                nMontoCuotaNeto = (nPrestamo / nNumCuotas) + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) / nNumCuotas + ((nTasaCompNeto * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) * 0.18 / nNumCuotas


                nGastoFijo = nMontoCuota - nMontoCuotaNeto
                nGastoFijo = Math.Round((nGastoFijo / 1.18), 1, MidpointRounding.AwayFromZero)
                nGastoFijoIgv = Math.Round(nGastoFijo * 0.18, 1, MidpointRounding.AwayFromZero)

                'D = MontoCuotaPorProdSinIGV(Str(nPrestamo).Trim, Str(nNumCuotas).Trim, Str(Math.Round(nTasaCompNeto * 1.18, 2)).Trim, Str(nTipoPlazo).Trim, Str(nTipoTarifario).Trim, Str(nProd).Trim, Str(nDiasTrans).Trim)
                'nMontoCuotaNeto = CDbl(D(0))


            End If

            '****************************************************************************


            'DEFINE TASAS
            nTasaCompNetoTempo = nTasaCompNeto
            nTasaCompTempo = nTasa


            Dim bCorrecto As Boolean
            Dim nDiaPago As Integer
            Dim nMesPago As Integer
            Dim bFechaEnc As Boolean
            Dim nDiasFeriado As Integer
            nDiasFeriado = 0

            bCorrecto = False


            Do While Not bCorrecto


                ReDim MatCal(nNumCuotas)
                dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                If nProd = 3 Then
                    dFechaTempo = dFechaTempo.AddDays(1)
                Else
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))

                End If

                '*******************************************************************
                'Si es campaña sin interes correr 1 dia si es diario
                ' o correr 7 dias si es semanal o mensual
                '*******************************************************************
                If nCampSinInt = 1 Then
                    dFechaTempo = CDate(Format(Now(), "dd/MM/yyyy"))
                    If nProd = 6 Or nProd = 7 Or nProd = 8 Or nProd = 10 Then
                        dFechaTempo = dFechaTempo.AddDays(7)
                    Else
                        dFechaTempo = dFechaTempo.AddDays(6)
                    End If

                End If

                nSaldoTempo = nPrestamo
                nNumCuotasTempo = nNumCuotas
                nDiaPago = Day(dFechaPrimPago)

                For i = 0 To nNumCuotas - 1

                    If nProd = 4 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaPrimPago
                    ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then

                        'Restar los feriados movidos para no cambiar de dia de fecha
                        If nDiasFeriado > 0 And nDiasFeriado < 7 Then
                            dFechaTempo = dFechaTempo.AddDays(7 - nDiasFeriado)
                        Else
                            dFechaTempo = dFechaTempo.AddDays(7)
                        End If
                        'Volver a cero variable que acumula feriados
                        nDiasFeriado = 0

                    ElseIf nProd = 5 Then
                        nDiasFeriado = 0
                        If i = 0 Then
                            dFechaTempo = dFechaPrimPago
                        Else
                            dFechaTempo = DateAdd(DateInterval.Day, 1, MatCal(i - 1).Fecha)
                            'dFechaTempo = CDate("01/" + Microsoft.VisualBasic.Right("00" + Str(dFechaTempo.Month), 2) + "/" + Str(dFechaTempo.Year).Trim)
                            nMesPago = dFechaTempo.Month
                            bFechaEnc = False
                            Do While Not bFechaEnc
                                If dFechaTempo.Day = nDiaPago Or DateDiff(DateInterval.Month, MatCal(i - 1).Fecha, dFechaTempo) >= 2 Then
                                    bFechaEnc = True
                                    Exit Do
                                End If
                                dFechaTempo = dFechaTempo.AddDays(1)
                            Loop
                        End If
                    ElseIf nProd = 10 Then
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(30)
                    Else
                        nDiasFeriado = 0
                        dFechaTempo = dFechaTempo.AddDays(1)
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If


                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    'SI es feriado incrementa en 1
                    bFeriado = False
                    For j = 0 To MatFeriados.Count - 1
                        If MatFeriados(j).Day = dFechaTempo.Day And MatFeriados(j).Month = dFechaTempo.Month And MatFeriados(j).Year = dFechaTempo.Year Then
                            bFeriado = True
                            Exit For
                        End If
                    Next
                    If bFeriado Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                        bFeriado = False
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Saturday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    If dFechaTempo.DayOfWeek = DayOfWeek.Sunday Then
                        dFechaTempo = dFechaTempo.AddDays(1)
                        nDiasFeriado = nDiasFeriado + 1
                    End If

                    MatCal(i).Fecha = CDate(Format(dFechaTempo, "dd/MM/yyyy"))
                    MatCal(i).NumCuota = i + 1

                    If nProd = 4 Then

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * nNumCuotas * nPrestamo) / nNumCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * nNumCuotas * nPrestamo) / nNumCuotas, "#0.00")



                    ElseIf nProd = 6 Or nProd = 7 Or nProd = 8 Then
                        nDiasTrans = 7

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) / nNumCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (nNumCuotas / 4) * nPrestamo) / nNumCuotas, "#0.00")

                    ElseIf nProd = 5 Or nProd = 10 Then

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (nNumCuotas / 1) * nPrestamo) / nNumCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (nNumCuotas / 1) * nPrestamo) / nNumCuotas, "#0.00")

                    ElseIf nProd = 11 Or nProd = 12 Then
                        nDiasTrans = 7

                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) / nNumCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (nNumCuotas / 2) * nPrestamo) / nNumCuotas, "#0.00")


                    Else
                        nInteresTempo = Format(((nTasaCompTempo * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) / nNumCuotas, "#0.00")
                        nInteresNetoTempo = Format(((nTasaCompNetoTempo * 1.0 / 100.0) * (nNumCuotas / 20) * nPrestamo) / nNumCuotas, "#0.00")
                    End If



                    MatCal(i).Interes = nInteresTempo
                    MatCal(i).InteresNeto = nInteresNetoTempo
                    nGastoFijo = MatCal(i).Interes - MatCal(i).InteresNeto
                    MatCal(i).Gasto = nGastoFijo
                    MatCal(i).Igv = CDbl(Format(MatCal(i).Interes * 0.18, "#0.00"))

                    If i = (nNumCuotas - 1) Then
                        MatCal(i).Capital = Format(nSaldoTempo, "#0.00")
                    Else
                        MatCal(i).Capital = Format(nMontoCuota - MatCal(i).Interes - MatCal(i).Igv, "#0.00")
                    End If
                    If MatCal(i).Capital < 0 Then
                        MatCal(i).Capital = 0
                    End If

                    nSaldoTempo = nSaldoTempo - MatCal(i).Capital

                Next i


                If nNumCuotas > 1 Then
                    If nProd = 3 Then

                        i = nNumCuotas - 1
                        If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                            bCorrecto = False
                            nTasaCompTempo = nTasaCompTempo - 0.0001
                        ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                            nTasaCompTempo = nTasaCompTempo + 0.0001
                        Else
                            bCorrecto = True
                        End If

                    Else   ' Para Producto Mensual
                        'i = nNumCuotas - 1
                        'If (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) > 0 Then
                        '    bCorrecto = False
                        '    nMontoCuota = nMontoCuota + 0.01
                        'ElseIf (CDbl(Format(MatCal(i).Capital + MatCal(i).Interes + MatCal(i).Igv, "#0.00")) - nMontoCuota) < -1.0 Then
                        '    nMontoCuota = nMontoCuota - 0.01
                        'Else
                        '    bCorrecto = True
                        'End If

                        bCorrecto = True

                    End If

                Else
                    bCorrecto = True
                End If

            Loop


            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            If nProd = 10 Then
                nNumCuotas = nNumCuotas * 20
            End If

            For i = 0 To nNumCuotas - 1

                oComm = New SqlClient.SqlCommand("PA_RegistraCuota", oConex)
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = CInt(psCodCred)
                oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
                oComm.Parameters.Add("@nNroCuota", SqlDbType.Int).Value = MatCal(i).NumCuota
                oComm.Parameters.Add("@dFecVenc", SqlDbType.Date).Value = MatCal(i).Fecha
                oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = MatCal(i).Capital
                oComm.Parameters.Add("@nIntComp", SqlDbType.Money).Value = MatCal(i).Interes
                oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = MatCal(i).Igv
                oComm.Parameters.Add("@nIntComOrig", SqlDbType.Money).Value = MatCal(i).InteresNeto
                oComm.Parameters.Add("@nComi", SqlDbType.Money).Value = MatCal(i).Gasto

                oComm.ExecuteNonQuery()

            Next

            oConex.Close()

            pnResp = 1
            RegistraCalendarioPorProducto_FLAT = "Registro con Exito"

        Catch ex As Exception
            pnResp = 0
            RegistraCalendarioPorProducto_FLAT = ex.Message

        End Try



    End Function



    Public Function Movil2_RegistraDesembolsoAmpliacionXProd() As String()

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim i As Integer
        Dim nCodAge As Integer
        Dim nCodCred As Integer
        Dim psResp As String
        Dim nResp As Integer
        Dim oConect As New clsConecta
        Dim sCadConexion As String

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        nCodAge = CInt(psCodage)
        nCodCred = CInt(psCodcred)


        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_ValidaEstadoVigente", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
        oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = CInt(psCodcred)

        oComm.Parameters.Add("@sResp", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()
        i = 0


        ReDim D(2)

        If oComm.Parameters("@nResp").Value = 1 Then
            D(0) = "0"
            D(1) = oComm.Parameters("@sResp").Value

            oConex.Close()
            Return D

        End If

        oConex.Close()



        If CInt(psCodage) = 1 Then
            RegistraCalendarioPorProducto(psCodage, psCodcred, psResp)
        Else
            RegistraCalendarioPorProducto_FLAT(psCodage, psCodcred, psResp)
        End If


        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_RegistraDesembolso_Retanqueo2", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
        oComm.Parameters.Add("@nCodcred", SqlDbType.Int).Value = CInt(psCodcred)
        oComm.Parameters.Add("@Lat", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@Long", SqlDbType.Int).Value = 0
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@sResp", SqlDbType.VarChar, 2000).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = psCodUsu


        oComm.ExecuteNonQuery()
        i = 0


        ReDim D(2)

        If oComm.Parameters("@nResp").Value = 1 Then
            D(0) = "1"
            D(1) = "01-PROCESO FINALIZADO CON EXITO"

        Else
            D(0) = "0"
            D(1) = oComm.Parameters("@sResp").Value

        End If


        oConex.Close()

        Return D


    End Function



    Public Function Movil2_RegistraDesembolsoPorProducto() As String()

        Dim dr As SqlClient.SqlDataReader

        Dim oConex As SqlConnection
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim D() As String
        Dim i As Integer
        Dim psResp As String
        Dim nResp As Integer
        Dim oConect As New clsConecta
        Dim sCadConexion As String


        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_ValidaEstadoVigente", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
        oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = CInt(psCodcred)

        oComm.Parameters.Add("@sResp", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()
        i = 0


        ReDim D(2)

        If oComm.Parameters("@nResp").Value = 1 Then
            D(0) = "0"
            D(1) = oComm.Parameters("@sResp").Value

            oConex.Close()
            Return D

        End If

        oConex.Close()



        If CInt(psCodage) = 1 Then
            RegistraCalendarioPorProducto(psCodage, psCodcred, psResp)
        Else
            RegistraCalendarioPorProducto_FLAT(psCodage, psCodcred, psResp)
        End If



        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        oComm = New SqlClient.SqlCommand("PA_Movil_RegistraDesembolso2", oConex)
        oComm.CommandTimeout = 0
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = CInt(psCodage)
        oComm.Parameters.Add("@ncodcred", SqlDbType.Int).Value = CInt(psCodcred)
        oComm.Parameters.Add("@Lat", SqlDbType.Float).Value = 0
        oComm.Parameters.Add("@Long", SqlDbType.Float).Value = 0
        oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = psCodUsu

        oComm.Parameters.Add("@sResp", SqlDbType.VarChar, 2500).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@nResp", SqlDbType.Int).Direction = ParameterDirection.Output


        oComm.ExecuteNonQuery()
        i = 0


        ReDim D(2)

        If oComm.Parameters("@nResp").Value = 1 Then
            D(0) = "1"
            D(1) = "01-PROCESO FINALIZADO CON EXITO"

        Else
            D(0) = "0"
            D(1) = oComm.Parameters("@sResp").Value
        End If


        oConex.Close()



        Return D


    End Function

End Class
