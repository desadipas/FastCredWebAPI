Imports System.Data.SqlClient

Public Class dTransacPago

    Public Property nId As Long
    Public Property sCodAge As String
    Public Property sCodCred As String
    Public Property sMontoPago As String
    Public Property cUser As String
    Public Property psMetLiquid As String
    Public Property sLat As String
    Public Property sLong As String

    Public Property oListaPagoLoteCreditos As clsPagoLoteItem()

    Public Function DevuelveVarNegocio(ByVal pnCodVariable As Integer) As String

        Dim objCon As New clsConecta
        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_DevuelveVarNegocio", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCodVar", SqlDbType.Int).Value = pnCodVariable
        oComm.Parameters.Add("@psValor", SqlDbType.VarChar, 250).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        DevuelveVarNegocio = oComm.Parameters("@psValor").Value

        objCon.CierraConexion()


    End Function

    Public Function DeudaALAFecha(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Double

        Dim oComm As SqlClient.SqlCommand
        Dim objCon As New clsConecta

        oComm = New SqlClient.SqlCommand("PA_DeudaALAFecha", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nDeudaALAFecha", SqlDbType.Money).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        DeudaALAFecha = oComm.Parameters("@nDeudaALAFecha").Value

        objCon.CierraConexion()

    End Function


    Public Sub RecuperaValoresAFecha(ByRef pnIntComp As Double, ByRef pnIntMor As Double, ByRef pnGasto As Double, ByRef pnCapital As Double, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer)

        Dim oComm As SqlClient.SqlCommand
        Dim objCon As New clsConecta

        oComm = New SqlClient.SqlCommand("WEB_RecuperaValoresAFecha_2", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@pnIntComp", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@pnIntMor", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@pnGastos", SqlDbType.Money).Direction = ParameterDirection.Output
        oComm.Parameters.Add("@pnCapital", SqlDbType.Money).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        pnIntComp = oComm.Parameters("@pnIntComp").Value
        pnIntMor = oComm.Parameters("@pnIntMor").Value
        pnGasto = oComm.Parameters("@pnGastos").Value
        pnCapital = oComm.Parameters("@pnCapital").Value

        objCon.CierraConexion()


    End Sub


    Public Function Web_PagarCredito(ByVal psCodAge As String, ByVal psCodCred As String, ByVal psMontoPago As String,
        ByVal cUser As String, ByVal psMetLiquid As String, ByVal sLat As String, ByVal sLong As String) As String()

        Dim nIgvValor As Double
        Dim objCon As New clsConecta
        Dim nDeudaALAFecha As Double
        Dim reader As SqlClient.SqlDataReader
        Dim MatCalend(1000, 13) As String
        Dim nFilas As Integer
        Dim nAmortizacion As Double
        Dim dFecPago As Date
        Dim nEstadoCred As Integer
        Dim nNroTran As Long

        Dim pnCodAge As Integer = CInt(psCodAge)
        Dim pnCodCred As Integer = CInt(psCodCred)
        Dim pnMontoPago As Double = CDbl(psMontoPago)
        Dim nLat As Decimal = CDec(sLat)
        Dim nLong As Decimal = CDec(sLong)
        Dim nProd As Integer
        Dim nPagoMinimo As Double
        Dim D() As String
        Dim nIntPend As Double

        Try
            Dim oComm As SqlClient.SqlCommand
            Dim sRespTope As String
            '***************************************************************************
            'Validar que no halla pasado el tope Permitido para Pagos

            oComm = New SqlClient.SqlCommand("ValidaTopeParaRealizarPago", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@pnMontoPago", SqlDbType.Money).Value = CDbl(psMontoPago)
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            sRespTope = oComm.Parameters("@psResul").Value

            objCon.CierraConexion()


            'If sRespTope.Trim <> "OK" Then

            '    ReDim D(2)
            '    D(0) = "1"
            '    D(1) = "ERROR : " + sRespTope
            '    Return D

            'End If

            '***************************************************************************


            '***************************************************************************
            'Validar que Credito esté Vigente

            oComm = New SqlClient.SqlCommand("ValidaReglas", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            sRespTope = oComm.Parameters("@psResul").Value

            objCon.CierraConexion()


            If sRespTope.Trim <> "OK" Then

                ReDim D(2)
                D(0) = "0"
                D(1) = "ERROR : " + sRespTope
                Return D
            End If


            '******************************************************************************
            nIgvValor = CDbl(DevuelveVarNegocio(4)) / 100.0

            dFecPago = Now()

            nDeudaALAFecha = Format(DeudaALAFecha(pnCodAge, pnCodCred), "#0.00")

            If pnMontoPago > nDeudaALAFecha Then

                ReDim D(2)
                D(0) = "0"
                D(1) = Str(pnCodAge) + " " + Str(pnCodCred) + " Monto de Pago no puede ser mayor a deuda a la fecha : " & Format(nDeudaALAFecha, "#,0.00")
                Return D
                Exit Function
            End If


            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosCreditoPorCodCred", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@pnCodcred", SqlDbType.Int).Value = pnCodCred

            reader = oComm.ExecuteReader()
            Do While reader.Read()
                nProd = reader.Item("nProd")
                nPagoMinimo = reader.Item("nPagoMinimo")
            Loop
            reader.Close()
            objCon.CierraConexion()


            oComm = New SqlClient.SqlCommand("PA_Recupera_Calendario", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred

            reader = oComm.ExecuteReader()


            'Carga Calendario
            nFilas = 0
            Do While reader.Read()
                nFilas = nFilas + 1
                MatCalend(nFilas, 1) = reader.Item("dFecVenc")
                MatCalend(nFilas, 2) = reader.Item("nNroCuota")
                MatCalend(nFilas, 3) = Format(CDbl(reader.Item("nCapital")) - CDbl(reader.Item("nCapPag")), "#,0.00")
                MatCalend(nFilas, 4) = "0.00"
                MatCalend(nFilas, 5) = Format(CDbl(reader.Item("nIntComp")) - CDbl(reader.Item("nIntPag")), "#,0.00")
                MatCalend(nFilas, 6) = "0.00"
                MatCalend(nFilas, 7) = Format(CDbl(reader.Item("nIntMor")) - CDbl(reader.Item("nIntMorPag")), "#,0.00")
                MatCalend(nFilas, 8) = "0.00"
                MatCalend(nFilas, 9) = Format(CDbl(reader.Item("nIgv")) - CDbl(reader.Item("nIgvPag")), "#,0.00")
                MatCalend(nFilas, 10) = "0.00"
                MatCalend(nFilas, 11) = "0"
                '*****
                MatCalend(nFilas, 12) = Format(CDbl(reader.Item("nComi")) - CDbl(reader.Item("nComiPag")), "#,0.00")
                MatCalend(nFilas, 13) = "0"

            Loop

            reader.Close()

            objCon.CierraConexion()

            '**********************************************
            'AMORTIZAR CALENDARIO
            '**********************************************
            Dim i As Integer
            Dim j As Integer

            nAmortizacion = pnMontoPago

            '****************************************************************************
            Dim nSaldoIntComp As Double
            Dim nSaldoIntMor As Double
            Dim nSaldoGastos As Double
            Dim nSaldoCapital As Double
            Dim nSaldoPago As Double

            Call RecuperaValoresAFecha(nSaldoIntComp, nSaldoIntMor, nSaldoGastos, nSaldoCapital, pnCodAge, pnCodCred)

            nIntPend = 0
            If nProd = 4 Then
                If pnMontoPago < nSaldoIntComp Then
                    nIntPend = nSaldoIntComp - pnMontoPago
                    'ReDim D(2)
                    'D(0) = "0"
                    'D(1) = "Monto de Pago es Menor al Pago Minimo"
                    'Return D
                End If
            End If


            'FORZAR PAGO HORIZONTAL
            If nProd = 4 Or nProd = 5 Then
                psMetLiquid = "REV1"

            Else
                psMetLiquid = "GMIC"

            End If
            nSaldoPago = pnMontoPago

            '**************SI MONTO DE PAGO NO CUBRE DEUDA A LA FECHA ********************
            If pnMontoPago < nDeudaALAFecha Then
                'SI ES REVOLVENTE, todo las amortizaciones es a la fecha, corren las fechas de los vencimientos
                If psMetLiquid = "REV1" Or psMetLiquid = "REV2" Then

                    If nSaldoPago > nSaldoIntMor Then
                        nAmortizacion = nSaldoIntMor
                        nSaldoPago = nSaldoPago - nSaldoIntMor
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If


                    For i = 1 To nFilas
                        'Amortiza Mora
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                                MatCalend(i, 8) = MatCalend(i, 7)
                                MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                                nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                            Else
                                MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If

                        End If
                    Next

                    'Amortiza Interes + Gastos Adm
                    If nSaldoPago > nSaldoIntComp Then
                        nAmortizacion = nSaldoIntComp
                        nSaldoPago = nSaldoPago - nSaldoIntComp
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                                MatCalend(i, 6) = MatCalend(i, 5)
                                'MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                                'nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                                MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 9))), "#,0.00")
                                nAmortizacion = nAmortizacion - CDbl(Format(CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 9)), "#,0.00"))
                            Else
                                MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If

                        End If
                    Next


                    If nSaldoPago > nSaldoGastos Then
                        nAmortizacion = nSaldoGastos
                        nSaldoPago = nSaldoPago - nSaldoGastos
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Amortiza Gasto Fijo
                            If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                                If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                                    MatCalend(i, 13) = MatCalend(i, 12)
                                Else
                                    MatCalend(i, 13) = MatCalend(i, 6)
                                End If
                            End If
                        End If
                    Next


                    'Amortiza Capital
                    If nSaldoPago > nSaldoCapital Then
                        nAmortizacion = nSaldoCapital
                        nSaldoPago = nSaldoPago - nSaldoCapital
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                                MatCalend(i, 4) = MatCalend(i, 3)
                                'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                                nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                            Else
                                MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                                'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If
                        End If
                    Next


                    'Genera un nuevo calendario
                    If psMetLiquid = "REV2" Then
                        If pnMontoPago > (nSaldoIntComp + nSaldoIntMor + nSaldoGastos) Then



                        End If
                    End If
                Else
                    For i = 1 To nFilas
                        For j = 1 To 4
                            If Mid(psMetLiquid, j, 1) = "M" Then
                                'Amortiza Mora
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                                        MatCalend(i, 8) = MatCalend(i, 7)
                                        MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                                        nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                                    Else
                                        MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                        MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If

                                End If
                            End If


                            If Mid(psMetLiquid, j, 1) = "I" Then
                                'Amortiza Interes + Gastos Adm
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                                        MatCalend(i, 6) = MatCalend(i, 5)
                                        'MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                                        'nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                                        MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 9))), "#,0.00")
                                        nAmortizacion = nAmortizacion - CDbl(Format(CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 9)), "#,0.00"))
                                    Else
                                        MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                        MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If

                                End If
                            End If

                            If Mid(psMetLiquid, j, 1) = "G" Then
                                'Amortiza Gasto Fijo
                                If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                                    If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                                        MatCalend(i, 13) = MatCalend(i, 12)
                                    Else
                                        MatCalend(i, 13) = MatCalend(i, 6)
                                    End If
                                End If
                            End If

                            If Mid(psMetLiquid, j, 1) = "C" Then
                                'Amortiza Capital
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                                        MatCalend(i, 4) = MatCalend(i, 3)
                                        'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                                        nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                                    Else
                                        MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                                        'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If
                                End If
                            End If

                        Next
                    Next
                End If
            Else
                '**************SI MONTO DE PAGO  CUBRE DEUDA A LA FECHA ********************
                'AMORTIZA TODO CAPITAL
                For i = 1 To nFilas
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                            MatCalend(i, 4) = MatCalend(i, 3)
                            'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                        Else
                            MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                            'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If
                    End If
                Next

                'AMORTIZA TODO INTERES
                For i = 1 To nFilas

                    'Amortiza Mora
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                            MatCalend(i, 8) = MatCalend(i, 7)
                            MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                        Else
                            MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                            MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If

                    End If

                    'Amortiza Interes + Gastos Adm
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                            MatCalend(i, 6) = MatCalend(i, 5)
                            MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                        Else
                            MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                            MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If

                    End If

                    'Amortiza Gasto Fijo
                    'If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                    '    If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                    '        MatCalend(i, 13) = MatCalend(i, 12)
                    '    Else
                    '        MatCalend(i, 13) = MatCalend(i, 6)
                    '    End If
                    'End If
                Next

            End If

            '****************************************************************************
            For i = 1 To nFilas
                'valida si todos los conceptos ya estan pagados
                If (CDbl(MatCalend(i, 3)) - CDbl(MatCalend(i, 4))) <= 0.05 And (CDbl(MatCalend(i, 5)) - CDbl(MatCalend(i, 6))) <= 0.05 And (CDbl(MatCalend(i, 7)) - CDbl(MatCalend(i, 8))) <= 0.05 Then
                    MatCalend(i, 11) = "1"
                End If

            Next

            'Cancelado
            nEstadoCred = 5
            For i = 1 To nFilas
                If MatCalend(i, 11) = "0" Then
                    nEstadoCred = 4
                    Exit For
                End If
            Next

            'Para que cencele el credito si es que paga la deuda a la fecha
            If pnMontoPago >= nDeudaALAFecha Then
                nEstadoCred = 5
            End If


            Dim nCapital As Double
            Dim nInteres As Double
            Dim nMora As Double
            Dim nIgv As Double
            Dim nGastoFijo As Double

            nCapital = 0
            nInteres = 0
            nMora = 0
            nIgv = 0
            nGastoFijo = 0
            For i = 1 To nFilas

                nCapital = nCapital + CDbl(MatCalend(i, 4))
                nInteres = nInteres + CDbl(MatCalend(i, 6))
                nMora = nMora + CDbl(MatCalend(i, 8))
                nIgv = nIgv + CDbl(MatCalend(i, 10))
                nGastoFijo = nGastoFijo + CDbl(MatCalend(i, 13))
            Next

            'Actualiza Cabeceras
            oComm = New SqlClient.SqlCommand("PA_RegistraPagoPorProd_2", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nID", SqlDbType.Int).Value = nId
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
            oComm.Parameters.Add("@nEstadoCred", SqlDbType.Int).Value = nEstadoCred
            oComm.Parameters.Add("@nMontopago", SqlDbType.Money).Value = pnMontoPago
            oComm.Parameters.Add("@Capital", SqlDbType.Money).Value = nCapital
            oComm.Parameters.Add("@nInteres", SqlDbType.Money).Value = nInteres - nGastoFijo
            'Nuevo Contrato
            oComm.Parameters.Add("@nGastoFijo", SqlDbType.Money).Value = nGastoFijo
            oComm.Parameters.Add("@Mora", SqlDbType.Int).Value = nMora
            oComm.Parameters.Add("@dFechaPago", SqlDbType.DateTime).Value = dFecPago
            oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = nIgv
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLat
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = nLong
            oComm.Parameters.Add("@nNroTran", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nIntPend", SqlDbType.Money).Value = nIntPend


            oComm.ExecuteNonQuery()

            nNroTran = oComm.Parameters("@nNroTran").Value

            objCon.CierraConexion()


            For i = 1 To nFilas
                '

                If CDbl(MatCalend(i, 4)) > 0 Or CDbl(MatCalend(i, 6)) > 0 Or CDbl(MatCalend(i, 8)) > 0 Then

                    oComm = New SqlClient.SqlCommand("PA_AmortizaCalendario", objCon.DameConexion)
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
                    oComm.Parameters.Add("@nNumCuota", SqlDbType.Int).Value = MatCalend(i, 2)
                    oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = MatCalend(i, 4)
                    oComm.Parameters.Add("@nInteres", SqlDbType.Money).Value = MatCalend(i, 6)
                    '************
                    oComm.Parameters.Add("@nGastoFijo", SqlDbType.Money).Value = MatCalend(i, 13)

                    oComm.Parameters.Add("@nMora", SqlDbType.Money).Value = MatCalend(i, 8)
                    oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = MatCalend(i, 11)
                    oComm.Parameters.Add("@dFecgaPago", SqlDbType.DateTime).Value = dFecPago
                    oComm.Parameters.Add("@nNroTran", SqlDbType.Int).Value = nNroTran
                    oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
                    oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = MatCalend(i, 10)

                    '
                    oComm.ExecuteNonQuery()

                End If

            Next

            'Aplica Ampliacion
            'oComm = New SqlClient.SqlCommand("PA_ActualizaAplicaAmpliacion", objCon.DameConexion)
            'oComm.CommandType = CommandType.StoredProcedure

            'oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            'oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred


            'oComm.ExecuteNonQuery()



            objCon.CierraConexion()

            'Renueva Calendario Para creditos a 1 sola cuota
            If pnMontoPago < nDeudaALAFecha Then
                If nIntPend <= 0 Then
                    oComm = New SqlClient.SqlCommand("Web_RenuevaCalendarioUnaCuota_2", objCon.DameConexion)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
                    oComm.ExecuteNonQuery()

                    objCon.CierraConexion()
                End If

            End If


            Dim pnDeudaAFecha As Double
            Dim dProxfexVen As Date
            pnDeudaAFecha = DeudaCalendario(pnCodAge, pnCodCred)
            dProxfexVen = ProxFechaPago(pnCodAge, pnCodCred)

            ReDim D(2)
            D(0) = "1"
            D(1) = "Pago Correcto, Proxima fecha " + Format(dProxfexVen, "dd/MM/yyyy") + " Deuda : " + Format(pnDeudaAFecha, "#0.00")

            Return D

        Catch ex As Exception
            ReDim D(2)
            D(0) = "0"
            D(1) = "Error : " + Err.Description
            Return D


        End Try


    End Function

    Public Function Web_PagarCreditoLote(ByVal psCodAge As String, ByVal psCodCred As String, ByVal psMontoPago As String,
        ByVal cUser As String, ByVal psMetLiquid As String, ByVal sLat As String, ByVal sLong As String, nIDPagolote As Long) As String()

        Dim nIgvValor As Double
        Dim objCon As New clsConecta
        Dim nDeudaALAFecha As Double
        Dim reader As SqlClient.SqlDataReader
        Dim MatCalend(1000, 13) As String
        Dim nFilas As Integer
        Dim nAmortizacion As Double
        Dim dFecPago As Date
        Dim nEstadoCred As Integer
        Dim nNroTran As Long

        Dim pnCodAge As Integer = CInt(psCodAge)
        Dim pnCodCred As Integer = CInt(psCodCred)
        Dim pnMontoPago As Double = CDbl(psMontoPago)
        Dim nLat As Decimal = CDec(sLat)
        Dim nLong As Decimal = CDec(sLong)
        Dim nProd As Integer
        Dim nPagoMinimo As Double
        Dim D() As String
        Dim nIntPend As Double

        Try
            Dim oComm As SqlClient.SqlCommand
            Dim sRespTope As String
            '***************************************************************************
            'Validar que no halla pasado el tope Permitido para Pagos

            oComm = New SqlClient.SqlCommand("ValidaTopeParaRealizarPago", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@pnMontoPago", SqlDbType.Money).Value = CDbl(psMontoPago)
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            sRespTope = oComm.Parameters("@psResul").Value

            objCon.CierraConexion()


            'If sRespTope.Trim <> "OK" Then

            '    ReDim D(2)
            '    D(0) = "1"
            '    D(1) = "ERROR : " + sRespTope
            '    Return D

            'End If

            '***************************************************************************


            '***************************************************************************
            'Validar que Credito esté Vigente

            oComm = New SqlClient.SqlCommand("ValidaReglas", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
            oComm.Parameters.Add("@psResul", SqlDbType.VarChar, 1500).Direction = ParameterDirection.Output
            oComm.ExecuteNonQuery()

            sRespTope = oComm.Parameters("@psResul").Value

            objCon.CierraConexion()


            If sRespTope.Trim <> "OK" Then

                ReDim D(2)
                D(0) = "0"
                D(1) = "ERROR : " + sRespTope
                Return D
            End If


            '******************************************************************************
            nIgvValor = CDbl(DevuelveVarNegocio(4)) / 100.0

            dFecPago = Now()

            nDeudaALAFecha = Format(DeudaALAFecha(pnCodAge, pnCodCred), "#0.00")

            'If pnMontoPago > nDeudaALAFecha Then

            '    ReDim D(2)
            '    D(0) = "0"
            '    D(1) = Str(pnCodAge) + " " + Str(pnCodCred) + " Monto de Pago no puede ser mayor a deuda a la fecha : " & Format(nDeudaALAFecha, "#,0.00")
            '    Return D
            '    Exit Function
            'End If


            oComm = New SqlClient.SqlCommand("PA_RecuperaDatosCreditoPorCodCred", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@pnCodcred", SqlDbType.Int).Value = pnCodCred

            reader = oComm.ExecuteReader()
            Do While reader.Read()
                nProd = reader.Item("nProd")
                nPagoMinimo = reader.Item("nPagoMinimo")
            Loop
            reader.Close()
            objCon.CierraConexion()


            oComm = New SqlClient.SqlCommand("PA_Recupera_Calendario", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred

            reader = oComm.ExecuteReader()


            'Carga Calendario
            nFilas = 0
            Do While reader.Read()
                nFilas = nFilas + 1
                MatCalend(nFilas, 1) = reader.Item("dFecVenc")
                MatCalend(nFilas, 2) = reader.Item("nNroCuota")
                MatCalend(nFilas, 3) = Format(CDbl(reader.Item("nCapital")) - CDbl(reader.Item("nCapPag")), "#,0.00")
                MatCalend(nFilas, 4) = "0.00"
                MatCalend(nFilas, 5) = Format(CDbl(reader.Item("nIntComp")) - CDbl(reader.Item("nIntPag")), "#,0.00")
                MatCalend(nFilas, 6) = "0.00"
                MatCalend(nFilas, 7) = Format(CDbl(reader.Item("nIntMor")) - CDbl(reader.Item("nIntMorPag")), "#,0.00")
                MatCalend(nFilas, 8) = "0.00"
                MatCalend(nFilas, 9) = Format(CDbl(reader.Item("nIgv")) - CDbl(reader.Item("nIgvPag")), "#,0.00")
                MatCalend(nFilas, 10) = "0.00"
                MatCalend(nFilas, 11) = "0"
                '*****
                MatCalend(nFilas, 12) = Format(CDbl(reader.Item("nComi")) - CDbl(reader.Item("nComiPag")), "#,0.00")
                MatCalend(nFilas, 13) = "0"

            Loop

            reader.Close()

            objCon.CierraConexion()

            '**********************************************
            'AMORTIZAR CALENDARIO
            '**********************************************
            Dim i As Integer
            Dim j As Integer

            nAmortizacion = pnMontoPago

            '****************************************************************************
            Dim nSaldoIntComp As Double
            Dim nSaldoIntMor As Double
            Dim nSaldoGastos As Double
            Dim nSaldoCapital As Double
            Dim nSaldoPago As Double

            Call RecuperaValoresAFecha(nSaldoIntComp, nSaldoIntMor, nSaldoGastos, nSaldoCapital, pnCodAge, pnCodCred)

            nIntPend = 0
            If nProd = 4 Then
                If pnMontoPago < nSaldoIntComp Then
                    nIntPend = nSaldoIntComp - pnMontoPago
                    'ReDim D(2)
                    'D(0) = "0"
                    'D(1) = "Monto de Pago es Menor al Pago Minimo"
                    'Return D
                End If
            End If


            'FORZAR PAGO HORIZONTAL
            If nProd = 4 Or nProd = 5 Then
                psMetLiquid = "REV1"

            Else
                psMetLiquid = "GMIC"

            End If
            nSaldoPago = pnMontoPago

            '**************SI MONTO DE PAGO NO CUBRE DEUDA A LA FECHA ********************
            If pnMontoPago < nDeudaALAFecha Then
                'SI ES REVOLVENTE, todo las amortizaciones es a la fecha, corren las fechas de los vencimientos
                If psMetLiquid = "REV1" Or psMetLiquid = "REV2" Then

                    If nSaldoPago > nSaldoIntMor Then
                        nAmortizacion = nSaldoIntMor
                        nSaldoPago = nSaldoPago - nSaldoIntMor
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If


                    For i = 1 To nFilas
                        'Amortiza Mora
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                                MatCalend(i, 8) = MatCalend(i, 7)
                                MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                                nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                            Else
                                MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If

                        End If
                    Next

                    'Amortiza Interes + Gastos Adm
                    If nSaldoPago > nSaldoIntComp Then
                        nAmortizacion = nSaldoIntComp
                        nSaldoPago = nSaldoPago - nSaldoIntComp
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                                MatCalend(i, 6) = MatCalend(i, 5)
                                'MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                                'nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                                MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 9))), "#,0.00")
                                nAmortizacion = nAmortizacion - CDbl(Format(CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 9)), "#,0.00"))
                            Else
                                MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If

                        End If
                    Next


                    If nSaldoPago > nSaldoGastos Then
                        nAmortizacion = nSaldoGastos
                        nSaldoPago = nSaldoPago - nSaldoGastos
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Amortiza Gasto Fijo
                            If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                                If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                                    MatCalend(i, 13) = MatCalend(i, 12)
                                Else
                                    MatCalend(i, 13) = MatCalend(i, 6)
                                End If
                            End If
                        End If
                    Next


                    'Amortiza Capital
                    If nSaldoPago > nSaldoCapital Then
                        nAmortizacion = nSaldoCapital
                        nSaldoPago = nSaldoPago - nSaldoCapital
                    Else
                        nAmortizacion = nSaldoPago
                        nSaldoPago = 0
                    End If

                    For i = 1 To nFilas
                        If nAmortizacion > 0 Then
                            'Si se puede cancelar todo se cancela
                            If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                                MatCalend(i, 4) = MatCalend(i, 3)
                                'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                                nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                            Else
                                MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                                'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                nAmortizacion = 0
                                Exit For
                            End If
                        End If
                    Next


                    'Genera un nuevo calendario
                    If psMetLiquid = "REV2" Then
                        If pnMontoPago > (nSaldoIntComp + nSaldoIntMor + nSaldoGastos) Then



                        End If
                    End If
                Else
                    For i = 1 To nFilas
                        For j = 1 To 4
                            If Mid(psMetLiquid, j, 1) = "M" Then
                                'Amortiza Mora
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                                        MatCalend(i, 8) = MatCalend(i, 7)
                                        MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                                        nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                                    Else
                                        MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                        MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If

                                End If
                            End If


                            If Mid(psMetLiquid, j, 1) = "I" Then
                                'Amortiza Interes + Gastos Adm
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                                        MatCalend(i, 6) = MatCalend(i, 5)
                                        'MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                                        'nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                                        MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 9))), "#,0.00")
                                        nAmortizacion = nAmortizacion - CDbl(Format(CDbl(MatCalend(i, 5)) + CDbl(MatCalend(i, 9)), "#,0.00"))
                                    Else
                                        MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                                        MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If

                                End If
                            End If

                            If Mid(psMetLiquid, j, 1) = "G" Then
                                'Amortiza Gasto Fijo
                                If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                                    If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                                        MatCalend(i, 13) = MatCalend(i, 12)
                                    Else
                                        MatCalend(i, 13) = MatCalend(i, 6)
                                    End If
                                End If
                            End If

                            If Mid(psMetLiquid, j, 1) = "C" Then
                                'Amortiza Capital
                                If nAmortizacion > 0 Then
                                    'Si se puede cancelar todo se cancela
                                    If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                                        MatCalend(i, 4) = MatCalend(i, 3)
                                        'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                                        nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                                    Else
                                        MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                                        'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                                        nAmortizacion = 0
                                        Exit For
                                    End If
                                End If
                            End If

                        Next
                    Next
                End If
            Else
                '**************SI MONTO DE PAGO  CUBRE DEUDA A LA FECHA ********************
                'AMORTIZA TODO CAPITAL
                For i = 1 To nFilas
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > (CDbl(MatCalend(i, 3))) Then
                            MatCalend(i, 4) = MatCalend(i, 3)
                            'MatCalend(nFilas, 10) = Format(CDbl(MatCalend(nFilas, 10)) + (CDbl(MatCalend(nFilas, 3)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - (CDbl(MatCalend(i, 3)))
                        Else
                            MatCalend(i, 4) = Format(nAmortizacion, "#,0.00")
                            'MatCalend(nFilas, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If
                    End If
                Next

                'AMORTIZA TODO INTERES
                For i = 1 To nFilas

                    'Amortiza Mora
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00")) Then
                            MatCalend(i, 8) = MatCalend(i, 7)
                            MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 7)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 7)) * (1 + nIgvValor)), "#,0.00"))
                        Else
                            MatCalend(i, 8) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                            MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If

                    End If

                    'Amortiza Interes + Gastos Adm
                    If nAmortizacion > 0 Then
                        'Si se puede cancelar todo se cancela
                        If nAmortizacion > CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00")) Then
                            MatCalend(i, 6) = MatCalend(i, 5)
                            MatCalend(i, 10) = Format(CDbl(MatCalend(i, 10)) + (CDbl(MatCalend(i, 5)) * nIgvValor), "#,0.00")
                            nAmortizacion = nAmortizacion - CDbl(Format((CDbl(MatCalend(i, 5)) * (1 + nIgvValor)), "#,0.00"))
                        Else
                            MatCalend(i, 6) = Format(nAmortizacion / (1 + nIgvValor), "#,0.00")
                            MatCalend(i, 10) = Format((nAmortizacion / (1 + nIgvValor)) * nIgvValor, "#,0.00")
                            nAmortizacion = 0
                            Exit For
                        End If

                    End If

                    'Amortiza Gasto Fijo
                    'If CDbl(MatCalend(i, 6)) > 0 And CDbl(MatCalend(i, 12)) > 0 Then
                    '    If CDbl(MatCalend(i, 6)) >= CDbl(MatCalend(i, 12)) Then
                    '        MatCalend(i, 13) = MatCalend(i, 12)
                    '    Else
                    '        MatCalend(i, 13) = MatCalend(i, 6)
                    '    End If
                    'End If
                Next

            End If

            '****************************************************************************
            For i = 1 To nFilas
                'valida si todos los conceptos ya estan pagados
                If (CDbl(MatCalend(i, 3)) - CDbl(MatCalend(i, 4))) <= 0.05 And (CDbl(MatCalend(i, 5)) - CDbl(MatCalend(i, 6))) <= 0.05 And (CDbl(MatCalend(i, 7)) - CDbl(MatCalend(i, 8))) <= 0.05 Then
                    MatCalend(i, 11) = "1"
                End If

            Next

            'Cancelado
            nEstadoCred = 5
            For i = 1 To nFilas
                If MatCalend(i, 11) = "0" Then
                    nEstadoCred = 4
                    Exit For
                End If
            Next

            'Para que cencele el credito si es que paga la deuda a la fecha
            If pnMontoPago >= nDeudaALAFecha Then
                nEstadoCred = 5
            End If


            Dim nCapital As Double
            Dim nInteres As Double
            Dim nMora As Double
            Dim nIgv As Double
            Dim nGastoFijo As Double

            nCapital = 0
            nInteres = 0
            nMora = 0
            nIgv = 0
            nGastoFijo = 0
            For i = 1 To nFilas

                nCapital = nCapital + CDbl(MatCalend(i, 4))
                nInteres = nInteres + CDbl(MatCalend(i, 6))
                nMora = nMora + CDbl(MatCalend(i, 8))
                nIgv = nIgv + CDbl(MatCalend(i, 10))
                nGastoFijo = nGastoFijo + CDbl(MatCalend(i, 13))
            Next

            'Actualiza Cabeceras
            oComm = New SqlClient.SqlCommand("PA_RegistraPagoPorProd_PagoLote", objCon.DameConexion)
            oComm.CommandType = CommandType.StoredProcedure

            oComm.Parameters.Add("@nID", SqlDbType.Int).Value = nId
            oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
            oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
            oComm.Parameters.Add("@nEstadoCred", SqlDbType.Int).Value = nEstadoCred
            oComm.Parameters.Add("@nMontopago", SqlDbType.Money).Value = pnMontoPago
            oComm.Parameters.Add("@Capital", SqlDbType.Money).Value = nCapital
            oComm.Parameters.Add("@nInteres", SqlDbType.Money).Value = nInteres - nGastoFijo
            'Nuevo Contrato
            oComm.Parameters.Add("@nGastoFijo", SqlDbType.Money).Value = nGastoFijo
            oComm.Parameters.Add("@Mora", SqlDbType.Int).Value = nMora
            oComm.Parameters.Add("@dFechaPago", SqlDbType.DateTime).Value = dFecPago
            oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = nIgv
            oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@nLat", SqlDbType.Float).Value = nLat
            oComm.Parameters.Add("@nLong", SqlDbType.Float).Value = nLong
            oComm.Parameters.Add("@nNroTran", SqlDbType.Int).Direction = ParameterDirection.Output
            oComm.Parameters.Add("@nIntPend", SqlDbType.Money).Value = nIntPend
            oComm.Parameters.Add("@nIDPagoLote", SqlDbType.Int).Value = nIDPagolote


            oComm.ExecuteNonQuery()

            nNroTran = oComm.Parameters("@nNroTran").Value

            objCon.CierraConexion()


            For i = 1 To nFilas
                '

                If CDbl(MatCalend(i, 4)) > 0 Or CDbl(MatCalend(i, 6)) > 0 Or CDbl(MatCalend(i, 8)) > 0 Then

                    oComm = New SqlClient.SqlCommand("PA_AmortizaCalendario", objCon.DameConexion)
                    oComm.CommandType = CommandType.StoredProcedure

                    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
                    oComm.Parameters.Add("@nNumCuota", SqlDbType.Int).Value = MatCalend(i, 2)
                    oComm.Parameters.Add("@nCapital", SqlDbType.Money).Value = MatCalend(i, 4)
                    oComm.Parameters.Add("@nInteres", SqlDbType.Money).Value = MatCalend(i, 6)
                    '************
                    oComm.Parameters.Add("@nGastoFijo", SqlDbType.Money).Value = MatCalend(i, 13)

                    oComm.Parameters.Add("@nMora", SqlDbType.Money).Value = MatCalend(i, 8)
                    oComm.Parameters.Add("@nEstado", SqlDbType.Int).Value = MatCalend(i, 11)
                    oComm.Parameters.Add("@dFecgaPago", SqlDbType.DateTime).Value = dFecPago
                    oComm.Parameters.Add("@nNroTran", SqlDbType.Int).Value = nNroTran
                    oComm.Parameters.Add("@cUser", SqlDbType.VarChar, 50).Value = cUser
                    oComm.Parameters.Add("@nIgv", SqlDbType.Money).Value = MatCalend(i, 10)

                    '
                    oComm.ExecuteNonQuery()

                End If

            Next

            'Aplica Ampliacion
            'oComm = New SqlClient.SqlCommand("PA_ActualizaAplicaAmpliacion", objCon.DameConexion)
            'oComm.CommandType = CommandType.StoredProcedure

            'oComm.Parameters.Add("@pnCodAge", SqlDbType.Int).Value = pnCodAge
            'oComm.Parameters.Add("@pnCodCred", SqlDbType.Int).Value = pnCodCred


            'oComm.ExecuteNonQuery()



            objCon.CierraConexion()

            'Renueva Calendario Para creditos a 1 sola cuota
            If pnMontoPago < nDeudaALAFecha Then
                If nIntPend <= 0 Then
                    oComm = New SqlClient.SqlCommand("Web_RenuevaCalendarioUnaCuota_2", objCon.DameConexion)
                    oComm.CommandType = CommandType.StoredProcedure
                    oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
                    oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
                    oComm.ExecuteNonQuery()

                    objCon.CierraConexion()
                End If

            End If


            Dim pnDeudaAFecha As Double
            Dim dProxfexVen As Date
            pnDeudaAFecha = DeudaCalendario(pnCodAge, pnCodCred)
            dProxfexVen = ProxFechaPago(pnCodAge, pnCodCred)

            ReDim D(2)
            D(0) = "1"
            D(1) = "Pago Correcto, Proxima fecha " + Format(dProxfexVen, "dd/MM/yyyy") + " Deuda : " + Format(pnDeudaAFecha, "#0.00")

            Return D

        Catch ex As Exception
            ReDim D(2)
            D(0) = "0"
            D(1) = "Error : " + Err.Description
            Return D


        End Try


    End Function

    Public Function DeudaCalendario(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Double
        Dim objCon As New clsConecta

        Dim oComm As SqlClient.SqlCommand

        oComm = New SqlClient.SqlCommand("PA_DeudaALAFecha", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@nDeudaALAFecha", SqlDbType.Money).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        DeudaCalendario = oComm.Parameters("@nDeudaALAFecha").Value

        objCon.CierraConexion()

    End Function

    Public Function ProxFechaPago(ByVal pnCodAge As Integer, ByVal pnCodCred As Integer) As Date

        Dim oComm As SqlClient.SqlCommand
        Dim objCon As New clsConecta


        oComm = New SqlClient.SqlCommand("PA_ProxfechaPago", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nCodAge", SqlDbType.Int).Value = pnCodAge
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = pnCodCred
        oComm.Parameters.Add("@dProxfecPago", SqlDbType.Date).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ProxFechaPago = oComm.Parameters("@dProxfecPago").Value

        objCon.CierraConexion()

    End Function

    Public Function ValidaTransaccion() As Integer


        Dim oComm As SqlClient.SqlCommand
        Dim objCon As New clsConecta


        oComm = New SqlClient.SqlCommand("PA_ValidaIDPago", objCon.DameConexion)
        oComm.CommandType = CommandType.StoredProcedure

        oComm.Parameters.Add("@nID", SqlDbType.Int).Value = nId
        oComm.Parameters.Add("@nCodCred", SqlDbType.Int).Value = CInt(sCodCred)
        oComm.Parameters.Add("@nResultado", SqlDbType.Int).Direction = ParameterDirection.Output

        oComm.ExecuteNonQuery()

        ValidaTransaccion = oComm.Parameters("@nResultado").Value

        objCon.CierraConexion()

    End Function

    Public Function RegistrarPago() As outNegocio
        Dim R As New outNegocio
        Dim sResp As Object

        Try

            If ValidaTransaccion() = 1 Or nId = -1 Then
                sResp = Web_PagarCredito(sCodAge.Trim, sCodCred.Trim, sMontoPago.Trim, cUser.Trim, psMetLiquid.Trim, sLat.Trim, sLong.Trim)

                R.nResultado = CInt(sResp(0))
                R.sRespuesta = sResp(1)

            Else
                R.nResultado = 1
                R.sRespuesta = "OK"
            End If

        Catch ex As Exception

            R.nResultado = 0
            'R.sRespuesta = ex.Message
            R.sRespuesta = sCodAge.Trim + "," + sCodCred.Trim + "," + sMontoPago.Trim + "," + cUser.Trim + "," + psMetLiquid.Trim + "," + sLat.Trim + "," + sLong.Trim

        End Try

        Return R

    End Function

    Public Function RegistrarPagoGrupalLote() As outNegocio
        Dim R As New outNegocio
        Dim RData As New inNegocio
        Dim sResp As Object
        Dim nIDPagolote As Long = 0


        Dim oConex As SqlConnection
        Dim oConect As New clsConecta
        Dim sCadConexion As String
        Dim C As SqlCommand
        Dim oComm As SqlClient.SqlCommand
        Dim nResp As Integer
        Dim sRespuesta As String
        Dim dr As SqlClient.SqlDataReader

        nResp = 0
        sRespuesta = ""

        sCadConexion = oConect.gsCadConex

        oConex = New SqlConnection(sCadConexion)

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing


        Try


            oComm = New SqlClient.SqlCommand("PA_RegistraPagolotecab", oConex)
            oComm.CommandTimeout = 0
            oComm.CommandType = CommandType.StoredProcedure
            oComm.Parameters.Add("@nIDGrupo", SqlDbType.Int).Value = CLng(sCodCred)
            oComm.Parameters.Add("@cCodUsu", SqlDbType.VarChar, 50).Value = cUser
            oComm.Parameters.Add("@nIDPagolote", SqlDbType.Int).Direction = ParameterDirection.Output

            oComm.ExecuteNonQuery()

            nIDPagolote = oComm.Parameters("@nIDPagolote").Value


            Dim i As Integer = 0

            For i = 0 To oListaPagoLoteCreditos.Count - 1

                If oListaPagoLoteCreditos(i).nMontoPago > 0 Then

                    sCodAge = Str(oListaPagoLoteCreditos(i).nCodAge)
                    sCodCred = Str(oListaPagoLoteCreditos(i).nCodCred)
                    sMontoPago = Str(oListaPagoLoteCreditos(i).nMontoPago)

                    If ValidaTransaccion() = 1 Or nId = -1 Then
                        sResp = Web_PagarCreditoLote(sCodAge.Trim, sCodCred.Trim, sMontoPago.Trim, cUser.Trim, psMetLiquid.Trim, sLat.Trim, sLong.Trim, nIDPagolote)

                        R.nResultado = CInt(sResp(0))
                        R.sRespuesta = sResp(1)

                    Else
                        R.nResultado = 1
                        R.sRespuesta = "OK"
                    End If

                End If
            Next



            oConex.Close()

        Catch ex As Exception

            R.nResultado = 0
            R.sRespuesta = sCodAge.Trim + "," + sCodCred.Trim + "," + sMontoPago.Trim + "," + cUser.Trim + "," + psMetLiquid.Trim + "," + sLat.Trim + "," + sLong.Trim

        End Try

        Return R

    End Function


End Class
