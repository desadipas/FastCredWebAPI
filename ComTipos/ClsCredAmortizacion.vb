
<Serializable()> Public Class ClsCredAmortizacion

    Public oCuotas As clsDataset
    Public oDatosCredito As clsDataset
    Public oPorcentajes As clsDataset
    Public nMontoPago As Double

    Private nMontoPagoCapital As Double
    Private nMontoPagoInteres As Double
    Private nMontoPagoInteresGracia As Double
    Private nMontoPagoMora As Double
    Private nMontoPagoGastos As Double

    'Valores de deuda a la fecha


    Public Function DeudaCapital() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DeudaCapital = nSuma
    End Function

    Public Function DeudaInteres() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DeudaInteres = nSuma
    End Function

    Public Function DevuelveInteresPagado() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("INTERESPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DevuelveInteresPagado = nSuma
    End Function

    Public Function DeudaInteresGracia() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("INTGRACIA")) - CDbl(oCuotas.ObtenerDato("INTGRACIAPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DeudaInteresGracia = nSuma
    End Function

    Public Function DeudaInteresMoratorio() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DeudaInteresMoratorio = nSuma
    End Function


    Public Function DeudaGastos() As Double
        Dim nSuma As Double
        nSuma = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            nSuma = nSuma + CDbl(Format((CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))), "#,0.00"))
            oCuotas.Siguiente()
        Loop
        DeudaGastos = nSuma
    End Function




    Public Function GenerarPagoVertical(ByVal pnItf As Double, ByVal pnIGV As Double) As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        nMontoRestante = 0
        nDiff = 0
        nPendiente = 0

        nMontoPagoCapital = Format(CDbl(oPorcentajes.ObtenerDato("nPorCap")) * nMontoPago / 100, "#,0.00")
        nMontoPagoInteres = Format(CDbl(oPorcentajes.ObtenerDato("nPorInt")) * nMontoPago / 100, "#,0.00")
        nMontoPagoInteresGracia = Format(CDbl(oPorcentajes.ObtenerDato("nPorcIntGracia")) * nMontoPago / 100, "#,0.00")
        nMontoPagoMora = Format(CDbl(oPorcentajes.ObtenerDato("nPorMora")) * nMontoPago / 100, "#,0.00")
        nMontoPagoGastos = Format(CDbl(oPorcentajes.ObtenerDato("nPorGasto")) * nMontoPago / 100, "#,0.00")


        'Ajustar el monto de pago a los montos de porcentaje obtenidos
        nDiff = (nMontoPagoCapital + nMontoPagoInteres + nMontoPagoInteresGracia + nMontoPagoMora + nMontoPagoGastos) - nMontoPago

        If nDiff > 0 And nMontoPagoGastos > 0 And nMontoPagoGastos > nDiff Then
            nMontoPagoGastos = nMontoPagoGastos - nDiff
            nDiff = 0
        End If

        If nDiff > 0 And nMontoPagoMora > 0 And nMontoPagoMora > nDiff Then
            nMontoPagoMora = nMontoPagoMora - nDiff
        End If

        If nDiff > 0 And nMontoPagoInteresGracia > 0 And nMontoPagoInteresGracia > nDiff Then
            nMontoPagoInteresGracia = nMontoPagoInteresGracia - nDiff
        End If

        If nDiff > 0 And nMontoPagoInteres > 0 And nMontoPagoInteres > nDiff Then
            nMontoPagoInteres = nMontoPagoInteres - nDiff
        End If

        'Amortiza Capital
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                    nMontoPagoCapital = nMontoPagoCapital - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                    nMontoPagoCapital = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Interes + Igv
        If pnItf = 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPagoInteres >= (CDbl(oCuotas.ObtenerDato("INTERES")) * (1 + pnIGV)) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPagoInteres = Format(nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES")), "#,0.00")
                        oCuotas.AsignarDato(Format((CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00"), "IGVPAG")
                        nMontoPagoInteres = Format(nMontoPagoInteres - (CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00")
                    Else
                        oCuotas.AsignarDato(Format((nMontoPagoInteres / (1 + pnIGV)), "#,0.00"), "INTERESPAG")
                        oCuotas.AsignarDato(Format((nMontoPagoInteres - (nMontoPagoInteres / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                        nMontoPagoInteres = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        Else
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPagoInteres >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPagoInteres = nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                        nMontoPagoInteres = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop

        End If



        'Amortizo Interes de Gracia
        If nMontoPagoInteresGracia > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                    If nMontoPagoInteresGracia >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("[INTGRACIA]"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If

        'Amortizo Interes Moratorio
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                If nMontoPagoMora >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                    nMontoPagoMora = nMontoPagoMora - CDbl(oCuotas.ObtenerDato("MORA"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoMora, "#,0.00"), "MORAPAG")
                    nMontoPagoMora = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Gastos
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                If nMontoPagoGastos >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - CDbl(oCuotas.ObtenerDato("GASTOS"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoPagoGastos = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Si algun concepto quedo con sobrante mandar a conceptos que quedaron pendientes
        'Si quedo Sobrante 
        nMontoRestante = nMontoPagoCapital + nMontoPagoInteres + nMontoPagoInteresGracia + nMontoPagoMora + nMontoPagoGastos


        If nMontoRestante > 0 And DeudaCapital() > 0 Then
            nMontoPagoCapital = nMontoRestante

            'Amortiza Capital
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = Format(CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")), "#,0.00")
                If nPendiente > 0 Then
                    If nMontoPagoCapital >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nPendiente, "CAPITALPAG")
                        nMontoPagoCapital = nMontoPagoCapital - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                        nMontoPagoCapital = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoCapital

        End If

        If nMontoRestante > 0 And DeudaInteres() > 0 Then
            nMontoPagoInteres = nMontoRestante

            'Amortizo Interes + Igv
            If pnItf = 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= (nPendiente * (1 + pnIGV)) Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nPendiente * pnIGV), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = Format(nMontoPagoInteres - (nPendiente * pnIGV), "#,0.00")
                        Else
                            oCuotas.AsignarDato(Format((nMontoPagoInteres / (1 + pnIGV)) + CDbl(oCuotas.ObtenerDato("INTERESPAG")), "#,0.00"), "INTERESPAG")
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nMontoPagoInteres - (nMontoPagoInteres / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If
                    oCuotas.Siguiente()
                Loop

            Else

                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= nPendiente Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                        Else
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop

            End If
            nMontoRestante = nMontoPagoInteres

        End If

        If nMontoRestante > 0 And DeudaInteresGracia() > 0 Then
            nMontoPagoInteresGracia = nMontoRestante
            'Amortizo Interes de Gracia
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("INTGRACIA")) - CDbl(oCuotas.ObtenerDato("INTGRACIAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoInteresGracia >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nPendiente, "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoInteresGracia

        End If

        If nMontoRestante > 0 And DeudaInteresMoratorio() > 0 Then
            nMontoPagoMora = nMontoRestante
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoMora >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nPendiente, "MORAPAG")
                        nMontoPagoMora = nMontoPagoMora - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nMontoPagoMora, "#,0.00"), "MORAPAG")
                        nMontoPagoMora = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoMora
        End If


        If nMontoRestante > 0 And DeudaGastos() > 0 Then
            nMontoPagoGastos = nMontoRestante
            'Amortizo Gastos
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))
                If nMontoPagoGastos >= nPendiente Then
                    oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nPendiente, "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - nPendiente
                Else
                    oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoPagoGastos = 0
                    Exit Do
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoGastos

        End If

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")


            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoVertical = nMontoRestante



    End Function


    Public Function GenerarPagoTodoCapital(ByVal pnItf As Double, ByVal pnIGV As Double) As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        nMontoRestante = 0
        nDiff = 0
        nPendiente = 0



        'Amortiza Capital
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                If nMontoPago >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                    nMontoPago = nMontoPago - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPago, "#,0.00"), "CAPITALPAG")
                    nMontoPago = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Interes + Igv
        If pnItf = 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPago >= (CDbl(oCuotas.ObtenerDato("INTERES")) * (1 + pnIGV)) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPago = Format(nMontoPago - CDbl(oCuotas.ObtenerDato("INTERES")), "#,0.00")
                        oCuotas.AsignarDato(Format((CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00"), "IGVPAG")
                        nMontoPago = Format(nMontoPago - (CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00")
                    Else
                        oCuotas.AsignarDato(Format((nMontoPago / (1 + pnIGV)), "#,0.00"), "INTERESPAG")
                        oCuotas.AsignarDato(Format((nMontoPago - (nMontoPago / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                        nMontoPago = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        Else
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPago >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPago = nMontoPago - CDbl(oCuotas.ObtenerDato("INTERES"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPago, "#,0.00"), "INTERESPAG")
                        nMontoPago = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop

        End If

        'Amortizo Interes de Gracia
        If nMontoPago > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                    If nMontoPago >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("[INTGRACIA]"), "INTGRACIAPAG")
                        nMontoPago = nMontoPago - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPago, "#,0.00"), "INTGRACIAPAG")
                        nMontoPago = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If

        'Amortizo Interes Moratorio
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                If nMontoPago >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                    nMontoPago = nMontoPago - CDbl(oCuotas.ObtenerDato("MORA"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPago, "#,0.00"), "MORAPAG")
                    nMontoPago = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Gastos
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                If nMontoPago >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPago = nMontoPago - CDbl(oCuotas.ObtenerDato("GASTOS"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPago, "#,0.00"), "GASTOSPAG")
                    nMontoPago = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Si algun concepto quedo con sobrante mandar a conceptos que quedaron pendientes
        'Si quedo Sobrante 
        nMontoRestante = nMontoPago


        If nMontoRestante > 0 And DeudaCapital() > 0 Then
            nMontoPagoCapital = nMontoRestante

            'Amortiza Capital
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = Format(CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")), "#,0.00")
                If nPendiente > 0 Then
                    If nMontoPagoCapital >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nPendiente, "CAPITALPAG")
                        nMontoPagoCapital = nMontoPagoCapital - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                        nMontoPagoCapital = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoCapital

        End If

        If nMontoRestante > 0 And DeudaInteres() > 0 Then
            nMontoPagoInteres = nMontoRestante

            'Amortizo Interes + Igv
            If pnItf = 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= (nPendiente * (1 + pnIGV)) Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nPendiente * pnIGV), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = Format(nMontoPagoInteres - (nPendiente * pnIGV), "#,0.00")
                        Else
                            oCuotas.AsignarDato(Format((nMontoPagoInteres / (1 + pnIGV)) + CDbl(oCuotas.ObtenerDato("INTERESPAG")), "#,0.00"), "INTERESPAG")
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nMontoPagoInteres - (nMontoPagoInteres / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If
                    oCuotas.Siguiente()
                Loop

            Else

                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= nPendiente Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                        Else
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop

            End If
            nMontoRestante = nMontoPagoInteres

        End If

        If nMontoRestante > 0 And DeudaInteresGracia() > 0 Then
            nMontoPagoInteresGracia = nMontoRestante
            'Amortizo Interes de Gracia
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("INTGRACIA")) - CDbl(oCuotas.ObtenerDato("INTGRACIAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoInteresGracia >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nPendiente, "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoInteresGracia

        End If

        If nMontoRestante > 0 And DeudaInteresMoratorio() > 0 Then
            nMontoPagoMora = nMontoRestante
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoMora >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nPendiente, "MORAPAG")
                        nMontoPagoMora = nMontoPagoMora - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nMontoPagoMora, "#,0.00"), "MORAPAG")
                        nMontoPagoMora = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoMora
        End If


        If nMontoRestante > 0 And DeudaGastos() > 0 Then
            nMontoPagoGastos = nMontoRestante
            'Amortizo Gastos
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))
                If nMontoPagoGastos >= nPendiente Then
                    oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nPendiente, "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - nPendiente
                Else
                    oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoPagoGastos = 0
                    Exit Do
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoGastos

        End If

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")


            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoTodoCapital = nMontoRestante



    End Function

    Public Function DevuelveDeudaAFecha(ByVal pdFechaHoy As DateTime, ByVal pnIgv As Double, ByVal pnITF As Double) As Double
        Dim nInteresAFecha As Double
        Dim dFecInicioCuota As DateTime
        Dim nDiasTrans As Integer

        dFecInicioCuota = CDate(oDatosCredito.ObtenerDato("dFechaInicioCuota"))
        If dFecInicioCuota = CDate(oDatosCredito.ObtenerDato("dFecVig")) Then
            dFecInicioCuota = dFecInicioCuota.AddDays(CInt(oDatosCredito.ObtenerDato("nNroDiasGracia")))
        End If

        nInteresAFecha = 0

        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            'Si la cuota aun no ha vencido
            If pdFechaHoy < CDate(oCuotas.ObtenerDato("VENCIMIENTO")) Then
                nDiasTrans = DateDiff(DateInterval.Day, dFecInicioCuota, pdFechaHoy)
                If nDiasTrans > 0 Then
                    nInteresAFecha = nInteresAFecha + Format((((1 + (CDbl(oDatosCredito.ObtenerDato("nTasaComp")) / 100)) ^ (nDiasTrans / 30)) - 1) * CDbl(oCuotas.ObtenerDato("nSaldoCapCuota")), "#,0.00")
                End If
                Exit Do
            Else
                'Si la cuota ya venció
                nInteresAFecha = nInteresAFecha + CDbl(oCuotas.ObtenerDato("INTERES"))
            End If
            dFecInicioCuota = CDate(oCuotas.ObtenerDato("VENCIMIENTO"))
            oCuotas.Siguiente()
        Loop

        nInteresAFecha = Format(nInteresAFecha - DevuelveInteresPagado(), "#,0.00")
        If nInteresAFecha < 0 Then
            nInteresAFecha = 0
        End If

        DevuelveDeudaAFecha = nInteresAFecha + DeudaInteresGracia() + DeudaInteresMoratorio() + DeudaGastos() + DeudaCapital()
        If pnITF = 0 Then
            DevuelveDeudaAFecha = Format(DevuelveDeudaAFecha + CDbl(Format(((nInteresAFecha * (1 + pnIgv)) - nInteresAFecha), "#,0.00")), "#,0.00")
        Else
            DevuelveDeudaAFecha = Format(DevuelveDeudaAFecha * pnITF, "#,0.00")
        End If


    End Function

    Public Function GenerarPagoNotaCreditoCuotas(ByVal pnItf As Double, ByVal pnIGV As Double, _
    ByVal pnMontoCapital As Double, ByVal pnMontoIntComp As Double, ByVal pnMontoIntMor As Double, ByVal pnMontoIntGra As Double, ByVal pnMontoGastos As Double) As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        nMontoRestante = 0
        nDiff = 0
        nPendiente = 0

        nMontoPagoCapital = pnMontoCapital
        nMontoPagoInteres = pnMontoIntComp
        nMontoPagoInteresGracia = pnMontoIntGra
        nMontoPagoMora = pnMontoIntMor
        nMontoPagoGastos = pnMontoGastos


        'Ajustar el monto de pago a los montos de porcentaje obtenidos
        nDiff = (nMontoPagoCapital + nMontoPagoInteres + nMontoPagoInteresGracia + nMontoPagoMora + nMontoPagoGastos) - nMontoPago

        If nDiff > 0 And nMontoPagoGastos > 0 And nMontoPagoGastos > nDiff Then
            nMontoPagoGastos = nMontoPagoGastos - nDiff
            nDiff = 0
        End If

        If nDiff > 0 And nMontoPagoMora > 0 And nMontoPagoMora > nDiff Then
            nMontoPagoMora = nMontoPagoMora - nDiff
        End If

        If nDiff > 0 And nMontoPagoInteresGracia > 0 And nMontoPagoInteresGracia > nDiff Then
            nMontoPagoInteresGracia = nMontoPagoInteresGracia - nDiff
        End If

        If nDiff > 0 And nMontoPagoInteres > 0 And nMontoPagoInteres > nDiff Then
            nMontoPagoInteres = nMontoPagoInteres - nDiff
        End If

        'Amortiza Capital
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                    nMontoPagoCapital = nMontoPagoCapital - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                    nMontoPagoCapital = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Interes + Igv
        If pnItf = 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPagoInteres >= (CDbl(oCuotas.ObtenerDato("INTERES")) * (1 + pnIGV)) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPagoInteres = Format(nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES")), "#,0.00")
                        oCuotas.AsignarDato(Format((CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00"), "IGVPAG")
                        nMontoPagoInteres = Format(nMontoPagoInteres - (CDbl(oCuotas.ObtenerDato("INTERES")) * pnIGV), "#,0.00")
                    Else
                        oCuotas.AsignarDato(Format((nMontoPagoInteres / (1 + pnIGV)), "#,0.00"), "INTERESPAG")
                        oCuotas.AsignarDato(Format((nMontoPagoInteres - (nMontoPagoInteres / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                        nMontoPagoInteres = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        Else
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPagoInteres >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPagoInteres = nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                        nMontoPagoInteres = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop

        End If



        'Amortizo Interes de Gracia
        If nMontoPagoInteresGracia > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                    If nMontoPagoInteresGracia >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("[INTGRACIA]"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If

        'Amortizo Interes Moratorio
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                If nMontoPagoMora >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                    nMontoPagoMora = nMontoPagoMora - CDbl(oCuotas.ObtenerDato("MORA"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoMora, "#,0.00"), "MORAPAG")
                    nMontoPagoMora = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Gastos
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                If nMontoPagoGastos >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - CDbl(oCuotas.ObtenerDato("GASTOS"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoPagoGastos = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Si algun concepto quedo con sobrante mandar a conceptos que quedaron pendientes
        'Si quedo Sobrante 
        nMontoRestante = nMontoPagoCapital + nMontoPagoInteres + nMontoPagoInteresGracia + nMontoPagoMora + nMontoPagoGastos


        If nMontoRestante > 0 And DeudaCapital() > 0 Then
            nMontoPagoCapital = nMontoRestante

            'Amortiza Capital
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = Format(CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")), "#,0.00")
                If nPendiente > 0 Then
                    If nMontoPagoCapital >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nPendiente, "CAPITALPAG")
                        nMontoPagoCapital = nMontoPagoCapital - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                        nMontoPagoCapital = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoCapital

        End If

        If nMontoRestante > 0 And DeudaInteres() > 0 Then
            nMontoPagoInteres = nMontoRestante

            'Amortizo Interes + Igv
            If pnItf = 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= (nPendiente * (1 + pnIGV)) Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nPendiente * pnIGV), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = Format(nMontoPagoInteres - (nPendiente * pnIGV), "#,0.00")
                        Else
                            oCuotas.AsignarDato(Format((nMontoPagoInteres / (1 + pnIGV)) + CDbl(oCuotas.ObtenerDato("INTERESPAG")), "#,0.00"), "INTERESPAG")
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("IGVPAG")) + (nMontoPagoInteres - (nMontoPagoInteres / (1 + pnIGV))), "#,0.00"), "IGVPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If
                    oCuotas.Siguiente()
                Loop

            Else

                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    nPendiente = (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    If nPendiente > 0 Then
                        If nMontoPagoInteres >= nPendiente Then
                            oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nPendiente, "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - nPendiente
                        Else
                            oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop

            End If
            nMontoRestante = nMontoPagoInteres

        End If

        If nMontoRestante > 0 And DeudaInteresGracia() > 0 Then
            nMontoPagoInteresGracia = nMontoRestante
            'Amortizo Interes de Gracia
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("INTGRACIA")) - CDbl(oCuotas.ObtenerDato("INTGRACIAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoInteresGracia >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nPendiente, "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) + nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoInteresGracia

        End If

        If nMontoRestante > 0 And DeudaInteresMoratorio() > 0 Then
            nMontoPagoMora = nMontoRestante
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG"))
                If nPendiente > 0 Then
                    If nMontoPagoMora >= nPendiente Then
                        oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nPendiente, "MORAPAG")
                        nMontoPagoMora = nMontoPagoMora - nPendiente
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nMontoPagoMora, "#,0.00"), "MORAPAG")
                        nMontoPagoMora = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoMora
        End If


        If nMontoRestante > 0 And DeudaGastos() > 0 Then
            nMontoPagoGastos = nMontoRestante
            'Amortizo Gastos
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                nPendiente = CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))
                If nMontoPagoGastos >= nPendiente Then
                    oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nPendiente, "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - nPendiente
                Else
                    oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoPagoGastos = 0
                    Exit Do
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoPagoGastos

        End If

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")


            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoNotaCreditoCuotas = nMontoRestante



    End Function

    Public Sub New()
        oCuotas = New clsDataset
        oPorcentajes = New clsDataset

    End Sub

    Public Function GenerarPagoPorcentual() As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        Dim nMontoCapitalPagado As Double = 0
        Dim nMontoInteresPagado As Double = 0
        Dim nMontoInteresGraciaPagado As Double = 0
        Dim nMontoMoraPagado As Double = 0
        Dim nMontoGastosPagado As Double = 0

        nMontoRestante = nMontoPago
        nDiff = 0
        nPendiente = 0

        'Primero recorrer Capital + Gastos
        nMontoPagoCapital = Format(90 * nMontoRestante / 100, "#,0.00")
        nMontoPagoGastos = Format(10 * nMontoRestante / 100, "#,0.00")

        'Amortiza Capital
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                    nMontoPagoCapital = nMontoPagoCapital - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                    nMontoCapitalPagado = nMontoCapitalPagado + CDbl(oCuotas.ObtenerDato("CAPITAL"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                    nMontoCapitalPagado = nMontoCapitalPagado + nMontoPagoCapital
                    nMontoPagoCapital = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Gastos
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                If nMontoPagoGastos >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - CDbl(oCuotas.ObtenerDato("GASTOS"))
                    nMontoGastosPagado = nMontoGastosPagado + CDbl(oCuotas.ObtenerDato("GASTOS"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoGastosPagado = nMontoGastosPagado + nMontoPagoGastos
                    nMontoPagoGastos = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Si algun concepto quedo con sobrante mandar a conceptos que quedaron pendientes
        'Si quedo Sobrante 
        nMontoRestante = nMontoPago - (nMontoCapitalPagado + nMontoGastosPagado)


        'Luego recorrer Interes + Gastos
        nMontoPagoInteres = Format(90 * nMontoRestante / 100, "#,0.00")
        nMontoPagoInteresGracia = Format(90 * nMontoRestante / 100, "#,0.00")
        nMontoPagoGastos = Format(10 * nMontoRestante / 100, "#,0.00")


        'Amortizo Interes 
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                If nMontoPagoInteres >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                    nMontoPagoInteres = nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES"))
                    nMontoInteresPagado = nMontoInteresPagado + CDbl(oCuotas.ObtenerDato("INTERES"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                    nMontoInteresPagado = nMontoInteresPagado + nMontoPagoInteres
                    nMontoPagoInteres = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Interes de Gracia
        If nMontoPagoInteresGracia > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                    If nMontoPagoInteresGracia >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTGRACIA"), "INTGRACIAPAG")
                        nMontoPagoInteresGracia = nMontoPagoInteresGracia - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                        nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                        nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + nMontoPagoInteresGracia
                        nMontoPagoInteresGracia = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If

        'Amortizo Gastos
        oCuotas.Inicio()
        nMontoGastosPagado = 0
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")) > 0 Then
                If nMontoPagoGastos >= (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")))
                    nMontoGastosPagado = nMontoGastosPagado + (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")))
                Else
                    'oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    oCuotas.AsignarDato(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoPagoGastos, "GASTOSPAG")
                    nMontoGastosPagado = nMontoGastosPagado + nMontoPagoGastos
                    nMontoPagoGastos = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        nMontoRestante = nMontoRestante - (nMontoInteresPagado + nMontoInteresGraciaPagado + nMontoGastosPagado)

        'Finalmente Mora + Gastos
        nMontoPagoMora = Format(90 * nMontoRestante / 100, "#,0.00")
        nMontoPagoGastos = Format(10 * nMontoRestante / 100, "#,0.00")

        'Amortizo Interes Moratorio
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                If nMontoPagoMora >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                    nMontoPagoMora = nMontoPagoMora - CDbl(oCuotas.ObtenerDato("MORA"))
                    nMontoMoraPagado = nMontoMoraPagado + CDbl(oCuotas.ObtenerDato("MORA"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoMora, "#,0.00"), "MORAPAG")
                    nMontoMoraPagado = nMontoMoraPagado + nMontoPagoMora
                    nMontoPagoMora = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        'Amortizo Gastos
        oCuotas.Inicio()
        nMontoGastosPagado = 0
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")) > 0 Then
                If nMontoPagoGastos >= (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                    nMontoPagoGastos = nMontoPagoGastos - (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")))
                    nMontoGastosPagado = nMontoGastosPagado + (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")))
                Else
                    'oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                    nMontoGastosPagado = nMontoGastosPagado + nMontoPagoGastos
                    nMontoPagoGastos = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop

        nMontoRestante = nMontoRestante - (nMontoMoraPagado + nMontoGastosPagado)
        nMontoRestante = Format(nMontoRestante, "#,0.00")

        '----------- Si quedo sobrante amortizar a Gastos ------------------
        If nMontoRestante > 0 Then
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")) > 0 Then
                    If nMontoRestante >= (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG"))) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                        nMontoRestante = nMontoRestante - (CDbl(oCuotas.ObtenerDato("GASTOS")) - CDbl(oCuotas.ObtenerDato("GASTOSPAG")))
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("GASTOSPAG")) + nMontoRestante, "#,0.00"), "GASTOSPAG")
                        nMontoRestante = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        '---------------Si quedo sobrante amortizar a Mora -----------------------------
        If nMontoRestante > 0 Then
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG")) > 0 Then
                    If nMontoRestante >= (CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG"))) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                        nMontoRestante = nMontoRestante - (CDbl(oCuotas.ObtenerDato("MORA")) - CDbl(oCuotas.ObtenerDato("MORAPAG")))
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("MORAPAG")) + nMontoRestante, "#,0.00"), "MORAPAG")
                        nMontoRestante = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        '---------------Si quedo sobrante amortizar a Interes -----------------------------
        If nMontoRestante > 0 Then
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")) > 0 Then
                    If nMontoRestante >= (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG"))) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoRestante = nMontoRestante - (CDbl(oCuotas.ObtenerDato("INTERES")) - CDbl(oCuotas.ObtenerDato("INTERESPAG")))
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("INTERESPAG")) + nMontoRestante, "#,0.00"), "INTERESPAG")
                        nMontoRestante = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        '---------------Si quedo sobrante amortizar a Capital -----------------------------
        If nMontoRestante > 0 Then
            'Amortizo Interes Moratorio
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")) > 0 Then
                    If nMontoRestante >= (CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG"))) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                        nMontoRestante = nMontoRestante - (CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")))
                    Else
                        oCuotas.AsignarDato(Format(CDbl(oCuotas.ObtenerDato("CAPITALPAG")) + nMontoRestante, "#,0.00"), "CAPITALPAG")
                        nMontoRestante = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        '----------------------------------------------------------------------------------

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")

            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoPorcentual = nMontoRestante



    End Function

    Public Function GenerarPagoCastigado() As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        Dim nMontoCapitalPagado As Double = 0
        Dim nMontoInteresPagado As Double = 0
        Dim nMontoInteresGraciaPagado As Double = 0
        Dim nMontoMoraPagado As Double = 0
        Dim nMontoGastosPagado As Double = 0

        nMontoRestante = nMontoPago
        nDiff = 0
        nPendiente = 0

        'Primero recorrer Capital 
        If DeudaCapital() > 0 Then
            nMontoPagoCapital = IIf(nMontoRestante > DeudaCapital(), DeudaCapital(), nMontoRestante)

            'Amortiza Capital
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                    If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                        'If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) - CDbl(oCuotas.ObtenerDato("CAPITALPAG")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                        nMontoPagoCapital = nMontoPagoCapital - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                        nMontoCapitalPagado = nMontoCapitalPagado + CDbl(oCuotas.ObtenerDato("CAPITAL"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                        nMontoCapitalPagado = nMontoCapitalPagado + nMontoPagoCapital
                        nMontoPagoCapital = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
            nMontoRestante = nMontoRestante - nMontoCapitalPagado
        End If

        If DeudaInteres() > 0 Then
            nMontoPagoInteres = IIf(nMontoRestante > DeudaInteres(), DeudaInteres(), nMontoRestante)

            'Amortizo Interes 
            If nMontoRestante > 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                        If nMontoPagoInteres >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                            oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                            nMontoPagoInteres = nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES"))
                            nMontoInteresPagado = nMontoInteresPagado + CDbl(oCuotas.ObtenerDato("INTERES"))
                        Else
                            oCuotas.AsignarDato(Format(nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                            nMontoInteresPagado = nMontoInteresPagado + nMontoPagoInteres
                            nMontoPagoInteres = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop
            End If
            nMontoRestante = nMontoRestante - nMontoInteresPagado
        End If

        If DeudaInteresGracia() > 0 Then
            nMontoPagoInteresGracia = IIf(nMontoRestante > DeudaInteresGracia(), DeudaInteresGracia(), nMontoRestante)

            'Amortizo Interes de Gracia
            If nMontoRestante > 0 Then
                If nMontoPagoInteresGracia > 0 Then
                    oCuotas.Inicio()
                    Do While Not oCuotas.EOF
                        If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                            If nMontoPagoInteresGracia >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                                oCuotas.AsignarDato(oCuotas.ObtenerDato("[INTGRACIA]"), "INTGRACIAPAG")
                                nMontoPagoInteresGracia = nMontoPagoInteresGracia - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                                nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                            Else
                                oCuotas.AsignarDato(Format(nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                                nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + nMontoPagoInteresGracia
                                nMontoPagoInteresGracia = 0
                                Exit Do
                            End If
                        End If

                        oCuotas.Siguiente()
                    Loop
                End If
            End If
            nMontoRestante = nMontoRestante - nMontoInteresGraciaPagado
        End If

        If DeudaGastos() > 0 Then
            nMontoPagoGastos = IIf(nMontoRestante > DeudaGastos(), DeudaGastos(), nMontoRestante)

            'Amortizo Gastos
            If nMontoRestante > 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                        If nMontoPagoGastos >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                            oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                            nMontoPagoGastos = nMontoPagoGastos - CDbl(oCuotas.ObtenerDato("GASTOS"))
                            nMontoGastosPagado = nMontoGastosPagado + CDbl(oCuotas.ObtenerDato("GASTOS"))
                        Else
                            oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                            nMontoGastosPagado = nMontoGastosPagado + nMontoPagoGastos
                            nMontoPagoGastos = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop
            End If
            nMontoRestante = nMontoRestante - nMontoGastosPagado
        End If

        If DeudaInteresMoratorio() > 0 Then
            nMontoPagoMora = IIf(nMontoRestante > DeudaInteresMoratorio(), DeudaInteresMoratorio(), nMontoRestante)

            'Amortizo Interes Moratorio
            If nMontoRestante > 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                        If nMontoPagoMora >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                            oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                            nMontoPagoMora = nMontoPagoMora - CDbl(oCuotas.ObtenerDato("MORA"))
                            nMontoMoraPagado = nMontoMoraPagado + CDbl(oCuotas.ObtenerDato("MORA"))
                        Else
                            oCuotas.AsignarDato(Format(nMontoPagoMora, "#,0.00"), "MORAPAG")
                            nMontoMoraPagado = nMontoMoraPagado + nMontoPagoMora
                            nMontoPagoMora = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop
            End If
            nMontoRestante = nMontoRestante - nMontoMoraPagado
        End If

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")

            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoCastigado = nMontoRestante

    End Function

    Public Function GenerarPagoVertical_NEW() As Double
        Dim nDiff As Double
        Dim nMontoRestante As Double
        Dim nPendiente As Double

        Dim nMontoCapitalPagado As Double = 0
        Dim nMontoInteresPagado As Double = 0
        Dim nMontoInteresGraciaPagado As Double = 0
        Dim nMontoMoraPagado As Double = 0
        Dim nMontoGastosPagado As Double = 0

        nMontoRestante = nMontoPago
        nDiff = 0
        nPendiente = 0

        'Primero recorrer Capital 
        nMontoPagoCapital = IIf(nMontoRestante > DeudaCapital(), nMontoRestante - DeudaCapital(), nMontoRestante)

        'Amortiza Capital
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) > 0 Then
                If nMontoPagoCapital >= CDbl(oCuotas.ObtenerDato("CAPITAL")) Then
                    oCuotas.AsignarDato(oCuotas.ObtenerDato("CAPITAL"), "CAPITALPAG")
                    nMontoPagoCapital = nMontoPagoCapital - CDbl(oCuotas.ObtenerDato("CAPITAL"))
                    nMontoCapitalPagado = nMontoCapitalPagado + CDbl(oCuotas.ObtenerDato("CAPITAL"))
                Else
                    oCuotas.AsignarDato(Format(nMontoPagoCapital, "#,0.00"), "CAPITALPAG")
                    nMontoCapitalPagado = nMontoCapitalPagado + nMontoPagoCapital
                    nMontoPagoCapital = 0
                    Exit Do
                End If
            End If

            oCuotas.Siguiente()
        Loop
        nMontoRestante = nMontoRestante - nMontoCapitalPagado

        nMontoPagoGastos = IIf(nMontoRestante > DeudaGastos(), nMontoRestante - DeudaGastos(), nMontoRestante)

        'Amortizo Gastos
        If nMontoRestante > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("GASTOS")) > 0 Then
                    If nMontoPagoGastos >= CDbl(oCuotas.ObtenerDato("GASTOS")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("GASTOS"), "GASTOSPAG")
                        nMontoPagoGastos = nMontoPagoGastos - CDbl(oCuotas.ObtenerDato("GASTOS"))
                        nMontoGastosPagado = nMontoGastosPagado + CDbl(oCuotas.ObtenerDato("GASTOS"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoGastos, "#,0.00"), "GASTOSPAG")
                        nMontoGastosPagado = nMontoGastosPagado + nMontoPagoGastos
                        nMontoPagoGastos = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        nMontoRestante = nMontoRestante - nMontoGastosPagado

        nMontoPagoInteres = IIf(nMontoRestante > DeudaInteres(), nMontoRestante - DeudaInteres(), nMontoRestante)

        'Amortizo Interes 
        If nMontoRestante > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("INTERES")) > 0 Then
                    If nMontoPagoInteres >= CDbl(oCuotas.ObtenerDato("INTERES")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("INTERES"), "INTERESPAG")
                        nMontoPagoInteres = nMontoPagoInteres - CDbl(oCuotas.ObtenerDato("INTERES"))
                        nMontoInteresPagado = nMontoInteresPagado + CDbl(oCuotas.ObtenerDato("INTERES"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoInteres, "#,0.00"), "INTERESPAG")
                        nMontoInteresPagado = nMontoInteresPagado + nMontoPagoInteres
                        nMontoPagoInteres = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        nMontoRestante = nMontoRestante - nMontoInteresPagado

        nMontoPagoInteresGracia = IIf(nMontoRestante > DeudaInteresGracia(), nMontoRestante - DeudaInteresGracia(), nMontoRestante)

        'Amortizo Interes de Gracia
        If nMontoRestante > 0 Then
            If nMontoPagoInteresGracia > 0 Then
                oCuotas.Inicio()
                Do While Not oCuotas.EOF
                    If CDbl(oCuotas.ObtenerDato("INTGRACIA")) > 0 Then
                        If nMontoPagoInteresGracia >= CDbl(oCuotas.ObtenerDato("INTGRACIA")) Then
                            oCuotas.AsignarDato(oCuotas.ObtenerDato("[INTGRACIA]"), "INTGRACIAPAG")
                            nMontoPagoInteresGracia = nMontoPagoInteresGracia - CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                            nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + CDbl(oCuotas.ObtenerDato("INTGRACIA"))
                        Else
                            oCuotas.AsignarDato(Format(nMontoPagoInteresGracia, "#,0.00"), "INTGRACIAPAG")
                            nMontoInteresGraciaPagado = nMontoInteresGraciaPagado + nMontoPagoInteresGracia
                            nMontoPagoInteresGracia = 0
                            Exit Do
                        End If
                    End If

                    oCuotas.Siguiente()
                Loop
            End If
        End If
        nMontoRestante = nMontoRestante - nMontoInteresGraciaPagado

        nMontoPagoMora = IIf(nMontoRestante > DeudaInteresMoratorio(), nMontoRestante - DeudaInteresMoratorio(), nMontoRestante)

        'Amortizo Interes Moratorio
        If nMontoRestante > 0 Then
            oCuotas.Inicio()
            Do While Not oCuotas.EOF
                If CDbl(oCuotas.ObtenerDato("MORA")) > 0 Then
                    If nMontoPagoMora >= CDbl(oCuotas.ObtenerDato("MORA")) Then
                        oCuotas.AsignarDato(oCuotas.ObtenerDato("MORA"), "MORAPAG")
                        nMontoPagoMora = nMontoPagoMora - CDbl(oCuotas.ObtenerDato("MORA"))
                        nMontoMoraPagado = nMontoMoraPagado + CDbl(oCuotas.ObtenerDato("MORA"))
                    Else
                        oCuotas.AsignarDato(Format(nMontoPagoMora, "#,0.00"), "MORAPAG")
                        nMontoMoraPagado = nMontoMoraPagado + nMontoPagoMora
                        nMontoPagoMora = 0
                        Exit Do
                    End If
                End If

                oCuotas.Siguiente()
            Loop
        End If
        nMontoRestante = nMontoRestante - nMontoMoraPagado

        'Actualiza Estado de Cuota
        oCuotas.Inicio()
        Do While Not oCuotas.EOF
            If CDbl(oCuotas.ObtenerDato("CAPITAL")) = CDbl(oCuotas.ObtenerDato("CAPITALPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTERES")) = CDbl(oCuotas.ObtenerDato("INTERESPAG")) _
                And CDbl(oCuotas.ObtenerDato("INTGRACIA")) = CDbl(oCuotas.ObtenerDato("INTGRACIAPAG")) _
                And CDbl(oCuotas.ObtenerDato("MORA")) = CDbl(oCuotas.ObtenerDato("MORAPAG")) _
                And CDbl(oCuotas.ObtenerDato("GASTOS")) = CDbl(oCuotas.ObtenerDato("GASTOSPAG")) Then

                oCuotas.AsignarDato("CANCELADO", "ESTADO")

            Else

                oCuotas.AsignarDato("PENDIENTE", "ESTADO")

            End If

            oCuotas.Siguiente()
        Loop

        GenerarPagoVertical_NEW = nMontoRestante

    End Function

End Class
