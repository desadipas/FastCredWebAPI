Imports System.Data
Imports System.Data.SqlClient

<Serializable()> Public Class ClsCalendario
    Public oFeriados As ComTipos.clsDataset
    Dim nGasto As Double

    'Public Function GeneraMontoCuota_y_PrimFechaPago(ByVal pnGasto As Double, ByVal pnPrestamo As Double, _
    '         ByVal pnCuotas As Integer, ByVal pdFecDes As Date, ByVal pnDia As Integer, _
    '         ByVal pnDiasGracia As Integer, ByVal pnTasa As Double, ByRef pdFecPrimPago As Date, _
    '         ByVal pnTipoGracia As Integer) As Double

    '    Dim MatDatos(pnCuotas, 8) As String
    '    Dim nMonto As Double
    '    Dim dFecha As Date
    '    Dim MatFechas As Object
    '    Dim nGasto As Double
    '    Dim nTasaGasto As Double
    '    Dim nIGV As Double
    '    Dim nFinCuotaGracia As Integer

    '    nTasaGasto = Format(pnGasto / pnPrestamo * 100, "#0.00")
    '    nIGV = 19
    '    nGasto = Format(pnGasto / pnCuotas, "#0.00")

    '    nMonto = pnPrestamo
    '    dFecha = pdFecDes
    '    MatFechas = CalendarioFechas(pdFecDes, pnDia, pnCuotas, pnTipoGracia, pnDiasGracia, nFinCuotaGracia)
    '    pdFecPrimPago = CDate(MatFechas(1))
    '    GeneraMontoCuota_y_PrimFechaPago = CuotaFechaFija(pdFecDes, pnTasa, nIGV, pnDiasGracia, pnCuotas, MatFechas, pnPrestamo, nFinCuotaGracia)

    'End Function

    Public Function CalculaInteresComp(ByVal pnTasa As Double, ByVal pnDiasTrans As Integer, ByVal pnMonto As Double) As Double
        CalculaInteresComp = ((1 + pnTasa / 100) ^ (pnDiasTrans / 30) - 1) * pnMonto
    End Function

    'Public Function GeneraMontoCuota(ByVal pnGasto As Double, ByVal pnPrestamo As Double, _
    '         ByVal pnCuotas As Integer, ByVal pdFecDes As Date, ByVal pnDia As Integer, _
    '         ByVal pnDiasGracia As Integer, ByVal pnTasa As Double) As Double

    '    Dim MatDatos(pnCuotas, 8) As String
    '    Dim nMonto As Double
    '    Dim dFecha As Date
    '    Dim MatFechas As Object
    '    Dim nGasto As Double
    '    Dim nTasaGasto As Double
    '    Dim nIGV As Double

    '    nTasaGasto = Format(pnGasto / pnPrestamo * 100, "#0.00")
    '    nIGV = 19
    '    nGasto = Format(pnGasto / pnCuotas, "#0.00")

    '    nMonto = pnPrestamo
    '    dFecha = pdFecDes
    '    MatFechas = CalendarioFechas(pdFecDes, pnDia, pnCuotas)
    '    GeneraMontoCuota = CuotaFechaFija(pdFecDes, pnTasa, nIGV, pnDiasGracia, pnCuotas, MatFechas, pnPrestamo)

    'End Function

    Public Function GeneraCalendarioDataSet(ByVal pnMontoGasto As String, ByVal poFeriado As ComTipos.clsDataset, ByVal pnCodAge As Integer, ByVal pnCodCred As Integer, ByVal pnNroCalen As Integer, ByVal pnPrestamo As Double, _
        ByVal pnCuotas As Integer, ByVal pnTasa As Double, ByVal pnDia As Integer, _
        ByVal pdFecDes As Date, ByVal pnDiasGracia As Integer, _
        ByVal pnTipoGracia As Integer, ByVal pnGasto As Double, ByVal pnProd As Integer, _
        ByVal pnSubProd As Integer, ByVal pnIncluyeGasto As Boolean, ByVal pnTipoPeriodo As Integer, ByVal pnPlazo As Integer, ByVal pnItfValor As Double, Optional ByVal pnSoloIntereses As Boolean = False, _
        Optional ByVal pnTasaDesgravamen As Double = 0, Optional ByRef pdFechaPrimPago As Date = Nothing _
        , Optional ByVal pbCuotaDoble As Boolean = False, Optional ByVal bCDPrimeraGratif As Boolean = True) As ComTipos.clsDataset ' ARCV pnTasaDesgravamen (21-05-2009) ---vvillafranca 05-11-2012

        Dim oCal As Object
        Dim i As Integer
        Dim D As ComTipos.clsDataset
        Dim Columnas(18) As String
        Dim nCursor As Integer

        oCal = GeneraCalendario(pnMontoGasto, poFeriado, pnPrestamo, pnCuotas, pnTasa, pnDia, _
                        pdFecDes, pnDiasGracia, pnTipoGracia, pnGasto, pnProd, pnSubProd, pnIncluyeGasto, pnTipoPeriodo, pnPlazo, pnItfValor, pnSoloIntereses, _
                        pnTasaDesgravamen, , , , pdFechaPrimPago, , pbCuotaDoble, bCDPrimeraGratif)

        Columnas(1) = "nCodCred"
        Columnas(2) = "nCodAge"
        Columnas(3) = "nNroCuota"
        Columnas(4) = "dFecVenc"
        Columnas(5) = "dFecPago"
        Columnas(6) = "nCapital"
        Columnas(7) = "nIntComp"
        Columnas(8) = "nIntMor"
        Columnas(9) = "nIgv"
        Columnas(10) = "nCapPag"
        Columnas(11) = "nIntPag"
        Columnas(12) = "nIntMorPag"
        Columnas(13) = "nIgvPag"
        Columnas(14) = "nEstado"
        Columnas(15) = "nNroCalen"
        Columnas(16) = "nTipo"
        Columnas(17) = "nIntGra"
        Columnas(18) = "nIntGraPag"

        '1->  Fecha de Vencimiento
        '2 -> Nro Cuota
        '5 -> Interes            
        '7 -> Gastos
        '6 -> IGV
        '4 -> Capital
        '3 -> Cuota Total
        '8-> Saldo

        D = New ComTipos.clsDataset
        D.CrearVacio(Columnas)
        For i = 1 To UBound(oCal)
            nCursor = D.AddRegistro()
            D.Seek(nCursor)
            D.AsignarDato(Trim(Str(pnCodCred)), "nCodCred")
            D.AsignarDato(Trim(Str(pnCodAge)), "nCodAge")
            D.AsignarDato(oCal(i, 2), "nNroCuota")
            D.AsignarDato(Format(CDate(oCal(i, 1)), "yyyy/MM/dd"), "dFecVenc")
            D.AsignarDato("", "dFecPago")
            D.AsignarDato(oCal(i, 4), "nCapital")
            D.AsignarDato(oCal(i, 5), "nIntComp")
            D.AsignarDato("0", "nIntMor")
            D.AsignarDato(oCal(i, 6), "nIgv")
            D.AsignarDato("0", "nCapPag")
            D.AsignarDato("0", "nIntPag")
            D.AsignarDato("0", "nIntMorPag")
            D.AsignarDato("0", "nIgvPag")
            D.AsignarDato("0", "nEstado")
            D.AsignarDato(Str(pnNroCalen), "nNroCalen")
            D.AsignarDato("2", "nTipo")
            D.AsignarDato(oCal(i, 9), "nIntGra")
            D.AsignarDato("0", "nIntGraPag")
        Next

        GeneraCalendarioDataSet = D

        D = Nothing
        oCal = Nothing

    End Function



    Public Function HayMontosNegativos(ByVal pMatCal As Object) As Boolean
        Dim i As Integer

        HayMontosNegativos = False

        For i = 1 To UBound(pMatCal)
            If CDbl(pMatCal(i, 4)) < 0 Or CDbl(pMatCal(i, 5)) < 0 Then
                HayMontosNegativos = True
                Exit For
            End If
        Next
    End Function

    Public Function GeneraCalendario(ByVal pnMontoGasto As String, ByVal poFeriado As ComTipos.clsDataset, ByVal pnPrestamo As Double, _
        ByVal pnCuotas As Integer, ByVal pnTasa As Double, ByVal pnDia As Integer, _
        ByVal pdFecDes As Date, ByVal pnCuotasGracia As Integer, _
        ByVal pnTipoGracia As Integer, ByVal pnGasto As Double, Optional ByVal pnProd As Integer = 0, Optional ByVal pnSubProd As Integer = 0, _
        Optional ByVal pbIncluyeGasto As Boolean = False, Optional ByVal pnTipoPeriodo As Integer = 1, Optional ByVal pnPlazo As Integer = 30, _
        Optional ByVal pnValorITF As Double = 0, Optional ByVal pbSoloInteres As Boolean = False, _
        Optional ByVal pnTasaDesgravamen As Double = 0, Optional ByVal pnTipoDesembParcial As Boolean = False, Optional ByVal pMatrizDesemb As ComTipos.clsDataset = Nothing, _
        Optional ByRef MatCalenCuotasDesemParDetalle As ComTipos.clsDataset = Nothing, Optional ByRef pdFechaPrimerPago As DateTime = Nothing, _
        Optional ByVal pnMontoCuotaObj As Double = -1, _
        Optional ByVal bCuotaDoble As Boolean = False, Optional ByVal bCDPrimeraGratif As Boolean = True, _
    Optional ByVal bRedondeo As Boolean = False _
) As Object 'Napoleon Silva - Campañas Enlatados 12/09/2011

        Dim i, j As Integer
        'Dim MatDatos(pnCuotas, 9) As String
        Dim MatDatos(pnCuotas, 10) As String 'ARCV 21-05-2009
        Dim MatColumnasDesemPar(6) As String 'ARCV 21-05-2009
        Dim nMonto As Double
        Dim nInteres As Double
        Dim dFecha As Date
        Dim MatFechas As Object
        Dim nMontoCuota As Double

        Dim nTasaGasto As Double
        Dim nIGV As Double
        Dim nIntGracia As Double
        Dim nDiasGracia As Integer
        Dim pnFinCuotaGracia As Integer
        Dim nTasaDesgravamenDeducido As Double
        Dim nMontoDesgravamenDeducido As Double
        Dim bOK As Boolean
        Dim nIteraciones As Integer
        Dim nIntPend As Double
        Dim nIntComTempo As Double
        Dim nFila As Integer

        'vvillafranca 08-11-2012
        'Dim bHabEmparejador as Boolean

        '
        Dim nMontoCuotaBK As Double  'ARCV 08-11-2012

        Dim dFecPago1 As Date  'vvillafranca 14-11-2012
        Dim bHabBusq1Pago As Boolean = True  'vvillafranca 14-11-2012

        'cpua 21/04/2010
        If pnSubProd = 4 Then
            pdFechaPrimerPago = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
            pdFechaPrimerPago = DateAdd(DateInterval.Day, pnPlazo, pdFechaPrimerPago)
        Else
            If pdFechaPrimerPago = CDate("01/01/0001") Then
                If pnTipoPeriodo = 2 Then
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, pnPlazo, pdFechaPrimerPago)

                Else
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, 20, pdFecDes)
                    Do While Day(pdFechaPrimerPago) <> pnDia
                        pdFechaPrimerPago = DateAdd(DateInterval.Day, 1, pdFechaPrimerPago)
                    Loop
                    pnCuotasGracia = 0
                End If

            End If
        End If

        nTasaDesgravamenDeducido = 0
        nMontoDesgravamenDeducido = 0

        oFeriados = poFeriado

        nTasaGasto = Format(pnGasto / pnPrestamo * 100, "#0.00")

        If pnValorITF = 0 Then
            'nIGV = 19
            'Se cambio el valor del igv a 18%...22/02/2011 ....cpua
            nIGV = 18
        Else
            nIGV = 0
        End If

        nGasto = 0
        If pbIncluyeGasto Then
            nGasto = Format(pnMontoGasto / pnCuotas, "#0.00")
            'nGasto = Format(Format(ObtieneGastoTotalCal(pnPrestamo, 1, pnProd, pnSubProd), "#0.00") / pnCuotas, "#0.00")
        End If

        nMonto = pnPrestamo
        If pnCuotasGracia > 0 And pnTipoGracia <> 4 Then
            dFecha = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
            nDiasGracia = DateDiff(DateInterval.Day, pdFecDes, dFecha)
        Else
            dFecha = pdFecDes
        End If


        nTasaDesgravamenDeducido = pnTasaDesgravamen
        nMontoDesgravamenDeducido = 0
        '************************************************
        'SI ES FECHA FIJA
        '************************************************
        If pnTipoPeriodo = 1 Then
            MatFechas = CalendarioFechas(dFecha, pnDia, pnCuotas, pnTipoGracia, pnCuotasGracia, pnFinCuotaGracia, pdFechaPrimerPago)
            'Anulo la gracia y regreso a la fecha de desembolso real para 
            'que todo se cobre en la primera cuota como IC
            nDiasGracia = 0
            dFecha = pdFecDes
            pnCuotasGracia = 0
        Else
            '************************************************
            'SI ES PERIODO FIJO
            '************************************************
            MatFechas = CalendarioFechasCuotaFija(dFecha, pnPlazo, pnCuotas, pnTipoGracia, pnCuotasGracia, pnFinCuotaGracia, pdFechaPrimerPago)
            If pnPlazo < 30 Then
                nMontoDesgravamenDeducido = Format((((nTasaDesgravamenDeducido / 100) * ((pnPlazo * pnCuotas) / 20)) * pnPrestamo) / pnCuotas, "#,0.00")
            End If
        End If


        pnFinCuotaGracia = IIf(pnFinCuotaGracia < 0, 0, pnFinCuotaGracia)
        'nMontoCUota = CuotaFechaFija(pdFecDes, pnTasa, nIGV, pnDiasGracia, pnCuotas, MatFechas, pnPrestamo)
        'nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia) 
        'nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia, pnTasaDesgravamen * 100) 'ARCV 21-05-2009



        nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia, nTasaDesgravamenDeducido) 'ARCV 21-05-2009



        '----------

        If nDiasGracia > 0 Then
            nIntGracia = ((1 + pnTasa / 100) ^ (nDiasGracia / 30)) - 1 'Interes
            If pnTipoGracia <> 3 And pnTipoGracia <> 4 Then
                If pnTipoGracia = 1 Then    'Proratear
                    nIntGracia = Format((nIntGracia * nMonto) / pnCuotas, "#0.00")
                Else    'Primera Cuota
                    nIntGracia = Format((nIntGracia * nMonto), "#0.00")
                End If
            Else
                nIntGracia = 0
            End If
        End If

        ' Se es desembolso parcial generar calendario de Desembolsos Parciales
        'M(1) = "dFecha"
        'M(2) = "nPrestamo"
        'M(3) = "nItem"
        Dim nInteresDesPar As Double
        nInteresDesPar = 0
        nFila = 0
        If pnTipoDesembParcial Then


            MatDatos(1, 1) = MatFechas(1) 'Fecha de Vencimiento
            MatDatos(1, 2) = Trim(Str(1)) 'Nro Cuota

            pMatrizDesemb.Inicio()

            'Inicializa Matriz de Desembolsos Detallado
            MatCalenCuotasDesemParDetalle = Nothing
            MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset
            MatColumnasDesemPar(1) = "nItem"
            MatColumnasDesemPar(2) = "dFecha"
            MatColumnasDesemPar(3) = "nPrestamo"
            MatColumnasDesemPar(4) = "nEstado"
            MatColumnasDesemPar(5) = "dFechaVenc"
            MatColumnasDesemPar(6) = "nInteres"

            MatCalenCuotasDesemParDetalle.CrearVacio(MatColumnasDesemPar)

            pMatrizDesemb.Inicio()
            For i = 1 To pMatrizDesemb.nNumFilas
                nFila = MatCalenCuotasDesemParDetalle.AddRegistro()

                nMonto = CDbl(pMatrizDesemb.ObtenerDato("nPrestamo"))
                dFecha = CDate(pMatrizDesemb.ObtenerDato("dFecha"))

                nInteres = ((1 + pnTasa / 100) ^ (DateDiff(DateInterval.Day, dFecha, CDate(MatFechas(1))) / 30)) - 1 'Interes
                nInteres = nInteres * nMonto 'Interes

                nInteresDesPar = nInteresDesPar + nInteres
                MatCalenCuotasDesemParDetalle.AsignarDato(i, "nItem", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(nMonto, "nPrestamo", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(0, "nEstado", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(Format(CDate(MatFechas(1)), "dd/MM/yyyy"), "dFechaVenc", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(nInteres, "nInteres", nFila)
                pMatrizDesemb.Siguiente()

            Next i


            MatDatos(1, 5) = Format(nInteresDesPar, "#0.00") 'Interes  
            'ARCV 21-08-2009
            'MatDatos(1, 7) = "0.00"
            MatDatos(1, 7) = Format(nGasto, "#0.00") 'Gastos    
            '-------------------------------
            MatDatos(1, 9) = "0.00"

            If pnValorITF = 0 Then
                MatDatos(1, 6) = Format((CDbl(MatDatos(1, 5))) * (nIGV / 100), "#0.00") 'IGV
            Else
                MatDatos(1, 6) = 0
            End If

            MatDatos(1, 4) = Format(pnPrestamo, "#0.00")
            MatDatos(1, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(pnPrestamo * nTasaDesgravamenDeducido / 100, "#0.00"))
            MatDatos(1, 3) = Format(CDbl(MatDatos(1, 5)) + CDbl(MatDatos(1, 4)) + CDbl(MatDatos(1, 6)) + CDbl(MatDatos(1, 7)) + CDbl(MatDatos(1, 9)) + CDbl(MatDatos(1, 10)), "#0.00")
            MatDatos(1, 8) = "0.00"

        Else
            nIteraciones = 0
            bOK = False
            nIntPend = 0
            nIntComTempo = 0

            Do While Not bOK

                If bRedondeo Then
                    nMontoCuota = Math.Round(nMontoCuota, 0)
                End If
                nMonto = pnPrestamo
                ReDim MatDatos(pnCuotas, 10)
                If pnCuotasGracia > 0 And pnTipoGracia <> 4 Then
                    dFecha = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
                    nDiasGracia = DateDiff(DateInterval.Day, pdFecDes, dFecha)
                Else
                    dFecha = pdFecDes
                End If

                nMontoCuotaBK = nMontoCuota

                For i = 1 To pnCuotas




                    'ARCV 08-11-2012 *************************
                    If bCuotaDoble Then
                        '======================vvillafranca 14-11-2012=============
                        If Not bCDPrimeraGratif And bHabBusq1Pago = True Then
                            For l As Integer = 1 To UBound(MatFechas)
                                If (DatePart(DateInterval.Month, CDate(MatFechas(l))) = 7 Or _
                                    DatePart(DateInterval.Month, CDate(MatFechas(l))) = 12) And bHabBusq1Pago Then
                                    dFecPago1 = CDate(MatFechas(l))
                                    bHabBusq1Pago = False
                                End If
                            Next
                        End If
                        '=========================================================
                        If (DatePart(DateInterval.Month, CDate(MatFechas(i))) = 7 Or _
                            DatePart(DateInterval.Month, CDate(MatFechas(i))) = 12) And CDate(MatFechas(i)) <> dFecPago1 Then

                            nMontoCuota = nMontoCuota * 2
                        Else
                            nMontoCuota = nMontoCuotaBK
                        End If
                    Else
                        nMontoCuota = nMontoCuotaBK
                    End If
                    '**************************************

                    MatDatos(i, 1) = MatFechas(i) 'Fecha de Vencimiento
                    MatDatos(i, 2) = Trim(Str(i)) 'Nro Cuota
                    nInteres = ((1 + pnTasa / 100) ^ (DateDiff(DateInterval.Day, dFecha, CDate(MatFechas(i))) / 30)) - 1 'Interes
                    nInteres = nInteres * nMonto 'Interes
                    MatDatos(i, 5) = Format(nInteres, "#0.00") 'Interes  

                    If pnTipoGracia <> 4 Or i > pnFinCuotaGracia Then
                        MatDatos(i, 7) = Format(nGasto, "#0.00") 'Gastos            
                    Else
                        MatDatos(i, 7) = "0.00" 'Gastos            
                    End If

                    'ARCV 07-12-2007
                    If pnTipoGracia = 1 Then 'Proratear
                        MatDatos(i, 9) = Format(nIntGracia, "#0.00") 'Interes de Gracia
                    Else    'Primera Cuota
                        If i = 1 Then
                            MatDatos(i, 9) = Format(nIntGracia, "#0.00")
                        Else
                            MatDatos(i, 9) = "0.00"
                        End If
                    End If

                    If Not (pnTipoGracia <> 4 Or i > pnFinCuotaGracia) Then
                        MatDatos(i, 9) = MatDatos(i, 5)
                        MatDatos(i, 5) = "0.00"
                    End If
                    '--------
                    If pnValorITF = 0 Then
                        MatDatos(i, 6) = Format((CDbl(MatDatos(i, 5))) * (nIGV / 100), "#0.00") 'IGV
                    Else
                        MatDatos(i, 6) = 0
                    End If

                    If pnTipoGracia <> 4 Or i > pnFinCuotaGracia Then
                        If i <> pnCuotas Then

                            'MatDatos(i, 4) = Format(nMontoCUota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7))), "#0.00") 'Capital

                            'ARCV 21-05-2009
                            'MatDatos(i, 4) = Format(nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6))), "#0.00") 'Capital
                            'MatDatos(i, 4) = Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)))) / (1 + pnTasaDesgravamen / 100), "#0.00")
                            'MatDatos(i, 10) = IIf(pnTasaDesgravamen = 0, "0.00", Format(nMonto * pnTasaDesgravamen / 100, "#0.00"))
                            If nMontoDesgravamenDeducido > 0 Then
                                MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                            Else
                                MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(nMonto * nTasaDesgravamenDeducido / 100, "#0.00"))
                            End If

                            If nIntPend > 0 Then
                                MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) + nIntPend, "#,0.00")
                                MatDatos(i, 6) = Format(CDbl(MatDatos(i, 5)) * (nIGV / 100), "#,0.00")
                            End If
                            ' SI Calculo de interes rebasa al monto de la cuota
                            If CDbl(Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 10)))), "#0.00")) < 0 Then
                                MatDatos(i, 4) = "0.00"
                                nIntComTempo = CDbl(MatDatos(i, 5))
                                If pnValorITF = 0 Then
                                    MatDatos(i, 5) = Format(((nMontoCuota - CDbl(MatDatos(i, 10))) / (1 + nIGV / 100)), "#,0.00")
                                    MatDatos(i, 6) = Format(((nMontoCuota - CDbl(MatDatos(i, 10))) - CDbl(MatDatos(i, 5))), "#,0.00")
                                    nIntPend = nIntPend + nIntComTempo - CDbl(MatDatos(i, 5))
                                Else
                                    MatDatos(i, 5) = Format((nMontoCuota - CDbl(MatDatos(i, 10))), "#,0.00")
                                    nIntPend = nIntComTempo - CDbl(MatDatos(i, 5))
                                End If

                            Else
                                MatDatos(i, 4) = Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 10)))), "#0.00")
                                nIntPend = 0
                            End If

                            '---------------
                            '5: INTERES , 6: IGV
                        Else
                            If pnCuotas = 1 Then
                                MatDatos(i, 4) = Format(nMonto, "#0.00")
                                If nMontoDesgravamenDeducido > 0 Then
                                    MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                                Else
                                    MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(nMonto * nTasaDesgravamenDeducido / 100, "#0.00")) 'ARCV
                                End If

                            Else
                                MatDatos(i, 4) = MatDatos(i - 1, 8)
                                If nMontoDesgravamenDeducido > 0 Then
                                    MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                                Else
                                    MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(CDbl(MatDatos(i - 1, 8)) * nTasaDesgravamenDeducido / 100, "#0.00")) 'ARCV
                                End If
                            End If
                        End If
                    Else
                        'Para cuotas de gracia capital es cero
                        MatDatos(i, 4) = "0.00"
                    End If

                    If pbSoloInteres Then
                        If i <> pnCuotas Then
                            MatDatos(i, 4) = "0.00"
                        Else
                            MatDatos(i, 4) = Format(nMonto, "#0.00")
                        End If
                    End If

                    If pnValorITF = 0 Then
                        'Agrego el IGV del Gasto y de la Gracia si lo hubiera
                        MatDatos(i, 6) = Format(CDbl(MatDatos(i, 6)) + ((nIGV / 100) * CDbl(MatDatos(i, 9))), "#0.00")
                        MatDatos(i, 6) = Format(CDbl(MatDatos(i, 6)) + ((nIGV / 100) * nGasto), "#0.00")
                    Else
                        'No Se activa para que el ITF no se guarde en el Calendario en Base de Datos
                        'MatDatos(i, 6) = Format(CDbl(Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)), "#0.00")) * pnValorITF, "#,0.00")
                        MatDatos(i, 6) = "0.00"
                    End If


                    'MatDatos(i, 3) = Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)) , "#0.00") 'Cuota Total
                    MatDatos(i, 3) = Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)) + CDbl(MatDatos(i, 10)), "#0.00") 'Cuota Total 'ARCV 21-05-2009

                    'nMonto = IIf(i = pnCuotas, nMonto, Format(nMonto - CDbl(MatDatos(i, 4)), "#0.00")) 'ARCV 22-05-2009
                    nMonto = Format(nMonto - CDbl(MatDatos(i, 4)), "#0.00")
                    MatDatos(i, 8) = Format(nMonto, "#0.00") 'Saldo
                    dFecha = CDate(MatDatos(i, 1))

                Next i

                nMontoCuota = nMontoCuotaBK 'ARCV 08-11-2012 *************************************************

                If Not pbSoloInteres And (pnTipoGracia <> 4 Or nDiasGracia = 0) Then
                    If bRedondeo Then


                        If UBound(MatDatos) > 1 Then
                            If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 5 Then
                                nMontoCuota = CDbl(Format(nMontoCuota + 1, "#,0.00"))
                                nIteraciones = nIteraciones + 1
                                If nIteraciones > 1500 Then
                                    bOK = True
                                End If
                            ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) = 0 Then
                                bOK = True
                            Else
                                If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) <= -5 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota - 1, "#,0.00"))
                                    nIteraciones = nIteraciones + 1
                                Else
                                    nIteraciones = nIteraciones + 1
                                End If
                                If nIteraciones > 1500 Then
                                    bOK = True
                                End If
                            End If

                        Else
                            bOK = True
                        End If



                    Else

                        If bCuotaDoble Then
                            If UBound(MatDatos) > 1 Then
                                Dim X As Integer
                                X = UBound(MatDatos) - 1
                                If Month(CDate(MatDatos(X, 1))) = 12 Or Month(CDate(MatDatos(X, 1))) = 7 Then
                                    X = X - 1
                                End If

                                If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(X, 3))) > 10 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.5, "#,0.00"))

                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(X, 3))) > 1 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.1, "#,0.00"))

                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(X, 3))) > 0 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.01, "#,0.00"))

                                    'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -10 Then
                                    '    nMontoCuota = CDbl(Format(nMontoCuota - 0.5, "#,0.00"))
                                    '    nIteraciones = nIteraciones + 1
                                    'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -1 Then
                                    '    nMontoCuota = CDbl(Format(nMontoCuota - 0.1, "#,0.00"))
                                    '    nIteraciones = nIteraciones + 1
                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(X, 3))) = 0 Then
                                    bOK = True
                                Else
                                    If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(X, 3))) <= -10 Then
                                        nMontoCuota = CDbl(Format(nMontoCuota - 0.5, "#,0.00"))
                                        nIteraciones = nIteraciones + 1
                                    Else
                                        nMontoCuota = CDbl(Format(nMontoCuota - 0.01, "#,0.00"))
                                        nIteraciones = nIteraciones + 1
                                    End If
                                    If nIteraciones > 1500 Then
                                        bOK = True
                                    End If
                                End If

                            Else
                                bOK = True
                            End If

                        Else
                            If UBound(MatDatos) > 1 Then
                                If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 10 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.5, "#,0.00"))

                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 1 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.1, "#,0.00"))

                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 0 Then
                                    nMontoCuota = CDbl(Format(nMontoCuota + 0.01, "#,0.00"))

                                    'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -10 Then
                                    '    nMontoCuota = CDbl(Format(nMontoCuota - 0.5, "#,0.00"))
                                    '    nIteraciones = nIteraciones + 1
                                    'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -1 Then
                                    '    nMontoCuota = CDbl(Format(nMontoCuota - 0.1, "#,0.00"))
                                    '    nIteraciones = nIteraciones + 1
                                ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) = 0 Then
                                    bOK = True
                                Else
                                    If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) <= -10 Then
                                        nMontoCuota = CDbl(Format(nMontoCuota - 0.5, "#,0.00"))
                                        nIteraciones = nIteraciones + 1
                                    Else
                                        nMontoCuota = CDbl(Format(nMontoCuota - 0.01, "#,0.00"))
                                        nIteraciones = nIteraciones + 1
                                    End If
                                    If nIteraciones > 1500 Then
                                        bOK = True
                                    End If
                                End If

                            Else
                                bOK = True
                            End If
                        End If
                        


                    End If


                Else

                    bOK = True

                End If

            Loop

        End If


        'Napoleon Silva
        'Campañas Enlatados 12/09/2011
        'Actualiza Matriz si es Campaña de Enlatados
        Dim nDiferenciaCuota As Double
        Dim nDiferenciaTotal As Double = 0 'ARCV 18-10-2011
        If pnMontoCuotaObj > -1 Then
            For i = 1 To UBound(MatDatos)
                nDiferenciaCuota = Format(pnMontoCuotaObj - CDbl(MatDatos(i, 3)), "#,0.00")
                nDiferenciaTotal += nDiferenciaCuota 'ARCV 18-10-2011
                'Hay que bajar la cuota reduciendo los intereses por ser enlatado
                If nDiferenciaCuota < 0 Then
                    If CDbl(MatDatos(i, 5)) > (-1 * nDiferenciaCuota) Then
                        MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) - (-1 * nDiferenciaCuota), "#,0.00")
                        MatDatos(i, 3) = Format(CDbl(MatDatos(i, 3)) - (-1 * nDiferenciaCuota), "#,0.00")
                    End If
                Else
                    MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) + nDiferenciaCuota, "#,0.00")
                    MatDatos(i, 3) = Format(CDbl(MatDatos(i, 3)) + nDiferenciaCuota, "#,0.00")
                End If
            Next

            'ARCV 18-10-2011
            If pnTipoPeriodo = 1 Then 'SOLO para Fecha Fija
                MatDatos(UBound(MatDatos), 5) = Format(CDbl(MatDatos(UBound(MatDatos), 5)) - nDiferenciaTotal, "#,0.00")
                MatDatos(UBound(MatDatos), 3) = Format(CDbl(MatDatos(UBound(MatDatos), 3)) - nDiferenciaTotal, "#,0.00")
            End If
            '---------
        End If
        '============================== vvillafranca 05-11-2012 (Opcion Meses Pago doble Cuota) ========================
        'If bCuotaDoble Then
        '    ' Convierte el Dataset en Array 
        '    Dim arrayMeses(oMesesDoble.nNumFilas - 1) As Integer '= {12}
        '    Dim nPosicion As Integer = 0
        '    With oMesesDoble
        '        .Inicio()
        '        While Not .EOF
        '            arrayMeses(nPosicion) = CInt(oMesesDoble.ObtenerDato("nValor"))
        '            nPosicion += 1
        '            .Siguiente()
        '        End While
        '    End With

        '    'Ordena el Array en funcion a la Fecha de Primer Pago
        '    Dim p As Integer = 0
        '    Dim nAux As Integer = 0
        '    While arrayMeses(0) < Month(pdFechaPrimerPago) And UBound(arrayMeses) > 0
        '        nAux = arrayMeses(0)
        '        For q As Integer = 0 To UBound(arrayMeses)
        '            If q < UBound(arrayMeses) Then
        '                arrayMeses(q) = arrayMeses(q + 1)
        '            End If
        '        Next
        '        arrayMeses(UBound(arrayMeses)) = nAux
        '        p += 1
        '    End While

        '    'Crea el Array de Meses a Eliminar en el periodo de las cuotas
        '    'Dim bAvanza As Boolean = False
        '    'Dim bEliminaUltCuota As Boolean
        '    'Dim arrayMesesTotal(UBound(arrayMeses)) As Integer
        '    'Dim nMaxMatriz As Integer
        '    'arrayMesesTotal = arrayMeses
        '    'For k As Integer = 1 To UBound(MatDatos)
        '    '    For l As Integer = 0 To UBound(arrayMeses)
        '    '        If Month(MatDatos(k, 1)) = arrayMeses(l) Then
        '    '            nMaxMatriz = UBound(arrayMesesTotal)
        '    '            ReDim Preserve arrayMesesTotal(nMaxMatriz + 1)
        '    '            arrayMesesTotal(nMaxMatriz + 1) = arrayMeses(l)
        '    '        End If
        '    '    Next
        '    'Next


        '    Dim bAvanza As Boolean = False
        '    Dim bEliminaUltCuota As Boolean
        '    Dim arrayMesesTotal(UBound(arrayMeses)) As Integer
        '    Dim nMaxMatriz As Integer = 0
        '    'arrayMesesTotal = arrayMeses
        '    For k As Integer = 1 To UBound(MatDatos)
        '        For l As Integer = 0 To UBound(arrayMeses)
        '            If Month(MatDatos(k, 1)) = arrayMeses(l) Then
        '                ReDim Preserve arrayMesesTotal(nMaxMatriz)
        '                arrayMesesTotal(nMaxMatriz) = arrayMeses(l)
        '                nMaxMatriz += 1
        '            End If
        '        Next
        '    Next

        '    'Elimina la cuota en funcion al Array Total
        '    Dim MatDatosAux(UBound(MatDatos) - 1, 11)
        '    ReDim Preserve MatDatos(UBound(MatDatos), 11)
        '    For m As Integer = 0 To UBound(arrayMesesTotal)
        '        For x As Integer = 1 To UBound(MatDatos)

        '            If Month(MatDatos(x, 1)) = arrayMesesTotal(m) And MatDatos(x, 11) <> 1 And bAvanza = False And UBound(MatDatos) <> x Then
        '                MatDatos(x, 3) = Format(CDbl(MatDatos(x, 3)) + CDbl(MatDatos(x + 1, 3)), "#,0.00")
        '                MatDatos(x, 4) = Format(CDbl(MatDatos(x, 4)) + CDbl(MatDatos(x + 1, 4)), "#,0.00")
        '                MatDatos(x, 5) = Format(CDbl(MatDatos(x, 5)) + CDbl(MatDatos(x + 1, 5)), "#,0.00")
        '                MatDatos(x, 6) = Format(CDbl(MatDatos(x, 6)) + CDbl(MatDatos(x + 1, 6)), "#,0.00")
        '                MatDatos(x, 7) = Format(CDbl(MatDatos(x, 7)) + CDbl(MatDatos(x + 1, 7)), "#,0.00")
        '                MatDatos(x, 8) = Format(CDbl(MatDatos(x + 1, 8)), "#,0.00")
        '                MatDatos(x, 9) = Format(CDbl(MatDatos(x, 9)) + CDbl(MatDatos(x + 1, 9)), "#,0.00")
        '                MatDatos(x, 10) = Format(CDbl(MatDatos(x, 10)) + CDbl(MatDatos(x + 1, 10)), "#,0.00")
        '                MatDatos(x, 11) = 1
        '                bAvanza = True
        '                bEliminaUltCuota = True
        '            End If
        '            If bAvanza And x + 1 < UBound(MatDatos) Then
        '                MatDatos(x + 1, 3) = Format(CDbl(MatDatos(x + 2, 3)), "#,0.00")
        '                MatDatos(x + 1, 4) = Format(CDbl(MatDatos(x + 2, 4)), "#,0.00")
        '                MatDatos(x + 1, 5) = Format(CDbl(MatDatos(x + 2, 5)), "#,0.00")
        '                MatDatos(x + 1, 6) = Format(CDbl(MatDatos(x + 2, 6)), "#,0.00")
        '                MatDatos(x + 1, 7) = Format(CDbl(MatDatos(x + 2, 7)), "#,0.00")
        '                MatDatos(x + 1, 8) = Format(CDbl(MatDatos(x + 2, 8)), "#,0.00")
        '                MatDatos(x + 1, 9) = Format(CDbl(MatDatos(x + 2, 9)), "#,0.00")
        '                MatDatos(x + 1, 10) = Format(CDbl(MatDatos(x + 2, 10)), "#,0.00")
        '            End If
        '            If bEliminaUltCuota And x = UBound(MatDatos) Then
        '                MatDatosAux = MatDatos
        '                ReDim MatDatos(x - 1, 11)
        '                For h As Integer = 1 To UBound(MatDatos)
        '                    For f As Integer = 1 To 11
        '                        MatDatos(h, f) = MatDatosAux(h, f)
        '                    Next
        '                Next
        '            End If

        '        Next
        '        bAvanza = False
        '        bEliminaUltCuota = False
        '    Next
        '    'retorna el Matdatos a 10 columnas
        '    Dim MatDatosAux2(UBound(MatDatos), 10)
        '    For h As Integer = 1 To UBound(MatDatos)
        '        For f As Integer = 1 To 10
        '            MatDatosAux2(h, f) = MatDatos(h, f)
        '        Next
        '    Next
        '    ReDim MatDatos(UBound(MatDatosAux2), 10)
        '    For h As Integer = 1 To UBound(MatDatos)
        '        For f As Integer = 1 To 10
        '            MatDatos(h, f) = MatDatosAux2(h, f)
        '        Next
        '    Next

        'End If

        '==============================================================================
        GeneraCalendario = MatDatos

    End Function

    Public Function GeneraCalendarioReprogramacion(ByVal pnMontoGasto As String, ByVal poFeriado As ComTipos.clsDataset, ByVal pnPrestamo As Double, _
            ByVal pnCuotas As Integer, ByVal pnTasa As Double, ByVal pnDia As Integer, _
            ByVal pdFecDes As Date, ByVal pnCuotasGracia As Integer, _
            ByVal pnTipoGracia As Integer, ByVal pnGasto As Double, Optional ByVal pnProd As Integer = 0, Optional ByVal pnSubProd As Integer = 0, _
            Optional ByVal pbIncluyeGasto As Boolean = False, Optional ByVal pnTipoPeriodo As Integer = 1, Optional ByVal pnPlazo As Integer = 30, _
            Optional ByVal pnValorITF As Double = 0, Optional ByVal pbSoloInteres As Boolean = False, _
            Optional ByVal pnTasaDesgravamen As Double = 0, Optional ByVal pnTipoDesembParcial As Boolean = False, Optional ByVal pMatrizDesemb As ComTipos.clsDataset = Nothing, _
            Optional ByRef MatCalenCuotasDesemParDetalle As ComTipos.clsDataset = Nothing, Optional ByRef pdFechaPrimerPago As DateTime = Nothing, _
            Optional ByVal pnMontoCuotaObj As Double = -1, Optional ByVal pnNroProxCuota As Integer = 1, Optional ByVal bCambiaFechaVenc As Boolean = False, Optional ByVal pdFechaVencimiento As DateTime = Nothing) As Object 'Napoleon Silva - Campañas Enlatados 12/09/2011

        Dim i, j As Integer
        'Dim MatDatos(pnCuotas, 9) As String
        Dim MatDatos(pnCuotas, 10) As String 'ARCV 21-05-2009
        'Dim MatDatos((pnCuotas - pnNroProxCuota) + 1, 10) As String 'cpua 22/05/2012..se modifico para generar solo las cuotas para la reprogramacion
        Dim MatColumnasDesemPar(6) As String 'ARCV 21-05-2009
        Dim nMonto As Double
        Dim nInteres As Double
        Dim dFecha As Date
        Dim MatFechas As Object
        Dim nMontoCuota As Double

        Dim nTasaGasto As Double
        Dim nIGV As Double
        Dim nIntGracia As Double
        Dim nDiasGracia As Integer
        Dim pnFinCuotaGracia As Integer
        Dim nTasaDesgravamenDeducido As Double
        Dim nMontoDesgravamenDeducido As Double
        Dim bOK As Boolean
        Dim nIteraciones As Integer
        Dim nIntPend As Double
        Dim nIntComTempo As Double
        Dim nFila As Integer
        Dim nNroProxCuota As Integer

        'cpua 21/04/2010
        If pnSubProd = 4 Then
            pdFechaPrimerPago = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
            pdFechaPrimerPago = DateAdd(DateInterval.Day, pnPlazo, pdFechaPrimerPago)
        Else
            If pdFechaPrimerPago = CDate("01/01/0001") Then
                If pnTipoPeriodo = 2 Then
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, pnPlazo, pdFechaPrimerPago)

                Else
                    pdFechaPrimerPago = DateAdd(DateInterval.Day, 20, pdFecDes)
                    Do While Day(pdFechaPrimerPago) <> pnDia
                        pdFechaPrimerPago = DateAdd(DateInterval.Day, 1, pdFechaPrimerPago)
                    Loop
                    pnCuotasGracia = 0
                End If

            End If
        End If

        nTasaDesgravamenDeducido = 0
        nMontoDesgravamenDeducido = 0

        oFeriados = poFeriado

        nTasaGasto = Format(pnGasto / pnPrestamo * 100, "#0.00")

        If pnValorITF = 0 Then
            'nIGV = 19
            'Se cambio el valor del igv a 18%...22/02/2011 ....cpua
            nIGV = 18
        Else
            nIGV = 0
        End If

        nGasto = 0
        If pbIncluyeGasto Then
            nGasto = Format(pnMontoGasto / pnCuotas, "#0.00")
            'nGasto = Format(Format(ObtieneGastoTotalCal(pnPrestamo, 1, pnProd, pnSubProd), "#0.00") / pnCuotas, "#0.00")
        End If

        nMonto = pnPrestamo
        If pnCuotasGracia > 0 And pnTipoGracia <> 4 Then
            dFecha = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
            nDiasGracia = DateDiff(DateInterval.Day, pdFecDes, dFecha)
        Else
            dFecha = pdFecDes
        End If


        nTasaDesgravamenDeducido = pnTasaDesgravamen
        nMontoDesgravamenDeducido = 0
        '************************************************
        'SI ES FECHA FIJA
        '************************************************
        If pnTipoPeriodo = 1 Then
            MatFechas = CalendarioFechas(dFecha, pnDia, pnCuotas, pnTipoGracia, pnCuotasGracia, pnFinCuotaGracia, pdFechaPrimerPago)
            'Anulo la gracia y regreso a la fecha de desembolso real para 
            'que todo se cobre en la primera cuota como IC
            nDiasGracia = 0
            dFecha = pdFecDes
            pnCuotasGracia = 0
        Else
            '************************************************
            'SI ES PERIODO FIJO
            '************************************************
            MatFechas = CalendarioFechasCuotaFija(dFecha, pnPlazo, pnCuotas, pnTipoGracia, pnCuotasGracia, pnFinCuotaGracia, pdFechaPrimerPago)
            If pnPlazo < 30 Then
                nMontoDesgravamenDeducido = Format((((nTasaDesgravamenDeducido / 100) * ((pnPlazo * pnCuotas) / 20)) * pnPrestamo) / pnCuotas, "#,0.00")
            End If
        End If


        pnFinCuotaGracia = IIf(pnFinCuotaGracia < 0, 0, pnFinCuotaGracia)
        'nMontoCUota = CuotaFechaFija(pdFecDes, pnTasa, nIGV, pnDiasGracia, pnCuotas, MatFechas, pnPrestamo)
        'nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia) 
        'nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia, pnTasaDesgravamen * 100) 'ARCV 21-05-2009


        nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo, pnFinCuotaGracia, nTasaDesgravamenDeducido) 'ARCV 21-05-2009
        '----------

        If nDiasGracia > 0 Then
            nIntGracia = ((1 + pnTasa / 100) ^ (nDiasGracia / 30)) - 1 'Interes
            If pnTipoGracia <> 3 And pnTipoGracia <> 4 Then
                If pnTipoGracia = 1 Then    'Proratear
                    nIntGracia = Format((nIntGracia * nMonto) / pnCuotas, "#0.00")
                Else    'Primera Cuota
                    nIntGracia = Format((nIntGracia * nMonto), "#0.00")
                End If
            Else
                nIntGracia = 0
            End If
        End If

        ' Se es desembolso parcial generar calendario de Desembolsos Parciales
        'M(1) = "dFecha"
        'M(2) = "nPrestamo"
        'M(3) = "nItem"
        Dim nInteresDesPar As Double
        nInteresDesPar = 0
        nFila = 0
        If pnTipoDesembParcial Then


            MatDatos(1, 1) = MatFechas(1) 'Fecha de Vencimiento
            MatDatos(1, 2) = Trim(Str(1)) 'Nro Cuota

            pMatrizDesemb.Inicio()

            'Inicializa Matriz de Desembolsos Detallado
            MatCalenCuotasDesemParDetalle = Nothing
            MatCalenCuotasDesemParDetalle = New ComTipos.clsDataset
            MatColumnasDesemPar(1) = "nItem"
            MatColumnasDesemPar(2) = "dFecha"
            MatColumnasDesemPar(3) = "nPrestamo"
            MatColumnasDesemPar(4) = "nEstado"
            MatColumnasDesemPar(5) = "dFechaVenc"
            MatColumnasDesemPar(6) = "nInteres"

            MatCalenCuotasDesemParDetalle.CrearVacio(MatColumnasDesemPar)

            pMatrizDesemb.Inicio()
            For i = 1 To pMatrizDesemb.nNumFilas
                nFila = MatCalenCuotasDesemParDetalle.AddRegistro()

                nMonto = CDbl(pMatrizDesemb.ObtenerDato("nPrestamo"))
                dFecha = CDate(pMatrizDesemb.ObtenerDato("dFecha"))

                nInteres = ((1 + pnTasa / 100) ^ (DateDiff(DateInterval.Day, dFecha, CDate(MatFechas(1))) / 30)) - 1 'Interes
                nInteres = nInteres * nMonto 'Interes

                nInteresDesPar = nInteresDesPar + nInteres
                MatCalenCuotasDesemParDetalle.AsignarDato(i, "nItem", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(Format(dFecha, "dd/MM/yyyy"), "dFecha", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(nMonto, "nPrestamo", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(0, "nEstado", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(Format(CDate(MatFechas(1)), "dd/MM/yyyy"), "dFechaVenc", nFila)
                MatCalenCuotasDesemParDetalle.AsignarDato(nInteres, "nInteres", nFila)
                pMatrizDesemb.Siguiente()

            Next i


            MatDatos(1, 5) = Format(nInteresDesPar, "#0.00") 'Interes  
            'ARCV 21-08-2009
            'MatDatos(1, 7) = "0.00"
            MatDatos(1, 7) = Format(nGasto, "#0.00") 'Gastos    
            '-------------------------------
            MatDatos(1, 9) = "0.00"

            If pnValorITF = 0 Then
                MatDatos(1, 6) = Format((CDbl(MatDatos(1, 5))) * (nIGV / 100), "#0.00") 'IGV
            Else
                MatDatos(1, 6) = 0
            End If

            MatDatos(1, 4) = Format(pnPrestamo, "#0.00")
            MatDatos(1, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(pnPrestamo * nTasaDesgravamenDeducido / 100, "#0.00"))
            MatDatos(1, 3) = Format(CDbl(MatDatos(1, 5)) + CDbl(MatDatos(1, 4)) + CDbl(MatDatos(1, 6)) + CDbl(MatDatos(1, 7)) + CDbl(MatDatos(1, 9)) + CDbl(MatDatos(1, 10)), "#0.00")
            MatDatos(1, 8) = "0.00"

        Else
            nIteraciones = 0
            bOK = False
            nIntPend = 0
            nIntComTempo = 0

            Do While Not bOK

                nMonto = pnPrestamo
                ReDim MatDatos(pnCuotas, 10)
                'ReDim MatDatos((pnCuotas - pnNroProxCuota) + 1, 10)
                If pnCuotasGracia > 0 And pnTipoGracia <> 4 Then
                    dFecha = DateAdd(DateInterval.Day, pnCuotasGracia, pdFecDes)
                    nDiasGracia = DateDiff(DateInterval.Day, pdFecDes, dFecha)
                Else
                    dFecha = pdFecDes
                End If
                nNroProxCuota = pnNroProxCuota
                For i = 1 To pnCuotas '...se realizo la modificacion para los calendarios de las reprogramaciones...cpua 22/05/2012
                    'For i = 1 To (pnCuotas - pnNroProxCuota) + 1
                    If i = pnCuotas Then
                        MatDatos(i, 1) = IIf(bCambiaFechaVenc, MatFechas(i), pdFechaVencimiento) 'Fecha de Vencimiento
                    Else
                        MatDatos(i, 1) = MatFechas(i)
                    End If
                    'MatDatos(i, 2) = Trim(Str(i)) 'Nro Cuota
                    MatDatos(i, 2) = Trim(Str(nNroProxCuota)) 'Nro Cuota
                    nInteres = ((1 + pnTasa / 100) ^ (DateDiff(DateInterval.Day, dFecha, CDate(MatFechas(i))) / 30)) - 1 'Interes
                    nInteres = nInteres * nMonto 'Interes
                    MatDatos(i, 5) = Format(nInteres, "#0.00") 'Interes  

                    If pnTipoGracia <> 4 Or i > pnFinCuotaGracia Then
                        MatDatos(i, 7) = Format(nGasto, "#0.00") 'Gastos            
                    Else
                        MatDatos(i, 7) = "0.00" 'Gastos            
                    End If

                    'ARCV 07-12-2007
                    If pnTipoGracia = 1 Then 'Proratear
                        MatDatos(i, 9) = Format(nIntGracia, "#0.00") 'Interes de Gracia
                    Else    'Primera Cuota
                        If i = 1 Then
                            MatDatos(i, 9) = Format(nIntGracia, "#0.00")
                        Else
                            MatDatos(i, 9) = "0.00"
                        End If
                    End If

                    If Not (pnTipoGracia <> 4 Or i > pnFinCuotaGracia) Then
                        MatDatos(i, 9) = MatDatos(i, 5)
                        MatDatos(i, 5) = "0.00"
                    End If
                    '--------
                    If pnValorITF = 0 Then
                        MatDatos(i, 6) = Format((CDbl(MatDatos(i, 5))) * (nIGV / 100), "#0.00") 'IGV
                    Else
                        MatDatos(i, 6) = 0
                    End If

                    If pnTipoGracia <> 4 Or i > pnFinCuotaGracia Then
                        If i <> pnCuotas Then

                            'MatDatos(i, 4) = Format(nMontoCUota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7))), "#0.00") 'Capital

                            'ARCV 21-05-2009
                            'MatDatos(i, 4) = Format(nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6))), "#0.00") 'Capital
                            'MatDatos(i, 4) = Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)))) / (1 + pnTasaDesgravamen / 100), "#0.00")
                            'MatDatos(i, 10) = IIf(pnTasaDesgravamen = 0, "0.00", Format(nMonto * pnTasaDesgravamen / 100, "#0.00"))
                            If nMontoDesgravamenDeducido > 0 Then
                                MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                            Else
                                MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(nMonto * nTasaDesgravamenDeducido / 100, "#0.00"))
                            End If

                            If nIntPend > 0 Then
                                MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) + nIntPend, "#,0.00")
                                MatDatos(i, 6) = Format(CDbl(MatDatos(i, 5)) * (nIGV / 100), "#,0.00")
                            End If
                            ' SI Calculo de interes rebasa al monto de la cuota
                            If CDbl(Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 10)))), "#0.00")) < 0 Then
                                MatDatos(i, 4) = "0.00"
                                nIntComTempo = CDbl(MatDatos(i, 5))
                                If pnValorITF = 0 Then
                                    MatDatos(i, 5) = Format(((nMontoCuota - CDbl(MatDatos(i, 10))) / (1 + nIGV / 100)), "#,0.00")
                                    MatDatos(i, 6) = Format(((nMontoCuota - CDbl(MatDatos(i, 10))) - CDbl(MatDatos(i, 5))), "#,0.00")
                                    nIntPend = nIntPend + nIntComTempo - CDbl(MatDatos(i, 5))
                                Else
                                    MatDatos(i, 5) = Format((nMontoCuota - CDbl(MatDatos(i, 10))), "#,0.00")
                                    nIntPend = nIntComTempo - CDbl(MatDatos(i, 5))
                                End If

                            Else
                                MatDatos(i, 4) = Format((nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 10)))), "#0.00")
                                nIntPend = 0
                            End If

                            '---------------
                            '5: INTERES , 6: IGV
                        Else
                            If pnCuotas = 1 Then
                                MatDatos(i, 4) = Format(nMonto, "#0.00")
                                If nMontoDesgravamenDeducido > 0 Then
                                    MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                                Else
                                    MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(nMonto * nTasaDesgravamenDeducido / 100, "#0.00")) 'ARCV
                                End If

                            Else
                                MatDatos(i, 4) = MatDatos(i - 1, 8)
                                If nMontoDesgravamenDeducido > 0 Then
                                    MatDatos(i, 10) = Format(nMontoDesgravamenDeducido, "#,0.00")
                                Else
                                    MatDatos(i, 10) = IIf(nTasaDesgravamenDeducido = 0, "0.00", Format(CDbl(MatDatos(i - 1, 8)) * nTasaDesgravamenDeducido / 100, "#0.00")) 'ARCV
                                End If
                            End If
                        End If
                    Else
                        'Para cuotas de gracia capital es cero
                        MatDatos(i, 4) = "0.00"
                    End If

                    If pbSoloInteres Then
                        If i <> pnCuotas Then
                            MatDatos(i, 4) = "0.00"
                        Else
                            MatDatos(i, 4) = Format(nMonto, "#0.00")
                        End If
                    End If

                    If pnValorITF = 0 Then
                        'Agrego el IGV del Gasto y de la Gracia si lo hubiera
                        MatDatos(i, 6) = Format(CDbl(MatDatos(i, 6)) + ((nIGV / 100) * CDbl(MatDatos(i, 9))), "#0.00")
                        MatDatos(i, 6) = Format(CDbl(MatDatos(i, 6)) + ((nIGV / 100) * nGasto), "#0.00")
                    Else
                        'No Se activa para que el ITF no se guarde en el Calendario en Base de Datos
                        'MatDatos(i, 6) = Format(CDbl(Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)), "#0.00")) * pnValorITF, "#,0.00")
                        MatDatos(i, 6) = "0.00"
                    End If


                    'MatDatos(i, 3) = Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)) , "#0.00") 'Cuota Total
                    MatDatos(i, 3) = Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)) + CDbl(MatDatos(i, 10)), "#0.00") 'Cuota Total 'ARCV 21-05-2009

                    'nMonto = IIf(i = pnCuotas, nMonto, Format(nMonto - CDbl(MatDatos(i, 4)), "#0.00")) 'ARCV 22-05-2009
                    nMonto = Format(nMonto - CDbl(MatDatos(i, 4)), "#0.00")
                    MatDatos(i, 8) = Format(nMonto, "#0.00") 'Saldo
                    dFecha = CDate(MatDatos(i, 1))
                    nNroProxCuota = nNroProxCuota + 1
                Next i

                If Not pbSoloInteres And (pnTipoGracia <> 4 Or nDiasGracia = 0) Then

                    If UBound(MatDatos) > 1 Then
                        If (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 10 Then
                            nMontoCuota = CDbl(Format(nMontoCuota + 0.5, "#,0.00"))
                        ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 1 Then
                            nMontoCuota = CDbl(Format(nMontoCuota + 0.1, "#,0.00"))
                        ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) > 0 Then
                            nMontoCuota = CDbl(Format(nMontoCuota + 0.01, "#,0.00"))

                            'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -10 Then
                            '    nMontoCuota = CDbl(Format(nMontoCuota - 0.5, "#,0.00"))
                            '    nIteraciones = nIteraciones + 1
                            'ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) < -1 Then
                            '    nMontoCuota = CDbl(Format(nMontoCuota - 0.1, "#,0.00"))
                            '    nIteraciones = nIteraciones + 1
                        ElseIf (CDbl(MatDatos(UBound(MatDatos), 3)) - CDbl(MatDatos(UBound(MatDatos) - 1, 3))) = 0 Then
                            bOK = True
                        Else
                            nMontoCuota = CDbl(Format(nMontoCuota - 0.01, "#,0.00"))
                            nIteraciones = nIteraciones + 1
                            If nIteraciones > 1500 Then
                                bOK = True
                            End If
                        End If

                    Else
                        bOK = True
                    End If
                Else

                    bOK = True

                End If

            Loop

        End If


        'Napoleon Silva
        'Campañas Enlatados 12/09/2011
        'Actualiza Matriz si es Campaña de Enlatados
        Dim nDiferenciaCuota As Double
        Dim nDiferenciaTotal As Double = 0 'ARCV 18-10-2011
        If pnMontoCuotaObj > -1 Then
            For i = 1 To UBound(MatDatos)
                nDiferenciaCuota = Format(pnMontoCuotaObj - CDbl(MatDatos(i, 3)), "#,0.00")
                nDiferenciaTotal += nDiferenciaCuota 'ARCV 18-10-2011
                'Hay que bajar la cuota reduciendo los intereses por ser enlatado
                If nDiferenciaCuota < 0 Then
                    If CDbl(MatDatos(i, 5)) > (-1 * nDiferenciaCuota) Then
                        MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) - (-1 * nDiferenciaCuota), "#,0.00")
                        MatDatos(i, 3) = Format(CDbl(MatDatos(i, 3)) - (-1 * nDiferenciaCuota), "#,0.00")
                    End If
                Else
                    MatDatos(i, 5) = Format(CDbl(MatDatos(i, 5)) + nDiferenciaCuota, "#,0.00")
                    MatDatos(i, 3) = Format(CDbl(MatDatos(i, 3)) + nDiferenciaCuota, "#,0.00")
                End If
            Next

            'ARCV 18-10-2011
            If pnTipoPeriodo = 1 Then 'SOLO para Fecha Fija
                MatDatos(UBound(MatDatos), 5) = Format(CDbl(MatDatos(UBound(MatDatos), 5)) - nDiferenciaTotal, "#,0.00")
                MatDatos(UBound(MatDatos), 3) = Format(CDbl(MatDatos(UBound(MatDatos), 3)) - nDiferenciaTotal, "#,0.00")
            End If
            '---------
        End If


        GeneraCalendarioReprogramacion = MatDatos

    End Function

    'Public Function GeneraCalendarioITF(ByVal pnPrestamo As Double, _
    '    ByVal pnCuotas As Integer, ByVal pnTasa As Double, ByVal pnDia As Integer, _
    '    ByVal pdFecDes As Date, ByVal pnCuotasGracia As Integer, _
    '    ByVal pnTipoGracia As Integer, ByVal pnGasto As Double, Optional ByVal pnProd As Integer = 0, Optional ByVal pnSubProd As Integer = 0, _
    '    Optional ByVal pbIncluyeGasto As Boolean = False, Optional ByVal pnTipoPeriodo As Integer = 1, Optional ByVal pnPlazo As Integer = 30) As Object

    '    Dim i As Integer
    '    Dim MatDatos(pnCuotas, 9) As String
    '    Dim nMonto As Double
    '    Dim nInteres As Double
    '    Dim dFecha As Date
    '    Dim MatFechas As Object
    '    Dim nMontoCuota As Double
    '    Dim nGasto As Double
    '    Dim nTasaGasto As Double
    '    Dim nIGV As Double
    '    Dim nIntGracia As Double
    '    Dim nDiasGracia As Integer

    '    nTasaGasto = Format(pnGasto / pnPrestamo * 100, "#0.00")

    '    nIGV = 0
    '    nGasto = 0

    '    If pbIncluyeGasto Then
    '        nGasto = Format(Format(nGasto, "#0.00") / pnCuotas, "#0.00")
    '    End If

    '    nMonto = pnPrestamo
    '    If pnCuotasGracia > 0 Then
    '        dFecha = DateAdd(DateInterval.Month, pnCuotasGracia, pdFecDes)
    '        nDiasGracia = DateDiff(DateInterval.Day, pdFecDes, dFecha)
    '    Else
    '        dFecha = pdFecDes
    '    End If

    '    '************************************************
    '    'SI ES FECHA FIJA
    '    '************************************************
    '    If pnTipoPeriodo = 1 Then
    '        MatFechas = CalendarioFechas(dFecha, pnDia, pnCuotas)
    '    Else
    '        '************************************************
    '        'SI ES PERIODO FIJO
    '        '************************************************
    '        'MatFechas = CalendarioFechasCuotaFija(dFecha, pnPlazo, pnCuotas)
    '    End If
    '    'Valida si Agrega Gracia Automaticamente
    '    'If DateDiff(DateInterval.Day, dFecha, MatFechas(1)) > 30 Then
    '    '    nDiasGracia = nDiasGracia + (DateDiff(DateInterval.Day, dFecha, MatFechas(1)) - 30)
    '    '    dFecha = DateAdd(DateInterval.Day, pnDiasGracia, pdFecDes)
    '    '    MatFechas = CalendarioFechas(dFecha, pnDia, pnCuotas)
    '    'End If

    '    'nMontoCUota = CuotaFechaFija(pdFecDes, pnTasa, nIGV, pnDiasGracia, pnCuotas, MatFechas, pnPrestamo)
    '    nMontoCuota = CuotaFechaFija(dFecha, pnTasa, nIGV, 0, pnCuotas, MatFechas, pnPrestamo)

    '    If nDiasGracia > 0 Then
    '        nIntGracia = ((1 + pnTasa / 100) ^ (nDiasGracia / 30)) - 1 'Interes
    '        If pnTipoGracia = 1 Then    'Proratear
    '            nIntGracia = Format((nIntGracia * nMonto) / pnCuotas, "#0.00")
    '        Else    'Primera Cuota
    '            nIntGracia = Format((nIntGracia * nMonto), "#0.00")
    '        End If
    '    End If

    '    For i = 1 To pnCuotas
    '        MatDatos(i, 1) = MatFechas(i) 'Fecha de Vencimiento
    '        MatDatos(i, 2) = Trim(Str(i)) 'Nro Cuota
    '        nInteres = ((1 + pnTasa / 100) ^ (DateDiff(DateInterval.Day, dFecha, CDate(MatFechas(i))) / 30)) - 1 'Interes
    '        nInteres = nInteres * nMonto 'Interes
    '        MatDatos(i, 5) = Format(nInteres, "#0.00") 'Interes            
    '        MatDatos(i, 7) = Format(nGasto, "#0.00") 'Gastos            
    '        'ARCV 07-12-2007
    '        If pnTipoGracia = 1 Then 'Proratear
    '            MatDatos(i, 9) = Format(nIntGracia, "#0.00") 'Interes de Gracia
    '        Else    'Primera Cuota
    '            If i = 1 Then
    '                MatDatos(i, 9) = Format(nIntGracia, "#0.00")
    '            Else
    '                MatDatos(i, 9) = "0.00"
    '            End If
    '        End If
    '        '--------
    '        MatDatos(i, 6) = "0.00"
    '        If i <> pnCuotas Then
    '            MatDatos(i, 4) = Format(nMontoCuota - (CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 6))), "#0.00") 'Capital
    '        Else
    '            If pnCuotas = 1 Then
    '                MatDatos(i, 4) = Format(nMonto, "#0.00")
    '            Else
    '                MatDatos(i, 4) = MatDatos(i - 1, 8)
    '            End If
    '        End If

    '        MatDatos(i, 3) = Format(CDbl(MatDatos(i, 5)) + CDbl(MatDatos(i, 4)) + CDbl(MatDatos(i, 6)) + CDbl(MatDatos(i, 7)) + CDbl(MatDatos(i, 9)), "#0.00") 'Cuota Total
    '        nMonto = Format(nMonto - CDbl(MatDatos(i, 4)), "#0.00")
    '        MatDatos(i, 8) = Format(nMonto, "#0.00") 'Saldo
    '        dFecha = CDate(MatDatos(i, 1))
    '    Next i

    '    GeneraCalendarioITF = MatDatos
    'End Function

    Public Function CuotaFechaFija(ByVal pdFecDes As Date, ByVal pTasaM As Double, _
            ByVal pnTasaSeguro As Double, ByVal psDiasGracia As Integer, _
            ByVal pnCuotas As Integer, ByVal pMatFechas As Object, ByVal vMonto As Double, _
            ByVal pnFinCuotaGracia As Integer, _
            Optional ByVal pnTasaDesgrav As Double = 0) As Double

        Dim I As Integer
        Dim SumTotal As Double
        Dim nTasaDiaria As Double
        Dim nValorActual As Double
        Dim nDiasTrans As Integer
        Dim dFecInicio As Date, dFecVenc As Date

        'nTasaDiaria = ((1 + (pTasaM / 100 * (1 + (pnTasaSeguro / 100)))) ^ (1 / 30) - 1) * 100
        'pTasaM = pTasaM * (1 + (pnTasaSeguro / 100))

        'ARCV 21-05-2009
        pTasaM = pTasaM * (1 + ((pnTasaSeguro + pnTasaDesgrav) / 100))
        '----------

        nTasaDiaria = (((1 + (pTasaM / 100)) ^ (1 / 30)) - 1) * 100


        SumTotal = 0
        If pnFinCuotaGracia > 0 Then
            dFecInicio = CDate(pMatFechas(pnFinCuotaGracia))
        Else
            dFecInicio = DateAdd(DateInterval.Day, psDiasGracia, pdFecDes)
        End If


        For I = pnFinCuotaGracia + 1 To pnCuotas
            dFecVenc = CDate(pMatFechas(I))
            nDiasTrans = DateDiff(DateInterval.Day, dFecInicio, dFecVenc)
            nValorActual = (1 / (1 + (nTasaDiaria / 100))) ^ (nDiasTrans)
            'nValorActual = (1 / (1 + (nTasaDiaria / 100))) ^ (nDiasTrans / 30)
            SumTotal = SumTotal + nValorActual
        Next I

        CuotaFechaFija = vMonto / SumTotal


        CuotaFechaFija = CDbl(Format(CuotaFechaFija, "#0.00"))



    End Function



    Public Function ESFeriado(ByVal pdFecha As Date, _
                            Optional ByVal pbVerificaFeriados As Boolean = False) As Boolean
        'Dim objCon As ComConecta.clsConecta
        'Dim oComm As SqlCommand
        'Dim nResul As Integer

        ''Carga Nombres de Columnas de Tablas
        'objCon = New ComConecta.clsConecta
        'oComm = New SqlCommand("Cred_Esferiado", objCon.DameConexion)
        'oComm.CommandType = CommandType.StoredProcedure
        'oComm.Parameters.Add("@dFecha", SqlDbType.DateTime).Value = pdFecha
        'oComm.Parameters.Add("@nResul", SqlDbType.Bit)
        'oComm.Parameters.Item("@nResul").Direction = ParameterDirection.Output
        'oComm.ExecuteNonQuery()
        'nResul = IIf(UCase(oComm.Parameters.Item("@nResul").Value.ToString) = "FALSE", 0, 1)

        'If nResul = 1 Then
        '    ESFeriado = True
        'Else
        '    ESFeriado = False
        'End If

        'oComm = Nothing
        'objCon.CierraConexion()
        'objCon = Nothing

        '------------ARCV 09-08-2010
        'If oFeriados.Find("dFecha", Format(pdFecha, "dd/MM/yyyy")) <> 0 Then
        '    ESFeriado = True
        'Else
        '    ESFeriado = False
        'End If
        'ESFeriado = False
        '--------------------------

        '------------ARCV 23-08-2010
        If pbVerificaFeriados Then
            If oFeriados.Find("dFecha", Format(pdFecha, "dd/MM/yyyy")) <> 0 Then
                ESFeriado = True
            Else
                ESFeriado = False
            End If
        Else
            ESFeriado = False
        End If
        '--------------------------
    End Function

    Public Function CalendarioFechas(ByVal pdFechaDes As Date, ByVal pnDia As Integer, _
            ByRef pnCuotas As Integer, ByVal pnTipoGracia As Integer, ByVal pnDiasGracia As Integer, _
                ByRef pnFinCuotaGracia As Integer, ByVal pdFechaPrimerPago As Date) As Object

        Dim I As Integer
        Dim M As Integer
        Dim Y As Integer
        Dim pdFecha As Date
        Dim MatFechas() As String
        Dim dFechaDesembTempo As Date
        Dim nDiasGraciaTemp As Integer
        Dim nDiasGraciaEval As Integer
        Dim K As Integer
        Dim bCambioMes As Boolean

        dFechaDesembTempo = pdFechaDes
        pnFinCuotaGracia = 0

        If pnTipoGracia <> 4 Or pnDiasGracia = 0 Then
            ReDim MatFechas(pnCuotas)
            K = 1
        Else 'SI ES GRACIA EN CUOTAS HALLAR EL TOTAL DE CUOTAS
            nDiasGraciaTemp = pnDiasGracia
            pnCuotas = Int((pnDiasGracia / 30) + pnCuotas)
            If Int((pnDiasGracia / 30)) = 0 Then
                pnCuotas = pnCuotas + 1
            End If
            ReDim MatFechas(pnCuotas)

            K = 1
            Do While nDiasGraciaTemp >= 30
                dFechaDesembTempo = DateAdd(DateInterval.Day, 30, dFechaDesembTempo)
                nDiasGraciaEval = 30
                If ESFeriado(dFechaDesembTempo) Then
                    dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    nDiasGraciaEval = nDiasGraciaEval + 1
                    If ESFeriado(dFechaDesembTempo) Then
                        dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                        nDiasGraciaEval = nDiasGraciaEval + 1
                        If ESFeriado(dFechaDesembTempo) Then
                            dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                            nDiasGraciaEval = nDiasGraciaEval + 1
                        End If
                    End If
                End If

                MatFechas(K) = Format(dFechaDesembTempo, "dd/MM/yyyy")

                nDiasGraciaTemp = nDiasGraciaTemp - nDiasGraciaEval
                K = K + 1
            Loop

            dFechaDesembTempo = DateAdd(DateInterval.Day, nDiasGraciaTemp, dFechaDesembTempo)
            If ESFeriado(dFechaDesembTempo) Then
                dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                If ESFeriado(dFechaDesembTempo) Then
                    dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    If ESFeriado(dFechaDesembTempo) Then
                        dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    End If
                End If
            End If
            If K = 1 Then
                K = 2
            End If
            MatFechas(K - 1) = Format(dFechaDesembTempo, "dd/MM/yyyy")

            'Numero de Cuota que acaba la Gracia
            pnFinCuotaGracia = K - 1
        End If

        pdFecha = dFechaDesembTempo
        pdFechaDes = dFechaDesembTempo
        'si fecha fija
        'pdFecha = pdFechaDes

        M = 0
        For I = K To pnCuotas
            bCambioMes = False
            If I = K Then
                M = Month(pdFecha)
                If pnDia <= Day(pdFecha) Then
                    M = M + 1
                End If
            End If

            Y = Year(pdFecha)
            If M > 12 Then
                M = M Mod 12
                Y = Y + 1
            End If

            'verificar que en febrero no pase de el dia 28 a una fecha fija
            If M = 2 And pnDia > 28 Then
                If Y Mod 4 <> 0 Then
                    'obtener nueva fecha para año no bisiesto
                    pdFecha = CDate("28/" + Str(M) + "/" + Str(Y))
                Else
                    'obtener nueva fecha para año bisiesto 2000,2004,2008,....
                    pdFecha = CDate("29/" + Str(M) + "/" + Str(Y))
                End If
            Else
                If (M = 4 Or M = 6 Or M = 9 Or M = 11) And (pnDia > 30) Then
                    'obtener nueva fecha valida para el mes
                    pdFecha = CDate("30/" + Str(M) + "/" + Str(Y))
                Else
                    'obtener nueva fecha
                    pdFecha = CDate(Str(pnDia) + "/" + Str(M) + "/" + Str(Y))
                End If
            End If

            'If I = K And DateDiff(DateInterval.Day, pdFechaDes, pdFecha) < 20 Then
            '    pdFecha = DateAdd(DateInterval.Month, 1, pdFecha)
            '    M = Month(pdFecha)
            '    Y = Year(pdFecha)
            'End If

            If I = K Then
                pdFecha = pdFechaPrimerPago
                M = Month(pdFecha)
                Y = Year(pdFecha)
            End If

            '-------------------  ARCV 09-08-2010
            'If Weekday(pdFecha) = 7 Then
            '    pdFecha = DateAdd(DateInterval.Day, 2, pdFecha)
            '    If M <> Month(pdFecha) Then
            '        bCambioMes = True
            '    Else
            '        bCambioMes = False
            '    End If
            '    M = Month(pdFecha)
            '    Y = Year(pdFecha)
            'End If

            'If Weekday(pdFecha) = 1 Then
            '    pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
            '    If M <> Month(pdFecha) Then
            '        bCambioMes = True
            '    Else
            '        bCambioMes = False
            '    End If
            '    M = Month(pdFecha)
            '    Y = Year(pdFecha)
            'End If
            '---------------------------------------------

            If ESFeriado(pdFecha) Then
                pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                If Not bCambioMes Then
                    If M <> Month(pdFecha) Then
                        bCambioMes = True
                    Else
                        bCambioMes = False
                    End If
                End If
                M = Month(pdFecha)
                Y = Year(pdFecha)
                If ESFeriado(pdFecha) Then
                    pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)

                    If Not bCambioMes Then
                        If M <> Month(pdFecha) Then
                            bCambioMes = True
                        Else
                            bCambioMes = False
                        End If
                    End If

                    M = Month(pdFecha)
                    Y = Year(pdFecha)
                    If ESFeriado(pdFecha) Then
                        pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                        If Not bCambioMes Then
                            If M <> Month(pdFecha) Then
                                bCambioMes = True
                            Else
                                bCambioMes = False
                            End If
                        End If
                        M = Month(pdFecha)
                        Y = Year(pdFecha)
                    End If
                End If
            End If


            MatFechas(I) = Format(pdFecha, "dd/MM/yyyy")


            'If (pnDiaFijo2) > 0 Then
            '    If m = 2 And pnDiaFijo2 > 28 Then
            '        If y Mod 4 <> 0 Then
            '            'obtener nueva fecha para año no bisiesto
            '            pdFecha = CDate("28/" + Str(m) + "/" + Str(y))
            '            pMatFecha(I) = pdFecha
            '        Else
            '            'obtener nueva fecha para año bisiesto 2000,2004,2008,....
            '            pdFecha = CDate("29/" + Str(m) + "/" + Str(y))
            '            pMatFecha(I) = pdFecha
            '        End If
            '    Else
            '        If (m = 4 Or m = 6 Or m = 9 Or m = 11) And (pnDiaFijo2 > 30) Then
            '            'obtener nueva fecha valida para el mes
            '            pdFecha = CDate("30/" + Str(m) + "/" + Str(y))
            '            pMatFecha(I) = pdFecha
            '        Else
            '            'obtener nueva fecha
            '            pdFecha = CDate(Str(pnDiaFijo2) + "/" + Str(m) + "/" + Str(y))
            '        End If
            '    End If
            '    pMatFecha(I + 1) = pdFecha
            '    I = I + 1
            'End If

            If Not bCambioMes Then
                M = M + 1
            End If

        Next I

        CalendarioFechas = MatFechas
    End Function

    Public Function CalendarioFechasCuotaFija(ByVal pdFechaDes As Date, ByVal pnPeriodo As Integer, _
                ByRef pnCuotas As Integer, ByVal pnTipoGracia As Integer, ByVal pnDiasGracia As Integer, _
                ByRef pnFinCuotaGracia As Integer, ByVal pdFechaPrimerPago As Date) As Object

        Dim I As Integer
        Dim Y As Integer
        Dim pdFecha As Date
        Dim MatFechas() As String
        Dim dFechaDesembTempo As Date
        Dim nDiasGraciaTemp As Integer
        Dim nDiasGraciaEval As Integer

        'ARCV 23-08-2010
        Dim bVerificaFeriados As Boolean = False
        If pnPeriodo = 1 Then bVerificaFeriados = True ' Para los credidiarios no generar cuotas los dias feriados
        '------------------

        dFechaDesembTempo = pdFechaDes
        pnFinCuotaGracia = 0

        If pnTipoGracia <> 4 Then
            ReDim MatFechas(pnCuotas)
            Y = 1
        Else 'SI ES GRACIA EN CUOTAS HALLAR EL TOTAL DE CUOTAS
            nDiasGraciaTemp = pnDiasGracia
            pnCuotas = Int((pnDiasGracia / 30) + pnCuotas)
            ReDim MatFechas(pnCuotas)

            Y = 1
            Do While nDiasGraciaTemp >= 30
                dFechaDesembTempo = DateAdd(DateInterval.Day, 30, dFechaDesembTempo)
                nDiasGraciaEval = 30
                If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                    dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    nDiasGraciaEval = nDiasGraciaEval + 1
                    If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                        dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                        nDiasGraciaEval = nDiasGraciaEval + 1
                        If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                            dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                            nDiasGraciaEval = nDiasGraciaEval + 1
                        End If
                    End If
                End If

                MatFechas(Y) = Format(dFechaDesembTempo, "dd/MM/yyyy")

                nDiasGraciaTemp = nDiasGraciaTemp - nDiasGraciaEval
                Y = Y + 1
            Loop

            dFechaDesembTempo = DateAdd(DateInterval.Day, nDiasGraciaTemp, dFechaDesembTempo)
            If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                    dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    If ESFeriado(dFechaDesembTempo, bVerificaFeriados) Then
                        dFechaDesembTempo = DateAdd(DateInterval.Day, 1, dFechaDesembTempo)
                    End If
                End If
            End If

            MatFechas(Y - 1) = Format(dFechaDesembTempo, "dd/MM/yyyy")

            'Numero de Cuota que acaba la Gracia
            pnFinCuotaGracia = Y - 1
        End If

        'pdFecha = pdFechaDes
        pdFecha = dFechaDesembTempo

        For I = Y To pnCuotas

            If I = Y Then
                pdFecha = pdFechaPrimerPago
            Else
                pdFecha = DateAdd(DateInterval.Day, pnPeriodo, pdFecha)
            End If


            'ARCV 09-08-2010
            'If Weekday(pdFecha) = 7 Then
            '    pdFecha = DateAdd(DateInterval.Day, 2, pdFecha)
            'End If

            'If Weekday(pdFecha) = 1 Then
            '    pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
            'End If
            '------------------------
            'ARCV 23-08-2010
            If pnPeriodo = 1 Then
                If Weekday(pdFecha) = 7 Then
                    pdFecha = DateAdd(DateInterval.Day, 2, pdFecha)
                End If
                If Weekday(pdFecha) = 1 Then
                    pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                End If
            End If
            '------------------------

            If ESFeriado(pdFecha, bVerificaFeriados) Then
                pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                If ESFeriado(pdFecha, bVerificaFeriados) Then
                    pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                    If ESFeriado(pdFecha, bVerificaFeriados) Then
                        pdFecha = DateAdd(DateInterval.Day, 1, pdFecha)
                    End If
                End If
            End If

            'ARCV 23-08-2010
            If pnPeriodo = 1 Then
                If Weekday(pdFecha) = 7 Then
                    pdFecha = DateAdd(DateInterval.Day, 2, pdFecha)
                End If
            End If
            '------------------------

            MatFechas(I) = Format(pdFecha, "dd/MM/yyyy")

        Next I

        CalendarioFechasCuotaFija = MatFechas
    End Function

    Public Function GeneraCalendarioDataSetSimulador(ByVal pnMontoGasto As Double, _
        ByVal poFeriado As ComTipos.clsDataset, ByVal pnPrestamo As Double, _
        ByVal pnCuotas As Integer, ByVal pnTasa As Double, ByVal pnDia As Integer, _
        ByVal pdFecDes As Date, ByVal pnDiasGracia As Integer, _
        ByVal pnTipoGracia As Integer, ByVal pnGasto As Double, ByVal pnProd As Integer, _
        ByVal pnSubProd As Integer, ByVal pnIncluyeGasto As Boolean, ByVal pnTipoPeriodo As Integer, ByVal pnPlazo As Integer, ByVal pnItfValor As Double, Optional ByVal pnSoloIntereses As Boolean = False, _
        Optional ByVal pnTasaDesgravamen As Double = 0, Optional ByRef pdFechaPrimPago As Date = Nothing, _
        Optional ByVal pnTipoDesembParcial As Boolean = False, Optional ByVal pMatrizDesemb As ComTipos.clsDataset = Nothing, _
        Optional ByRef MatCalenCuotasDesemParDetalle As ComTipos.clsDataset = Nothing) As ComTipos.clsDataset ' ARCV pnTasaDesgravamen (21-05-2009)

        Dim oCal As Object
        Dim i As Integer
        Dim D As ComTipos.clsDataset
        Dim Columnas(10) As String
        Dim nCursor As Integer

        oCal = GeneraCalendario(pnMontoGasto, poFeriado, pnPrestamo, pnCuotas, pnTasa, pnDia, _
                        pdFecDes, pnDiasGracia, pnTipoGracia, pnGasto, pnProd, pnSubProd, pnIncluyeGasto, pnTipoPeriodo, pnPlazo, pnItfValor, pnSoloIntereses, _
                        pnTasaDesgravamen, pnTipoDesembParcial, pMatrizDesemb, MatCalenCuotasDesemParDetalle, pdFechaPrimPago)

        Columnas(1) = "Vencimiento"
        Columnas(2) = "Cuota"
        Columnas(3) = "Monto"
        Columnas(4) = "Capital"
        Columnas(5) = "Interes"
        Columnas(6) = "IntGracia"
        Columnas(7) = "Impuesto"
        Columnas(8) = "Gasto"
        Columnas(9) = "Seguro"
        Columnas(10) = "Saldo"

        D = New ComTipos.clsDataset
        D.CrearVacio(Columnas)
        For i = 1 To UBound(oCal)
            nCursor = D.AddRegistro()
            D.Seek(nCursor)

            D.AsignarDato(Format(CDate(oCal(i, 1)), "yyyy/MM/dd"), "Vencimiento")
            D.AsignarDato(oCal(i, 2), "Cuota")
            D.AsignarDato(oCal(i, 3), "Monto")
            D.AsignarDato(oCal(i, 4), "Capital")
            D.AsignarDato(oCal(i, 5), "Interes")
            D.AsignarDato(oCal(i, 9), "IntGracia")
            D.AsignarDato(oCal(i, 6), "Impuesto")
            D.AsignarDato(oCal(i, 7), "Gasto")
            D.AsignarDato(oCal(i, 10), "Seguro")
            D.AsignarDato(oCal(i, 8), "Saldo")

        Next

        GeneraCalendarioDataSetSimulador = D

        D = Nothing
        oCal = Nothing

    End Function

End Class

