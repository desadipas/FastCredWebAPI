Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports FastCredWebAPI.FastCredWebAPI

Public Class CuadreCajaUsuario
    Private db As New DBNEGOCIOEntities
    Public Property nCodAge As Integer
    Public Property cCodUsu As String
    Public Property nSaldoInicial As Double
    Public Property nIngresos As Double
    Public Property nEgresos As Double
    Public Property nSaldoFinal As Double
    Public Property nDineroEnCaja As Double
    Public Property nDiferencia As Double
    Public Property sMensaje As String
    Public Property ListaIngresosEgresos As IQueryable(Of clsIngresoEgreso)
    Public Property ListaDepositos As IQueryable(Of Transac)
    'Public Property ListaDepositos As String()

    Private Function UnicodeBytesToString(
    ByVal bytes() As Byte) As String

        Return System.Text.Encoding.Unicode.GetString(bytes)
    End Function
    Public Sub CargaData(ByVal psCodUsu As String, ByVal pdFecha As Date)

        Dim oBD As New clsBaseDatos
        Dim oFila As clsIngresoEgreso
        Dim vData As DataTable = oBD.ListaIngresos(psCodUsu, pdFecha)
        Dim oDatos() As clsIngresoEgreso
        Dim i As Integer


        i = 0

        cCodUsu = psCodUsu

        Dim vUsuario = From U In db.Usuarios
                       Where U.cCodUsu = psCodUsu
                       Select U

        For Each m In vUsuario
            nCodAge = m.nCodAge
            Exit For
        Next

        Dim vDepositos = From U In db.Transac
                         Where U.cUser = psCodUsu And U.nCodOpe = 4 And U.bExtornado = 0 And U.dFecha.Day = pdFecha.Day And U.dFecha.Month = pdFecha.Month And U.dFecha.Year = pdFecha.Year
                         Select U

        ListaDepositos = vDepositos



        i = 0

        For Each d As DataRow In vData.Rows
            oFila = New clsIngresoEgreso
            oFila.cOperacion = d("OPERACION")
            oFila.cTipoOperacion = "I"
            oFila.nMonto = d("MONTO")
            ReDim Preserve oDatos(i)
            oDatos(i) = oFila
            i = i + 1

        Next

        vData = oBD.ListaEgresos(psCodUsu, pdFecha)

        For Each d As DataRow In vData.Rows
            oFila = New clsIngresoEgreso
            oFila.cOperacion = d("OPERACION")
            oFila.cTipoOperacion = "E"
            oFila.nMonto = d("MONTO")

            ReDim Preserve oDatos(i)
            oDatos(i) = oFila
            i = i + 1
        Next


        If IsNothing(oDatos) Then
            ListaIngresosEgresos = Nothing
        Else
            ListaIngresosEgresos = oDatos.AsQueryable()
        End If


        vData = oBD.DevuelveSaldoCajaAnterior(psCodUsu, pdFecha)
        nSaldoInicial = 0
        For Each d As DataRow In vData.Rows
            nSaldoInicial = d("nSaldoEfectivo")
        Next

        vData = oBD.DevuelveSaldoCajaHoy(psCodUsu, pdFecha)
        nDineroEnCaja = 0
        For Each d As DataRow In vData.Rows
            nDineroEnCaja = d("nSaldoEfectivo")
        Next

        nIngresos = 0
        nEgresos = 0

        If Not IsNothing(ListaIngresosEgresos) Then
            For Each F In ListaIngresosEgresos
                If Not IsNothing(F) Then
                    If F.cTipoOperacion.Trim() = "I" Then
                        nIngresos = nIngresos + F.nMonto
                    Else
                        nEgresos = nEgresos + F.nMonto
                    End If
                End If

            Next
        End If

        nSaldoFinal = nSaldoInicial + nIngresos - nEgresos

        nDiferencia = nDineroEnCaja - nSaldoFinal

        If nDiferencia < 0 Then
            sMensaje = "FALTANTE"
        End If
        If nDiferencia > 0 Then
            sMensaje = "SOBRANTE"
        End If
        If nDiferencia = 0 Then
            sMensaje = "SIN DIFERENCIAS"
        End If

    End Sub

End Class
