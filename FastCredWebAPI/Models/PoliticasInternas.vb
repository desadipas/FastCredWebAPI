Imports System.Data.Entity.Core.Objects
Imports FastCredWebAPI.FastCredWebAPI

Public Class PoliticasInternas
    Public ReglasdeNegocio() As sistNegocioReglas
    Public ReglasdeNegocioParam() As sistNegocioReglasParam
    Public ReglasdeNegocioDependencias() As sistNegocioReglasDependencias

    Public ValoresPantalla() As SistFormularios

    Private db As New DBNEGOCIOEntities
    Public Function ValorControlPantalla(ByVal psNomControl As String) As String


        Dim dReglasDepend = From U In ValoresPantalla
                            Where U.cNomControl = psNomControl
                            Select U

        If dReglasDepend.Count() > 0 Then
            ValorControlPantalla = dReglasDepend.ElementAt(0).cValor.Trim()
        Else
            ValorControlPantalla = ""
        End If

    End Function
    Public Function ValidaPoliticas(ByRef psResultado As String) As Boolean
        Dim i As Integer
        Dim bResp As Boolean
        Dim bRespDepend As Boolean

        ValidaPoliticas = True

        Dim dReglasNegocio = From U In db.sistNegocioReglas
                             Join F In db.sistFunciones On F.cNomFuncion Equals U.cFuncionAValidar
                             Where U.cNomForm = "persona" And F.bEjecucionServidor = True
                             Select U

        ReglasdeNegocio = dReglasNegocio.ToArray

        Dim dReglasNegocioParam = From U In db.sistNegocioReglasParam
                                  Join NR In db.sistNegocioReglas On U.nIDPadre Equals NR.nID
                                  Join F In db.sistFunciones On F.cNomFuncion Equals NR.cFuncionAValidar
                                  Where NR.cNomForm = "persona" And F.bEjecucionServidor = True
                                  Select U

        ReglasdeNegocioParam = dReglasNegocioParam.ToArray

        Dim dReglasNegocioDependencias = From U In db.sistNegocioReglasDependencias
                                         Select U

        ReglasdeNegocioDependencias = dReglasNegocioDependencias.ToArray

        For i = 0 To ReglasdeNegocio.Count() - 1
            If ReglasdeNegocio(i).bDependencia = False Then
                bRespDepend = ValidaPoliticasDependencias(ReglasdeNegocio(i).nID)
                If bRespDepend Then
                    ValidaPoliticas = ValidaPolitica(psResultado, ReglasdeNegocio(i))
                    If Not ValidaPoliticas Then
                        Exit Function
                    End If
                End If
            End If
        Next


    End Function
    Public Function ValidaPoliticasDependencias(ByVal pnIDPadre As Long) As Boolean
        Dim psResultado As String = ""
        Dim i As Integer

        ValidaPoliticasDependencias = True

        Dim dReglasDepend = From U In ReglasdeNegocioDependencias
                            Join R In ReglasdeNegocio On R.nID Equals U.nIDReglaDep
                            Where U.nIDRegla = pnIDPadre
                            Select R

        For i = 0 To dReglasDepend.Count() - 1
            ValidaPoliticasDependencias = ValidaPolitica(psResultado, dReglasDepend(i))
            If Not ValidaPoliticasDependencias Then
                Exit Function
            End If
        Next


    End Function

    Public Function ValorDeControl(ByVal psParamBusq As String, ByVal psFuncionBusq As String, ByVal nIDPadre As Integer) As String
        Dim i As Integer

        Dim dControl = From U In ReglasdeNegocioParam
                       Join P In ReglasdeNegocio On P.nID Equals U.nIDPadre
                       Where U.cParamFunc = psParamBusq And P.cFuncionAValidar = psFuncionBusq And P.nID = nIDPadre
                       Select U

        Dim sNomControl As String
        sNomControl = dControl.ElementAt(0).cNomControl

        Dim ValoresP = From U In ValoresPantalla
                       Where U.cNomControl = sNomControl
                       Select U

        ValorDeControl = ValoresP.ElementAt(0).cValor


    End Function

    Public Function CalificacionSistemaFinanciero(ByVal psValorPantalla As String, ByVal psOperador As String, ByVal psValorAValidar As String) As Boolean

        Dim pnValor As Integer

        Dim dCalif = From U In db.vRCCTotal
                     Where U.Cod_Doc_Id.Trim() = psValorPantalla.Trim()
                     Select U

        pnValor = dCalif.First.cCalificacion

        Select Case psOperador
            Case ">"
                If pnValor > CInt(psValorAValidar) Then
                    CalificacionSistemaFinanciero = True
                Else
                    CalificacionSistemaFinanciero = False
                End If
            Case ">="
                If pnValor >= CInt(psValorAValidar) Then
                    CalificacionSistemaFinanciero = True
                Else
                    CalificacionSistemaFinanciero = False
                End If

            Case "="
                If pnValor = CInt(psValorAValidar) Then
                    CalificacionSistemaFinanciero = True
                Else
                    CalificacionSistemaFinanciero = False
                End If

            Case "<="
                If pnValor <= CInt(psValorAValidar) Then
                    CalificacionSistemaFinanciero = True
                Else
                    CalificacionSistemaFinanciero = False
                End If

            Case "<"
                If pnValor < CInt(psValorAValidar) Then
                    CalificacionSistemaFinanciero = True
                Else
                    CalificacionSistemaFinanciero = False
                End If

            Case Else
                CalificacionSistemaFinanciero = False
        End Select

    End Function

    Public Function ValidaPolitica(ByRef psResultado As String, ByVal oPolitica As sistNegocioReglas) As Boolean
        Dim psValores As String
        Dim P1 As ObjectParameter

        psValores = ValorDeControl("psDoc", oPolitica.cFuncionAValidar, oPolitica.nID)

        P1 = New ObjectParameter("psResultado", "")
        db.PA_ValidaPolitica(P1, oPolitica.cOperador, oPolitica.cValorPolitica, oPolitica.cFuncionAValidar, psValores, oPolitica.cNomValidacion)
        psResultado = P1.Value

        If UCase(psResultado) = "OK" Then
            ValidaPolitica = True
        Else
            ValidaPolitica = False
        End If

    End Function
End Class
