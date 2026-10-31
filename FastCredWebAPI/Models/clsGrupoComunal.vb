Imports System.Data.Entity
Imports System.Data.Entity.Validation
Imports System.Data.SqlClient
Imports FastCredWebAPI.FastCredWebAPI
Public Class clsGrupoComunal
    Private db As New DBNEGOCIOEntities
    Public Property nIdGrupo As Long
    Public Property cNombre As String
    Public Property sCodUsuCob As String
    Public Property sNomAge As String
    Public Property cLugarProxReunion As String
    Public Property dFechaReunion As Date
    Public Property nNumSocios As Integer
    Public Property cNomEstado As String
    Public Property nCodEstado As Integer
    Public Property nNumCiclo As Integer
    Public Property cDNISocio As String
    Public Property ListaSocios As IQueryable(Of clsGrupoComunalSocio)
    Public Property nDiasAtraso As Integer
    Public Property dFecVenc As Date
    Public Property nNumCuotas As Integer
    Public Property ListaCargosSocios As IQueryable(Of CatalogoCodigos)
    Public Property nFormaPago As Integer

    Public Sub CargaDatos(ByVal pnIdGrupo As Long)

        Dim sDato As String
        sDato = ""

        Dim vGrupos = From U In db.GruposComunales
                      Join CC In db.CatalogoCodigos On CC.cValor Equals U.nCodEstado
                      Where U.nID = pnIdGrupo And CC.nCodigo = 1025
                      Select U.nID, U.cNombre, U.cLugarProximaReunion, U.nNumSocios, CC.cDescripcion, U.nNumCiclo, U.cCodUsuCob

        For Each m In vGrupos
            nIdGrupo = m.nID
            cNombre = m.cNombre
            cLugarProxReunion = m.cLugarProximaReunion
            nNumSocios = m.nNumSocios
            cNomEstado = m.cDescripcion
            nNumCiclo = m.nNumCiclo
            sCodUsuCob = m.cCodUsuCob
        Next

        Dim vGriposDetalle = From U In db.GruposComunalesDetalle
                             Where U.nIdGrupo = pnIdGrupo
                             Select U.cDNI

        nNumSocios = vGriposDetalle.Count

        Dim vCargosSocios = From CC In db.CatalogoCodigos
                            Where CC.nCodigo = 1026 And CC.nCodigo <> CC.cValor
                            Select CC

        ListaCargosSocios = vCargosSocios


        Dim vSocios = From U In db.vGCSocios
                      Where U.nIdGrupo = pnIdGrupo
                      Select New clsGrupoComunalSocio With {
                                                  .cDNI = U.cDNI, .cApellidos = U.cApellidos, .cNombres = U.cNombres,
                          .nCodAge = U.nCodAge, .cCodCargo = U.cCodCargo, .nCodCred = U.nCodCred,
                          .nDiasAtraso = U.nDiasAtraso, .nMontoAhorroProgramado = U.nMontoAhorro, .nMontoCuota = U.nMontoCuota,
                          .nMontoPrestamo = U.nPrestamo, .sCuotas = "", .nNumCuotas = U.nCuotas, .cEstado = U.cDescripcion1, .cNomCargo = U.cDescripcion2, .nSaldoK = U.nSaldoK, .nIdGrupo = U.nIdGrupo
                          }

        ListaSocios = vSocios


        For Each d In ListaSocios
            If Not IsNothing(d.nCodCred) Then
                Dim vCalendario = From CAL In db.CredCalendario
                                  Where CAL.nCodCred = d.nCodCred And CAL.nEstado = 0
                                  Group By CAL.nCodCred Into myGroup = Group
                                  Select dFecVenc = myGroup.Min(Function(CAL) CAL.dFecVenc), nNroCuotaMin = myGroup.Min(Function(CAL) CAL.nNroCuota)


                For Each C In vCalendario
                    d.nNroCuotaMin = C.nNroCuotaMin
                    d.dFecVenc = C.dFecVenc
                    d.sCuotas = Str(C.nNroCuotaMin) + "/" + Str(d.nNumCuotas)
                    Exit For
                Next
            End If
        Next

        'Dim vCtaInt = From U In db.GruposComunalesCuentaInterna
        '              Where U.nIdGrupo = pnIdGrupo
        '              Select U


        'ListaCtaInterna = vCtaInt

    End Sub

    Public Sub CargaDatos_A_Aprobar(ByVal pnIdGrupo As Long)

        Dim sDato As String
        sDato = ""

        Dim vGrupos = From U In db.GruposComunales
                      Join CC In db.CatalogoCodigos On CC.cValor Equals U.nCodEstado
                      Where U.nID = pnIdGrupo And CC.nCodigo = 1025
                      Select U.nID, U.cNombre, U.cLugarProximaReunion, U.nNumSocios, CC.cDescripcion, U.nNumCiclo

        For Each m In vGrupos
            nIdGrupo = m.nID
            cNombre = m.cNombre
            cLugarProxReunion = m.cLugarProximaReunion
            nNumSocios = m.nNumSocios
            cNomEstado = m.cDescripcion
            nNumCiclo = m.nNumCiclo
        Next

        Dim vCargosSocios = From CC In db.CatalogoCodigos
                            Where CC.nCodigo = 1026 And CC.nCodigo <> CC.cValor
                            Select CC

        ListaCargosSocios = vCargosSocios

        Dim vSocios = From U In db.GruposComunales
                      Join GD In db.GruposComunalesDetalle On GD.nIdGrupo Equals U.nID
                      Join P In db.Persona On P.cDoc Equals GD.cDNI
                      From CP In db.CredPersonas.Where(Function(CP) CP.nCodPers = P.nCodPers And CP.nRelacion = 10).DefaultIfEmpty()
                      From C In db.Creditos.Where(Function(C) C.nCodAge = CP.nCodAge And C.nCodCred = CP.nCodCred).DefaultIfEmpty()
                      From CC In db.CredCalendCond.Where(Function(CC) CC.nCodAge = C.nCodAge And CC.nCodCred = C.nCodCred).DefaultIfEmpty()
                      From CC1 In db.CatalogoCodigos.Where(Function(CC1) CC1.nCodigo = 1005 And CC1.cValor = C.nEstado).DefaultIfEmpty()
                      From CC2 In db.CatalogoCodigos.Where(Function(CC2) CC2.nCodigo = 1026 And CC2.cValor = GD.cCodCargo).DefaultIfEmpty()
                      Where U.nID = pnIdGrupo And (C.nEstado = 1 Or C.nEstado = 3)
                      Select New clsGrupoComunalSocio With {
                                                  .cDNI = GD.cDNI, .cApellidos = P.cApellidos, .cNombres = P.cNombres,
                          .nCodAge = C.nCodAge, .cCodCargo = GD.cCodCargo, .nCodCred = C.nCodCred,
                          .nDiasAtraso = C.nDiasAtraso, .nMontoAhorroProgramado = GD.nMontoAhorro, .nMontoCuota = CC.nMontoCuota,
                          .nMontoPrestamo = C.nPrestamo, .sCuotas = "", .nNumCuotas = CC.nCuotas, .cEstado = CC1.cDescripcion, .cNomCargo = CC2.cDescripcion, .nSaldoK = C.nSaldoK, .nIdGrupo = GD.nIdGrupo,
                          .nCodEstado = C.nEstado
                                            }



        ListaSocios = vSocios


        For Each d In ListaSocios
            If Not IsNothing(d.nCodCred) Then
                Dim vCalendario = From CAL In db.CredCalendario
                                  Where CAL.nCodCred = d.nCodCred And CAL.nEstado = 0
                                  Group By CAL.nCodCred Into myGroup = Group
                                  Select dFecVenc = myGroup.Min(Function(CAL) CAL.dFecVenc), nNroCuotaMin = myGroup.Min(Function(CAL) CAL.nNroCuota)


                For Each C In vCalendario
                    d.nNroCuotaMin = C.nNroCuotaMin
                    d.dFecVenc = C.dFecVenc
                    d.sCuotas = Str(C.nNroCuotaMin) + "/" + Str(d.nNumCuotas)
                    Exit For
                Next
            End If
        Next

        'Dim vCtaInt = From U In db.GruposComunalesCuentaInterna
        '              Where U.nIdGrupo = pnIdGrupo
        '              Select U


        'ListaCtaInterna = vCtaInt

    End Sub
    Public Sub CargaDatos_A_Pagar(ByVal pnIdGrupo As Long)

        Dim sDato As String
        sDato = ""

        Dim vGrupos = From U In db.GruposComunales
                      Join CC In db.CatalogoCodigos On CC.cValor Equals U.nCodEstado
                      Where U.nID = pnIdGrupo And CC.nCodigo = 1025
                      Select U.nID, U.cNombre, U.cLugarProximaReunion, U.nNumSocios, CC.cDescripcion, U.nNumCiclo

        For Each m In vGrupos
            nIdGrupo = m.nID
            cNombre = m.cNombre
            cLugarProxReunion = m.cLugarProximaReunion
            nNumSocios = m.nNumSocios
            cNomEstado = m.cDescripcion
            nNumCiclo = m.nNumCiclo
        Next

        Dim vCargosSocios = From CC In db.CatalogoCodigos
                            Where CC.nCodigo = 1026 And CC.nCodigo <> CC.cValor
                            Select CC

        ListaCargosSocios = vCargosSocios

        Dim vSocios = From U In db.GruposComunales
                      Join GD In db.GruposComunalesDetalle On GD.nIdGrupo Equals U.nID
                      Join P In db.Persona On P.cDoc Equals GD.cDNI
                      From CP In db.CredPersonas.Where(Function(CP) CP.nCodPers = P.nCodPers And CP.nRelacion = 10).DefaultIfEmpty()
                      From C In db.Creditos.Where(Function(C) C.nCodAge = CP.nCodAge And C.nCodCred = CP.nCodCred).DefaultIfEmpty()
                      From CC In db.CredCalendCond.Where(Function(CC) CC.nCodAge = C.nCodAge And CC.nCodCred = C.nCodCred).DefaultIfEmpty()
                      From CC1 In db.CatalogoCodigos.Where(Function(CC1) CC1.nCodigo = 1005 And CC1.cValor = C.nEstado).DefaultIfEmpty()
                      From CC2 In db.CatalogoCodigos.Where(Function(CC2) CC2.nCodigo = 1026 And CC2.cValor = GD.cCodCargo).DefaultIfEmpty()
                      Where U.nID = pnIdGrupo And C.nEstado = 4
                      Select New clsGrupoComunalSocio With {
                                                  .cDNI = GD.cDNI, .cApellidos = P.cApellidos, .cNombres = P.cNombres,
                          .nCodAge = C.nCodAge, .cCodCargo = GD.cCodCargo, .nCodCred = C.nCodCred,
                          .nDiasAtraso = C.nDiasAtraso, .nMontoAhorroProgramado = GD.nMontoAhorro, .nMontoCuota = CC.nMontoCuota,
                          .nMontoPrestamo = C.nPrestamo, .sCuotas = "", .nNumCuotas = CC.nCuotas, .cEstado = CC1.cDescripcion, .cNomCargo = CC2.cDescripcion, .nSaldoK = C.nSaldoK, .nIdGrupo = GD.nIdGrupo,
                          .nCodEstado = C.nEstado
                                            }



        ListaSocios = vSocios


        For Each d In ListaSocios
            If Not IsNothing(d.nCodCred) Then
                Dim vCalendario = From CAL In db.CredCalendario
                                  Where CAL.nCodCred = d.nCodCred And CAL.nEstado = 0
                                  Group By CAL.nCodCred Into myGroup = Group
                                  Select dFecVenc = myGroup.Min(Function(CAL) CAL.dFecVenc), nNroCuotaMin = myGroup.Min(Function(CAL) CAL.nNroCuota)


                For Each C In vCalendario
                    d.nNroCuotaMin = C.nNroCuotaMin
                    d.dFecVenc = C.dFecVenc
                    d.sCuotas = Str(C.nNroCuotaMin) + "/" + Str(d.nNumCuotas)
                    Exit For
                Next
            End If
        Next

        'Dim vCtaInt = From U In db.GruposComunalesCuentaInterna
        '              Where U.nIdGrupo = pnIdGrupo
        '              Select U


        'ListaCtaInterna = vCtaInt

    End Sub

    Public Sub CargaDatos_A_Desembolsar(ByVal pnIdGrupo As Long)

        Dim sDato As String
        sDato = ""

        Dim vGrupos = From U In db.GruposComunales
                      Join CC In db.CatalogoCodigos On CC.cValor Equals U.nCodEstado
                      Where U.nID = pnIdGrupo And CC.nCodigo = 1025
                      Select U.nID, U.cNombre, U.cLugarProximaReunion, U.nNumSocios, CC.cDescripcion, U.nNumCiclo

        For Each m In vGrupos
            nIdGrupo = m.nID
            cNombre = m.cNombre
            cLugarProxReunion = m.cLugarProximaReunion
            nNumSocios = m.nNumSocios
            cNomEstado = m.cDescripcion
            nNumCiclo = m.nNumCiclo
        Next

        Dim vCargosSocios = From CC In db.CatalogoCodigos
                            Where CC.nCodigo = 1026 And CC.nCodigo <> CC.cValor
                            Select CC

        ListaCargosSocios = vCargosSocios

        Dim vSocios = From U In db.GruposComunales
                      Join GD In db.GruposComunalesDetalle On GD.nIdGrupo Equals U.nID
                      Join P In db.Persona On P.cDoc Equals GD.cDNI
                      From CP In db.CredPersonas.Where(Function(CP) CP.nCodPers = P.nCodPers And CP.nRelacion = 10).DefaultIfEmpty()
                      From C In db.Creditos.Where(Function(C) C.nCodAge = CP.nCodAge And C.nCodCred = CP.nCodCred).DefaultIfEmpty()
                      From CC In db.CredCalendCond.Where(Function(CC) CC.nCodAge = C.nCodAge And CC.nCodCred = C.nCodCred).DefaultIfEmpty()
                      From CC1 In db.CatalogoCodigos.Where(Function(CC1) CC1.nCodigo = 1005 And CC1.cValor = C.nEstado).DefaultIfEmpty()
                      From CC2 In db.CatalogoCodigos.Where(Function(CC2) CC2.nCodigo = 1026 And CC2.cValor = GD.cCodCargo).DefaultIfEmpty()
                      Where U.nID = pnIdGrupo And C.nEstado = 2
                      Select New clsGrupoComunalSocio With {
                                                  .cDNI = GD.cDNI, .cApellidos = P.cApellidos, .cNombres = P.cNombres,
                          .nCodAge = C.nCodAge, .cCodCargo = GD.cCodCargo, .nCodCred = C.nCodCred,
                          .nDiasAtraso = C.nDiasAtraso, .nMontoAhorroProgramado = GD.nMontoAhorro, .nMontoCuota = CC.nMontoCuota,
                          .nMontoPrestamo = C.nPrestamo, .sCuotas = "", .nNumCuotas = CC.nCuotas, .cEstado = CC1.cDescripcion, .cNomCargo = CC2.cDescripcion, .nSaldoK = C.nSaldoK, .nIdGrupo = GD.nIdGrupo,
                          .nCodEstado = C.nEstado
                                            }



        ListaSocios = vSocios


        For Each d In ListaSocios
            If Not IsNothing(d.nCodCred) Then
                Dim vCalendario = From CAL In db.CredCalendario
                                  Where CAL.nCodCred = d.nCodCred And CAL.nEstado = 0
                                  Group By CAL.nCodCred Into myGroup = Group
                                  Select dFecVenc = myGroup.Min(Function(CAL) CAL.dFecVenc), nNroCuotaMin = myGroup.Min(Function(CAL) CAL.nNroCuota)


                For Each C In vCalendario
                    d.nNroCuotaMin = C.nNroCuotaMin
                    d.dFecVenc = C.dFecVenc
                    d.sCuotas = Str(C.nNroCuotaMin) + "/" + Str(d.nNumCuotas)
                    Exit For
                Next
            End If
        Next

        'Dim vCtaInt = From U In db.GruposComunalesCuentaInterna
        '              Where U.nIdGrupo = pnIdGrupo
        '              Select U


        'ListaCtaInterna = vCtaInt

    End Sub

    Public Sub CargaSocio()
        Dim vSocios = From U In db.GruposComunales
                      Join GD In db.GruposComunalesDetalle On GD.nIdGrupo Equals U.nID
                      Join P In db.Persona On P.cDoc Equals GD.cDNI
                      From CP In db.CredPersonas.Where(Function(CP) CP.nCodPers = P.nCodPers And CP.nRelacion = 10).DefaultIfEmpty()
                      From C In db.Creditos.Where(Function(C) C.nCodAge = CP.nCodAge And C.nCodCred = CP.nCodCred).DefaultIfEmpty()
                      From CC In db.CredCalendCond.Where(Function(CC) CC.nCodAge = C.nCodAge And CC.nCodCred = C.nCodCred).DefaultIfEmpty()
                      Where U.nID = nIdGrupo And (C.nEstado = 1 Or C.nEstado = 4) And GD.cDNI = cDNISocio
                      Select New clsGrupoComunalSocio With {
                                                  .cDNI = GD.cDNI, .cApellidos = P.cApellidos, .cNombres = P.cNombres,
                          .nCodAge = C.nCodAge, .cCodCargo = GD.cCodCargo, .nCodCred = C.nCodCred,
                          .nDiasAtraso = C.nDiasAtraso, .nMontoAhorroProgramado = GD.nMontoAhorro, .nMontoCuota = CC.nMontoCuota,
                          .nMontoPrestamo = C.nPrestamo, .sCuotas = "", .nNumCuotas = CC.nCuotas
                                            }

        ListaSocios = vSocios

        For Each d In ListaSocios
            If Not IsNothing(d.nCodCred) Then
                Dim vCalendario = From CAL In db.CredCalendario
                                  Where CAL.nCodCred = d.nCodCred And CAL.nEstado = 0
                                  Group By CAL.nCodCred Into myGroup = Group
                                  Select dFecVenc = myGroup.Min(Function(CAL) CAL.dFecVenc), nNroCuotaMin = myGroup.Min(Function(CAL) CAL.nNroCuota)


                For Each C In vCalendario
                    d.nNroCuotaMin = C.nNroCuotaMin
                    d.dFecVenc = C.dFecVenc
                    d.sCuotas = Str(C.nNroCuotaMin) + "/" + Str(d.nNumCuotas)
                    Exit For
                Next
            End If
        Next

    End Sub

    Public Sub RegistraGrupoComunal(ByRef nResp As Integer, ByRef sResp As String)

        Dim GC As GruposComunales


        Try

            'Registra Grupo Comunal

            GC = New GruposComunales

            Dim vGC = From U In db.GruposComunales
                      Where U.nID = nIdGrupo
                      Select U

            If vGC.Count > 0 Then

                db.Configuration.AutoDetectChangesEnabled = False

                GC = db.GruposComunales.First(Function(p) p.nID = nIdGrupo)
                db.GruposComunales.Attach(GC)
                GC.cNombre = cNombre
                db.Entry(GC).Property("cNombre").IsModified = True

                Try
                    db.SaveChanges()

                Catch ex As Exception
                    nResp = 0
                    sResp = ex.Message
                End Try

                db.Configuration.AutoDetectChangesEnabled = True

            Else

                GC.cNombre = cNombre
                GC.dFechaReg = Now()
                GC.cCodUsuCob = sCodUsuCob
                GC.cLugarProximaReunion = ""
                GC.dFechaProxReunion = Now().Date
                GC.nNumSocios = 0
                GC.nCodEstado = 1
                GC.bActivo = True
                GC.nNumCiclo = 0

                db.GruposComunales.Add(GC)


                Try
                    db.SaveChanges()
                Catch ex As DbEntityValidationException
                    nResp = 0
                    sResp = ex.EntityValidationErrors(0).ValidationErrors(0).ErrorMessage

                End Try


            End If

            db.Database.Connection.Close()


            'Actualiza Tipo de Forma de Pago Mensual o Catorcenal



            Dim oConex As SqlConnection
            Dim oConect As New clsConecta
            Dim sCadConexion As String
            Dim C As SqlCommand
            Dim oComm As SqlCommand


            sCadConexion = oConect.gsCadConex

            oConex = New SqlConnection(sCadConexion)

            oConex.Open()
            C = oConex.CreateCommand()
            C.CommandText = "SET DATEFORMAT DMY"
            C.ExecuteNonQuery()
            C = Nothing

            nResp = GC.nID
            sResp = "Registro Correcto"

            Try
                oComm = New SqlClient.SqlCommand("PA_GruposComunalesActualizaFormaPago", oConex)
                oComm.Parameters.Add("@nIdGrupo", SqlDbType.Int).Value = GC.nID
                oComm.Parameters.Add("@nFormaPago", SqlDbType.Int).Value = nFormaPago
                oComm.CommandTimeout = 0
                oComm.CommandType = CommandType.StoredProcedure
                oComm.ExecuteNonQuery()

                oConex.Close()
            Catch ex As Exception
                sResp = ex.Message
            End Try





        Catch ex As Exception
            db.Database.Connection.Close()
        End Try

    End Sub


End Class
