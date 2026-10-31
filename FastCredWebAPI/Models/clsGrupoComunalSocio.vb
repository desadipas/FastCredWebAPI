Imports FastCredWebAPI.FastCredWebAPI
Public Class clsGrupoComunalSocio
    Private db As New DBNEGOCIOEntities

    Public Property nIdGrupo As Integer
    Public Property cDNI As String
    Public Property cApellidos As String
    Public Property cNombres As String
    Public Property nMontoAhorroProgramado As Double
    Public Property cCodCargo As String
    Public Property nCodCred As Integer
    Public Property nCodAge As Integer
    Public Property nMontoPrestamo As Integer
    Public Property sCuotas As String
    Public Property nNumCuotas As Integer
    Public Property nNroCuotaMin As Integer
    Public Property nMontoCuota As Double
    Public Property dFecVenc As Date
    Public Property nDiasAtraso As Integer
    Public Property cEstado As String
    Public Property nSaldoK As Double
    Public Property cNomCargo As String
    Public Property nCodEstado As Integer
    Public Sub EliminaSocio(ByRef nResp As Integer, ByRef sResp As String)
        Dim Socio As GruposComunalesDetalle

        Dim vSocio = db.GruposComunalesDetalle.First(Function(p) p.nIdGrupo = nIdGrupo And p.cDNI = cDNI)
        Socio = vSocio

        If Not IsNothing(vSocio) Then
            Try

                db.GruposComunalesDetalle.Remove(Socio)
                db.SaveChanges()
                nResp = 1
                sResp = "Proceso Correcto"

            Catch ex As Exception
                nResp = 0
                sResp = ex.Message
            End Try

        Else
            nResp = 0
            sResp = "No se encontró al Socio"
        End If




    End Sub


    Public Sub RegistraSocio(ByRef nResp As Integer, ByRef sResp As String)

        Dim vSocio = From S In db.GruposComunalesDetalle
                     Where S.nIdGrupo = nIdGrupo And S.cDNI = cDNI
                     Select S

        Dim oSocio As New GruposComunalesDetalle

        If vSocio.Count > 0 Then
            db.GruposComunalesDetalle.Attach(oSocio)
            oSocio.nIdGrupo = nIdGrupo
            oSocio.cCodCargo = cCodCargo
            db.Entry(cCodCargo).Property("cCodCargo").IsModified = True
            Try
                db.SaveChanges()
                nResp = 1
                sResp = "Registro Correcto"

            Catch ex As Exception
                nResp = 0
                sResp = ex.Message
            End Try
        Else
            oSocio.nIdGrupo = nIdGrupo
            oSocio.cDNI = cDNI
            oSocio.cCodCargo = cCodCargo
            db.GruposComunalesDetalle.Add(oSocio)
            Try
                db.SaveChanges()
                nResp = 1
                sResp = "Registro Correcto"

            Catch ex As Exception
                nResp = 0
                sResp = ex.Message

            End Try
        End If

    End Sub

End Class
