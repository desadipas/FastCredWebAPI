Imports System
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Public Class ClsPrejudicial : Inherits MarshalByRefObject
    Public Function RecuperaControlesPrejudicial(ByVal pnCodAge As Integer) As ComTipos.ClsPrejudicial
        Dim C As ComTipos.ClsPrejudicial
        C = New ComTipos.ClsPrejudicial
        C.GestorCobranza.CargaTabla(, , "Prejudicial_RecuperaGestorxAgencia " & pnCodAge)
        RecuperaControlesPrejudicial = C
        C = Nothing
    End Function

    Public Function RecuperaCalendarioActualPrejudicial(ByVal pnCodAge As Integer, _
    ByVal pnCodCred As Integer) As ComTipos.ClsPrejudicial
        Dim C As ComTipos.ClsPrejudicial
        C = New ComTipos.ClsPrejudicial
        C.CalendariActual.CargaTabla(, , "CredW_RecuperaCalendario " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioActualPrejudicial = C
        C = Nothing
    End Function
    Public Function RecuperaCalendarioNegociaciones(ByVal pnCodAge As Integer, _
    ByVal pnCodCred As Integer) As ComTipos.ClsPrejudicial

        Dim C As ComTipos.ClsPrejudicial
        C = New ComTipos.ClsPrejudicial
        C.CalendariNegociaciones.CargaTabla(, , "Prejudicial_RecuperaCalendarioNegociaciones " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioNegociaciones = C
        C = Nothing
    End Function
    Public Function FiltrarClientexGestorCobranzas(ByVal pnCodAge As Integer, ByVal pnCodgestor As Integer) As ComTipos.ClsPrejudicial

        Dim C As ComTipos.ClsPrejudicial
        C = New ComTipos.ClsPrejudicial
        C.ClientesxGestorCobranza.CargaTabla(, , "Prejudicial_FiltrarClientexGestor " & pnCodAge & "," & pnCodgestor)
        FiltrarClientexGestorCobranzas = C
        C = Nothing
    End Function

    Public Function RecuperaCalendarioNegociaciones_FIDEICOMISO(ByVal pnCodAge As Integer, _
ByVal pnCodCred As Integer) As ComTipos.ClsPrejudicial

        Dim C As ComTipos.ClsPrejudicial
        C = New ComTipos.ClsPrejudicial
        C.CalendariNegociaciones.CargaTabla(, , "DBCORE..Prejudicial_RecuperaCalendarioNegociaciones " & pnCodAge & "," & pnCodCred)
        RecuperaCalendarioNegociaciones_FIDEICOMISO = C
        C = Nothing
    End Function
End Class
