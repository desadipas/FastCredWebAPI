
Imports System.Data
Imports System.Data.SqlClient

Public Class clsConecta

    Private oConex As SqlConnection
    Private oTransac As SqlTransaction 'ARCV 09-05-2011

    Public Function ConexionActual() As SqlConnection
        Return oConex
    End Function

    Public Function AbrirTransaccion() As SqlTransaction
        AbrirTransaccion = oConex.BeginTransaction()

    End Function
    '    Public Function AbrirConexion2() As String
    '        Dim C As SqlCommand

    '        On Error GoTo Erro1
    '        oConex = New SqlConnection("Data Source=SRVLAB001;Initial Catalog=DBSOLIDA;Integrated Security=True;")
    '        'oConex = New SqlConnection("Data Source=SRVLAB001;Initial Catalog=DBSOLIDA;Integrated Security=SSPI;User Instance=true;")

    '        oConex.Open()
    '        C = oConex.CreateCommand()
    '        C.CommandText = "update creditos set nsaldok=1.5 where ncodage=14 and ncodcred=3368"
    '        C.ExecuteNonQuery()
    '        C = Nothing
    '        AbrirConexion2 = "OK"
    '        Exit Function

    'Erro1:
    '        AbrirConexion2 = Err.Description

    '    End Function

    Public Sub AbrirConexion()
        Dim C As SqlCommand


        'oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBCore;Data Source=PROYECTO01\SQLEXPRESS")
        'oConex = New SqlConnection("Initial Catalog=DBCORE;Data Source=SRVLAB01;User=sa;password=1234")
        'oConex = New SqlConnection("Initial Catalog=DBCORE;Data Source=PROYECTO01\SQLEXPRESS;User=sa;password=1234")
        'oConex = New SqlConnection("Initial Catalog=DBCore;Data Source=10.10.1.44;User=sa;password=1234")

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing
        Exit Sub



    End Sub

    Public Function DameConexion() As SqlConnection
        Dim C As SqlCommand

        'Cierre de Mes
        'oConex = New SqlConnection("Initial Catalog=DBSOLIDACierreMes;Data Source=10.100.3.15;User=sa;password=1234")
        'oConex = New SqlConnection("Initial Catalog=DBCierreMes;Data Source=10.100.3.15;User=sa;password=1234")

        'Produccion Edpyme Integrado
        'NO TOCAR O DESPUES DE PROBAR DEJAR IGUAL
        'oConex = New SqlConnection("Data Source=SRVDBSOLIDA;Initial Catalog=DBSOLIDA;Integrated Security=True")

        'Produccion Edpyme SA
        'oConex = New SqlConnection("Initial Catalog=DBSOLIDA;Data Source=10.100.8.75;User=sa;password=1234")

        'Pruebas Fideicomiso
        'oConex = New SqlConnection("Initial Catalog=DBCORE;Data Source=10.100.2.104;User=sa;password=1234")

        'Pruebas  Edpyme
        'oConex = New SqlConnection("Initial Catalog=DBSOLIDA;Data Source=10.100.8.74;User=dbaccess;password=1234")
        '*** DEMO
        'oConex = New SqlConnection("Initial Catalog=DBSOLIDA;Data Source=168.62.168.175;User=sa;password=mur@ll@2018")
        '***COOPAC
        oConex = New SqlConnection("Initial Catalog=DBSOLIDA;Data Source=190.117.160.99;User=dbaccess;password=MF4*2021")

        'Pruebas Edpyme Integrado
        'oConex = New SqlConnection("Data Source=SRVLAB001;Initial Catalog=DBSOLIDA;Integrated Security=True")

        oConex.Open()
        C = oConex.CreateCommand()
        C.CommandText = "SET DATEFORMAT DMY"
        C.ExecuteNonQuery()
        C = Nothing

        Return oConex
    End Function

    Public Function DameConexionConsolidada() As SqlConnection
        oConex = New SqlConnection("Initial Catalog=DBConsolidada;Data Source=10.10.1.185;User=sa;password=1234")
        oConex.Open()
        Return oConex
    End Function

    Public Function DameConexionFinanc() As SqlConnection
        'REAL
        'oConex = New SqlConnection("Initial Catalog=DBFinanc;Data Source=10.10.1.44;User=sa;password=1234")

        'PRUEBAS
        oConex = New SqlConnection("Integrated Security=SSPI;Initial Catalog=DBFinanc;Data Source=PROYECTO01\SQLEXPRESS")

        oConex.Open()
        Return oConex
    End Function

    Public Sub CierraConexion()
        oConex.Close()
        oConex = Nothing

    End Sub

    'ARCV 09-05-2011
    'Public Sub EjecutaTransaction(ByRef connection As SqlConnection, _
    '                                ByRef command As SqlClient.SqlCommand, _
    '                                Optional ByRef pTransaction As SqlTransaction = Nothing, _
    '                                Optional ByVal bIniciaTransac As Boolean = True, _
    '                                Optional ByVal bTerminaTransac As Boolean = True)
    '    Using connection
    '        'connection.Open()

    '        'Dim command As SqlCommand = connection.CreateCommand()
    '        Dim transaction As SqlTransaction

    '        ' Start a local transaction
    '        If bIniciaTransac Then
    '            If pTransaction Is Nothing Then
    '                transaction = connection.BeginTransaction("SampleTransaction")
    '            Else
    '                pTransaction = connection.BeginTransaction("SampleTransaction")
    '            End If
    '        End If

    '        ' Must assign both transaction object and connection
    '        ' to Command object for a pending local transaction.
    '        command.Connection = connection
    '        'command.Transaction = transaction
    '        command.Transaction = IIf(pTransaction Is Nothing, transaction, pTransaction)

    '        Try
    '            'command.CommandText = sComando
    '            command.ExecuteNonQuery()

    '            ' Attempt to commit the transaction.
    '            If bTerminaTransac Then
    '                If pTransaction Is Nothing Then
    '                    transaction.Commit()
    '                    transaction = Nothing
    '                Else
    '                    pTransaction.Commit()
    '                    pTransaction = Nothing
    '                End If
    '            End If

    '        Catch ex As Exception
    '            Console.WriteLine("Commit Exception Type: {0}", ex.GetType())
    '            Console.WriteLine("  Message: {0}", ex.Message)

    '            ' Attempt to roll back the transaction.
    '            Try
    '                transaction.Rollback()

    '            Catch ex2 As Exception
    '                ' This catch block will handle any errors that may have occurred
    '                ' on the server that would cause the rollback to fail, such as
    '                ' a closed connection.
    '                Console.WriteLine("Rollback Exception Type: {0}", ex2.GetType())
    '                Console.WriteLine("  Message: {0}", ex2.Message)
    '            End Try
    '        End Try
    '    End Using
    'End Sub

    Sub EjecutaTransaccion(ByVal oCommand As SqlCommand)

        Try
            oCommand.ExecuteNonQuery()
        Catch ex As Exception
            RevierteTransaccion()
        End Try
    End Sub

    Sub IniciaTransaccion()
        oTransac = oConex.BeginTransaction("SampleTransaction")
    End Sub

    Sub FinalizaTransaccion()
        oTransac.Commit()
        oTransac = Nothing
        oConex.Close()
    End Sub

    Sub RevierteTransaccion()
        oTransac.Rollback()
        oTransac = Nothing
        oConex.Close()
    End Sub

    Public Function TransaccionActual() As SqlTransaction
        Return oTransac
    End Function

End Class
