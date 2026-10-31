
Imports System.Runtime.Serialization
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary

Public Class ClsSerializa

    Public Sub Serializar(ByVal pClase As Object, ByVal psNomFile As String)
        Dim formatter As IFormatter = New BinaryFormatter 'formateador binario para realizar la serialización.
        ' Todo lo que necesita es crear una instancia de la secuencia y el formateador que desee utilizar, 
        Dim stream As Stream = New FileStream("C:\Tempo\" & psNomFile, FileMode.Create, FileAccess.Write, FileShare.Read)
        'y a continuación, llamar al método Serialize en el formateador
        formatter.Serialize(stream, pClase)
        stream.Close()
    End Sub

    Public Function DesSerializar(ByVal psNomFile As String) As Object

        Dim formatter As IFormatter = New BinaryFormatter
        Dim stream As Stream

        If Not File.Exists("C:\Tempo\" & psNomFile) Then
            DesSerializar = Nothing
            Exit Function
        End If

        stream = New FileStream("C:\Tempo\" & psNomFile, FileMode.Open, FileAccess.Read, FileShare.Read)
        'Dim stream As Stream = New FileStream(psNomFile, FileMode.Open, FileAccess.Read, FileShare.Read)
        'If stream Is Nothing Then
        ' DesSerializar = Nothing
        'Exit Function
        'End If
        DesSerializar = formatter.Deserialize(stream)
        stream.Close()
    End Function

    Public Sub DestruirArchivo(ByVal psNomFile As String)
        If Not File.Exists("C:\Tempo\" & psNomFile) Then
            Exit Sub
        End If
        Kill("C:\Tempo\" & psNomFile)
    End Sub

End Class
