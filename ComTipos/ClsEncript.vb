Imports System
Imports System.IO
Imports System.Security
Imports System.Security.Cryptography
Imports System.Text

Public Class ClsEncript
    Private Const sSecretKey As String = "Password"


    Private Sub EncryptFile(ByVal sInputFilename As String, _
                   ByVal sOutputFilename As String, _
                   ByVal sKey As String)


        Dim fsInput As New FileStream(sInputFilename, _
                                    FileMode.Open, FileAccess.Read)
        Dim fsEncrypted As New FileStream(sOutputFilename, _
                                    FileMode.Create, FileAccess.Write)
        Dim DES As New DESCryptoServiceProvider()
        'Generacion de Claves y Vector Direccional
        'Si no se asigna una clave el sistema generara una clave aleatoria pero 
        'no se podrá desencriptar.
        DES.Key = ASCIIEncoding.ASCII.GetBytes(sKey)
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)
        Dim desdecrypt As ICryptoTransform = DES.CreateEncryptor()

        Dim cryptostream As New CryptoStream(fsEncrypted, desdecrypt, CryptoStreamMode.Write)

        Dim bytearrayinput(fsInput.Length - 1) As Byte
        fsInput.Read(bytearrayinput, 0, bytearrayinput.Length)
        cryptostream.Write(bytearrayinput, 0, bytearrayinput.Length)

        cryptostream.Flush()
        cryptostream.Close()
        fsInput.Flush()
        fsInput.Close()

    End Sub
    'gsEntorno
    Public Function DevuelveIPServicio(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine

        objStreamReader.Close()

        DevuelveIPServicio = strLine

    End Function

    Public Function DevuelveEntorno(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine

        objStreamReader.Close()

        DevuelveEntorno = strLine

    End Function

    Public Function DevuelveDobleSistema(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine

        objStreamReader.Close()

        DevuelveDobleSistema = strLine

    End Function

    Public Function DevuelveTipoAuntenticacion(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine

        objStreamReader.Close()

        DevuelveTipoAuntenticacion = strLine

    End Function
    Public Function DevuelveMultipleEjecucion(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine

        objStreamReader.Close()

        DevuelveMultipleEjecucion = strLine

    End Function

    Public Function DevuelveIPReportes(ByVal psRuta As String) As String
        Dim objStreamReader As StreamReader
        Dim strLine As String

        objStreamReader = New StreamReader(psRuta)
        strLine = objStreamReader.ReadLine
        strLine = objStreamReader.ReadLine
        objStreamReader.Close()

        DevuelveIPReportes = strLine

    End Function

    Public Sub DecryptArchivo(ByVal sInputFilename As String, _
    ByVal sOutputFilename As String, _
    ByVal sKey As String)

        Dim DES As New DESCryptoServiceProvider()
        'Se requiere una clave de 64 bits y IV para este proveedor.
        'Establecer la clave secreta para el algoritmo DES.
        DES.Key() = ASCIIEncoding.ASCII.GetBytes(sKey)
        'Establecer el vector de inicialización.
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

        'crear la secuencia de archivos para volver a leer el archivo cifrado
        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
        'crear descriptor DES a partir de nuestra instancia de DES
        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
        'crear conjunto de secuencias de cifrado para leer y realizar 
        'una transformación de descifrado DES en los bytes entrantes
        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
        'imprimir el contenido de archivo descifrado

        Dim fsDecrypted As New StreamWriter(sOutputFilename)

        fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)
        fsDecrypted.Flush()
        fsDecrypted.Close()
    End Sub

    Public Sub DecryptFile(ByVal sInputFilename As String, _
    ByVal sOutputFilename As String, _
    ByVal sKey As String)

        Dim DES As New DESCryptoServiceProvider()
        'Se requiere una clave de 64 bits y IV para este proveedor.
        'Establecer la clave secreta para el algoritmo DES.
        DES.Key() = ASCIIEncoding.ASCII.GetBytes(sKey)
        'Establecer el vector de inicialización.
        DES.IV = ASCIIEncoding.ASCII.GetBytes(sKey)

        'crear la secuencia de archivos para volver a leer el archivo cifrado
        Dim fsread As New FileStream(sInputFilename, FileMode.Open, FileAccess.Read)
        'crear descriptor DES a partir de nuestra instancia de DES
        Dim desdecrypt As ICryptoTransform = DES.CreateDecryptor()
        'crear conjunto de secuencias de cifrado para leer y realizar 
        'una transformación de descifrado DES en los bytes entrantes
        Dim cryptostreamDecr As New CryptoStream(fsread, desdecrypt, CryptoStreamMode.Read)
        'imprimir el contenido de archivo descifrado
        Dim fsDecrypted As New StreamWriter(sOutputFilename)
        fsDecrypted.Write(New StreamReader(cryptostreamDecr).ReadToEnd)
        fsDecrypted.Flush()
        fsDecrypted.Close()
    End Sub


    Public Sub EncriptaArchivo(ByVal psRutaArchivo As String, ByVal psRutaArchivoEncriptado As String)
        EncryptFile(psRutaArchivo, psRutaArchivoEncriptado, sSecretKey)

    End Sub

    Public Sub DesencriptaArchivo(ByVal psRutaArchivoEncriptado As String, ByVal psRutaArchivoDesencriptado As String)
        DecryptFile(psRutaArchivoEncriptado, psRutaArchivoDesencriptado, sSecretKey)

    End Sub


End Class
