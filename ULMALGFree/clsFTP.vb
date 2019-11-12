Imports System.Net.FtpWebRequest
Imports System.Net
Imports System.IO
Imports System.Windows.Forms

Friend Class clsFTP
    Public Const sep As String = "||"
    Public host, dir, user, pass, port As String
    Public conexionOk As Boolean = False
    Public ColDatos As Dictionary(Of String, String()) ' key = nombrefichero.ext, value= AddDatos() As String = {"", "", ""}    '(0) fullUri, (1) tamaño en bytes, (2) Fecha modificacion

    Public Sub New(ByVal host As String, dir As String, ByVal user As String, ByVal pass As String, Optional port As String = "21")
        Me.host = host
        Me.dir = dir
        Me.user = user
        Me.pass = pass
        Me.port = port
        ColDatos = New Dictionary(Of String, String())
        If Me.host.Contains(":21") And Me.port <> "21" Then
            Me.host.Replace(":21", ":" & Me.port)
        End If
        Try
            Dim lista As String() = Me.FTP_ListaDir(dir)
            If lista.Length = 1 AndAlso lista(0).ToString.ToUpper.StartsWith("ERROR") Then
                ULMALGFree.clsBase.txtResponse = lista(0).Replace("ERROR", "").Replace(vbCrLf, "").Trim
                conexionOk = False
            Else
                conexionOk = True
            End If
        Catch ex As Exception
            Me.conexionOk = False
        End Try
    End Sub
    '
    Public Function FTP_ExisteDir(queDir As String) As String
        Dim mensaje As String = ""
        Dim ftpReq As FtpWebRequest = Nothing
        Dim ftpresp As FtpWebResponse = Nothing
        Try
            ftpReq = DirectCast(FtpWebRequest.Create(queDir), FtpWebRequest)
            ftpReq.Method = WebRequestMethods.Ftp.ListDirectory
            ftpReq.Credentials = New NetworkCredential(Me.user, Me.pass)
            ftpReq.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            ftpReq.Proxy = Nothing
            '
            ftpresp = DirectCast(ftpReq.GetResponse, FtpWebResponse)
            mensaje = ftpresp.StatusDescription
            conexionOk = True
        Catch ex As Exception
            mensaje = "ERROR " & ex.ToString
            conexionOk = False
        Finally
            ftpReq = Nothing
            If ftpresp IsNot Nothing Then
                ftpresp.Close()
            End If
            ftpresp = Nothing
        End Try

        Return mensaje
    End Function
    Public Function eliminarFichero(ByVal fichero As String) As String
        Dim peticionFTP As FtpWebRequest

        ' Creamos una petición FTP con la dirección del fichero a eliminar
        peticionFTP = CType(WebRequest.Create(New Uri(fichero)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Eliminar un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.DeleteFile
        peticionFTP.UsePassive = False

        Try
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuestaFTP.Close()
            ' Si todo ha ido bien, devolvemos String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return "ERROR " & ex.Message
        End Try
    End Function

    Public Function existeObjeto(ByVal dir As String) As Boolean
        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

        peticionFTP.UsePassive = False

        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            Return True
        Catch ex As Exception
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

    Public Function creaDirectorio(ByVal dir As String) As String
        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return "ERROR " & ex.Message
        End Try
    End Function

    Public Function FTP_Upload(ByVal ficheroLocal As String, ByVal ficheroFtp As String) As String
        Dim resultado As String = ""
        '
        'Upload File to FTP site
        'Create Request To Upload File'
        Dim wrUpload As FtpWebRequest = DirectCast(WebRequest.Create _
           (ficheroFtp), FtpWebRequest)
        'Specify Username & Password'
        wrUpload.Credentials = New NetworkCredential(Me.user, Me.pass)
        'Start Upload Process'
        wrUpload.Method = WebRequestMethods.Ftp.UploadFile

        'Locate File And Store It In Byte Array'
        Dim btfile() As Byte = IO.File.ReadAllBytes(ficheroLocal)
        Dim strFile As IO.Stream = Nothing
        Try
            'Get File'
            strFile = wrUpload.GetRequestStream()
            'Upload Each Byte'
            strFile.Write(btfile, 0, btfile.Length)
            resultado = "CORRECTO: " & ficheroFtp & " --> Subido"
            'resultado = wrUpload.ToString
        Catch ex As Exception
            resultado = "ERROR " & ex.ToString
        Finally
            If strFile IsNot Nothing Then
                'Close'
                strFile.Close()
            End If
        End Try
        'Free Memory'
        strFile = Nothing

        Return resultado
    End Function


    'Delete File On FTP Server'
    Public Function FTP_Borra(fiFtp As String) As String
        Dim resultado As String = ""
        If FTP_DameDatosFichero(fiFtp, queDatoFichero.tamaño).ToUpper.StartsWith("ERROR") Then
            ' No exite el fichero
            Return "ERROR: No existe el fichero " & fiFtp
            Exit Function
        End If
        'Create Request To Delete File'
        Dim wrDelete As FtpWebRequest =
             DirectCast(WebRequest.Create(fiFtp),
             FtpWebRequest)

        wrDelete.Credentials = New NetworkCredential(Me.user, Me.pass)

        'Specify That You Want To Delete A File'
        wrDelete.Method = WebRequestMethods.Ftp.DeleteFile
        'Response Object'
        Dim rDeleteResponse As FtpWebResponse =
             CType(wrDelete.GetResponse(),
             FtpWebResponse)
        'Show Status Of Delete'
        'Console.WriteLine("Delete status: {0}", rDeleteResponse.StatusDescription)
        'resultado = "Delete status: " & rDeleteResponse.StatusDescription
        resultado = "CORRECTO: Borrado " & fiFtp & rDeleteResponse.StatusDescription
        'Close'
        rDeleteResponse.Close()
        Return resultado
    End Function

    Public Sub DescargaFicheroSimple(ByVal oriFTP As String, ByVal desPC As String)
        Try
            'My.Computer.Network.DownloadFile(origen, destino, usuario, clave, True, 100, True)
            Dim pp As New Microsoft.VisualBasic.Devices.Network
            If pp.IsAvailable Then _
                pp.DownloadFile(oriFTP, desPC, Me.user, Me.pass, True, 100000, True)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Function DescargarFicheroFTP(ByVal DirFtp As String, ByVal user As String,
            ByVal pass As String, ByVal dirLocal As String, ByVal Fichero As String,
            ByRef terminado As Boolean, Optional ByRef quePb As ProgressBar = Nothing) As String

        If DirFtp.EndsWith("/") = False Then DirFtp &= "/"
        Dim localFile As String = Fichero
        Dim remoteFile As String = Fichero
        Dim host As String = DirFtp 'nombre de la carpeta en nuestro server FTP donde estan los archivos que deseamos descargar
        ' colocamos el nombre de usuario y password respectivo para acceder al server, si este no poseyera, dejar solo las comillas, osea ""
        Dim username As String = user
        Dim password As String = pass
        Dim URI As String = host & remoteFile ' nombre completo de la ruta del archivo  & remoteFile
        Dim ftp As System.Net.FtpWebRequest = CType(Net.FtpWebRequest.Create(URI), Net.FtpWebRequest)
        ftp.Credentials = New System.Net.NetworkCredential(username, password)
        ftp.KeepAlive = False
        ftp.UseBinary = True
        ftp.Method = System.Net.WebRequestMethods.Ftp.DownloadFile
        Dim tamFile As Long = 32000
        Try
            tamFile = FTP_DameDatosFichero(URI, queDatoFichero.tamaño)
        Catch ex As Exception
            ' Coger tamaño por defecto
        End Try
        Try
            Using response As System.Net.FtpWebResponse = CType(ftp.GetResponse, System.Net.FtpWebResponse)
                Using responseStream As IO.Stream = response.GetResponseStream
                    If quePb IsNot Nothing Then
                        quePb.Value = 0
                        quePb.Maximum = tamFile
                        quePb.Refresh()
                    End If
                    Using fs As New IO.FileStream(dirLocal & "\" & Fichero, IO.FileMode.Create)
                        Dim buffer(2047) As Byte
                        Dim read As Integer = 0
                        Do
                            read = responseStream.Read(buffer, 0, buffer.Length)
                            fs.Write(buffer, 0, read)
                            If quePb IsNot Nothing Then
                                If quePb.Value + read <= quePb.Maximum Then
                                    quePb.Value += read
                                    quePb.Refresh()
                                End If
                            End If
                        Loop Until read = 0
                        If quePb IsNot Nothing Then quePb.Value = quePb.Maximum : quePb.Refresh()
                        responseStream.Close()
                        fs.Flush()
                        fs.Close()
                    End Using
                    responseStream.Close()
                    'responseStream.Dispose()
                End Using
                response.Close()
            End Using
        Catch ex As Exception
            Return ex.Message
            Exit Function
        Finally
            ftp = Nothing
            terminado = True
        End Try
        Return "FIN"
    End Function

    Public Sub CargarFicheroSimple(ByVal oriPC As String, ByVal desFTP As String)
        'Dim oriPC As String = "C:\Temp\CommandNames.txt"
        'Dim desFTP As String = "ftp://80.34.88.43/2aCAD/2aCADAdvancedTools/Inventor2011/CommandNames.txt"
        Try
            'My.Computer.Network.DownloadFile(origen, destino, usuario, clave, True, 100, True)
            Dim pp As New Microsoft.VisualBasic.Devices.Network
            If pp.IsAvailable Then _
                pp.UploadFile(oriPC, desFTP, Me.user, Me.pass, True, 100000, FileIO.UICancelOption.DoNothing)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Function FTP_ListaDir(ByVal queDir As String, Optional fullpath As Boolean = False) As String()
        ColDatos.Clear()
        Dim mensaje(-1) As String
        Dim ftp As FtpWebRequest = Nothing
        Dim ftpresp As FtpWebResponse = Nothing
        Try
            If queDir.ToLower.StartsWith("ftp") Then
                ftp = DirectCast(FtpWebRequest.Create(queDir), FtpWebRequest)
            ElseIf Me.host.EndsWith("/") = False And queDir.StartsWith("/") = False Then
                ftp = DirectCast(FtpWebRequest.Create(Me.host & "/" & queDir), FtpWebRequest)
            Else
                ftp = DirectCast(FtpWebRequest.Create(Me.host & queDir), FtpWebRequest)
            End If
            '/// ************************

            '///NOTE if you need to authenticate with username / password you would add the following 2 lines ...
            Dim cred As New NetworkCredential(Me.user, Me.pass)
            ftp.Credentials = cred
            '/// ************************

            ftp.AuthenticationLevel = Security.AuthenticationLevel.MutualAuthRequested
            ftp.Method = WebRequestMethods.Ftp.ListDirectoryDetails
            ftp.Proxy = Nothing
            ftp.UsePassive = True

            ftpresp = DirectCast(ftp.GetResponse, FtpWebResponse)
            If ftpresp.StatusCode <> FtpStatusCode.OpeningData Then
                ReDim mensaje(0)
                mensaje(0) = "ERROR " & ftpresp.StatusCode.ToString & " = " & ftpresp.StatusDescription.ToString
                Return mensaje
                Exit Function
            End If
            'txtResponse = "StatusCode=" & ftpresp.StatusCode.ToString & " | StatusDescription=" & ftpresp.StatusDescription

            '/// **** these 3 lines are NOT needed, just shows you a bit of the server logon messages you would see ...
            'Console.WriteLine(ftp.Headers.ToString)
            'Console.WriteLine(ftpresp.BannerMessage)
            'Console.WriteLine(ftpresp.Headers.ToString)
            '/// ****

            Dim sreader As New IO.StreamReader(ftpresp.GetResponseStream)

            While Not sreader.Peek = -1
                Dim ftpList As String() = sreader.ReadLine.Split(" "c)
                'Dim ftpBites As String = FormatNumber(CDbl(ftpList(ftpList.GetUpperBound(0) - 4)) / 1000000, 2, TriState.True) & " Mb"
                Dim ftpBites As String = ftpList(ftpList.GetUpperBound(0) - 4)
                Dim ftpFecha As String = ftpList(ftpList.GetUpperBound(0) - 2) & "/" & ftpList(ftpList.GetUpperBound(0) - 3) & " " & ftpList(ftpList.GetUpperBound(0) - 1)
                Dim ftpfile As String = ftpList(ftpList.GetUpperBound(0))
                If ftpfile = "" Then Continue While ' Si está vacío (Es un subdirectorio), continuar
                'Console.WriteLine(ftpfile)
                'If ftpfile.Contains(".bsp") And Not ftpfile.Contains(".ztmp") Then
                'lwMaps.Items.Add(ftpfile, 6)
                ReDim Preserve mensaje(mensaje.GetUpperBound(0) + 1)
                mensaje(mensaje.GetUpperBound(0)) = IIf(fullpath = True,
                                                       IIf(queDir.EndsWith("/") = False,
                                                           queDir & "/" & ftpfile,
                                                           queDir & ftpfile),
                                                        ftpfile) '& vbTab & ftpBites & vbTab & ftpFecha
                'mensaje &= ftpfile & vbTab & ftpBites & vbTab & ftpFecha & vbCrLf
                'End If
                ColDatos.Add(ftpfile, {ftp.RequestUri.OriginalString & "/" & ftpfile, ftpBites, ftpFecha})
            End While
        Catch ex As Exception
            ReDim mensaje(0)
            mensaje(0) = "ERROR " & ex.Message.Replace(":", "|").Replace(";", "|").Replace(vbCrLf, "|")
            'MsgBox(mensaje(0))
            'Throw New System.Exception(ex.ToString)
        Finally
            ftp = Nothing
            If ftpresp IsNot Nothing Then
                ftpresp.Close()
            End If
            ftpresp = Nothing
        End Try

        Return mensaje
        '-		ftpList	{Length=17}	String()
        '(0)	"-rw-r--r--"	String
        '(1)	"1"	String
        '(2)	"ftp"	String
        '(3)	"ftp"	String
        '(4)	""	String
        '(5)	""	String
        '(6)	""	String
        '(7)	""	String
        '(8)	""	String
        '(9)	""	String
        '(10)	""	String
        '(11)	""	String
        '(12)	"234001"	String
        '(13)	"Dec"	String
        '(14)	"05"	String
        '(15)	"16:04"	String
        '(16)	"CommandNames.txt"	String
    End Function

    Public Function existeObjeto1(ByVal dir As String) As Boolean
        Dim peticionFTP As FtpWebRequest

        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

        peticionFTP.UsePassive = False

        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            Return True
        Catch ex As Exception
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

    Public Enum queDatoFichero
        tamaño = 0
        fecha = 1
        '-		ftpList	{Length=17}	String()
        '(0)	"-rw-r--r--"	String
        '(1)	"1"	String
        '(2)	"ftp"	String
        '(3)	"ftp"	String
        '(4)	""	String
        '(5)	""	String
        '(6)	""	String
        '(7)	""	String
        '(8)	""	String
        '(9)	""	String
        '(10)	""	String
        '(11)	""	String
        '(12)	"234001"	String
        '(13)	"Dec"	String
        '(14)	"05"	String
        '(15)	"16:04"	String
        '(16)	"CommandNames.txt"	String
    End Enum

    Public Function FTP_DameDatosFichero(ByVal FileFTP As String, ByVal queDato As queDatoFichero) As String
        Dim resultado As String = ""
        Dim peticionFTP As FtpWebRequest
        Try
            ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
            peticionFTP = CType(WebRequest.Create(New Uri(FileFTP)), FtpWebRequest)

            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)

            ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
            'peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp

            '' PETICIÓN QUE ESTAMOS SOLICITANDO AL SERVIDOR
            If queDato = queDatoFichero.fecha Then
                peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
            ElseIf queDato = queDatoFichero.tamaño Then
                peticionFTP.Method = WebRequestMethods.Ftp.GetFileSize
            End If

            peticionFTP.UsePassive = False

            Try
                ' Si el objeto existe, se devolverá True
                Dim respuestaFTP As FtpWebResponse
                respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
                If queDato = queDatoFichero.fecha Then
                    'Return respuestaFTP.LastModified
                    resultado = respuestaFTP.LastModified.ToLongDateString
                ElseIf queDato = queDatoFichero.tamaño Then
                    'Return respuestaFTP.ContentLength
                    resultado = respuestaFTP.ContentLength.ToString
                End If
            Catch ex As Exception
                resultado = "ERROR: " & ex.ToString
                ' Si el objeto no existe, se producirá un error y al entrar por el Catch
                ' se devolverá falso
            End Try
        Catch ex As Exception

        End Try
        Return resultado
    End Function
End Class
