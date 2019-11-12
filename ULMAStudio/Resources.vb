Imports System.IO
Imports System.Reflection

Namespace MyResource
    Public Class Resources
        Public Shared Function Texto_DameIncrustado(nombre As String, Optional ass As System.Reflection.Assembly = Nothing) As String
            Dim resultado As String = ""
            If ass Is Nothing Then ass = System.Reflection.Assembly.GetExecutingAssembly
            ' Para leer un fichero externo:
            ' en vez de nombre (new Uri(Path.Combine( _imageFolder, imageName ))
            '
            Dim nAss As String = ass.GetName.Name.ToString
            Dim nResource As String = IIf(nombre.StartsWith(nAss & "."), nombre, nAss & "." & nombre).ToString
            Dim s As IO.Stream = ass.GetManifestResourceStream(nResource)     ' En carpeta Resources (Como "Recurso incrustado")
            Try
                If s IsNot Nothing Then
                    Dim sReader As New StreamReader(s)
                    resultado = sReader.ReadToEnd()
                    sReader.Close()
                    s.Close()
                End If
            Catch ex As Exception
                ' deberíamos cargarla desde el disco duro.
                If IO.File.Exists(nombre) Then
                    resultado = IO.File.ReadAllText(nombre, Text.Encoding.UTF8)
                End If
            End Try
            Return resultado
        End Function

        Public Shared Function Imagen_DameIncrustada(nombre As String, Optional ass As System.Reflection.Assembly = Nothing) As System.Windows.Media.Imaging.BitmapImage
            Dim resultado As New System.Windows.Media.Imaging.BitmapImage
            If ass Is Nothing Then ass = System.Reflection.Assembly.GetExecutingAssembly
            ' Para leer un fichero externo:
            ' en vez de nombre (new Uri(Path.Combine( _imageFolder, imageName ))
            '
            Dim nAss As String = ass.GetName.Name.ToString
            Dim nResource As String = IIf(nombre.StartsWith(nAss & "."), nombre, nAss & "." & nombre).ToString
            Dim s As IO.Stream = ass.GetManifestResourceStream(nResource)     ' En carpeta Resources (Como "Recurso incrustado")
            Try
                If s IsNot Nothing Then
                    resultado.BeginInit()
                    resultado.StreamSource = s
                    resultado.CacheOption = Windows.Media.Imaging.BitmapCacheOption.OnLoad
                    resultado.EndInit()
                    s.Close()
                End If
            Catch ex As Exception
                ' deberíamos cargarla desde el disco duro.
                If IO.File.Exists(nombre) Then
                    resultado.BeginInit()
                    resultado.UriSource = New Uri(IO.Path.Combine(ULMALGFree.clsBase._imgFolder, nombre))
                    resultado.CacheOption = Windows.Media.Imaging.BitmapCacheOption.OnLoad
                    resultado.EndInit()
                End If
            End Try
            Return resultado
        End Function

        Public Shared Function Icono_DameIncrustado(nombre As String, Optional ass As System.Reflection.Assembly = Nothing) As System.Drawing.Icon
            Dim resultado As System.Drawing.Icon = Nothing
            If ass Is Nothing Then ass = System.Reflection.Assembly.GetExecutingAssembly
            ' Para leer un fichero externo:
            ' en vez de nombre (new Uri(Path.Combine( _imageFolder, imageName ))
            '
            If nombre.ToLower.EndsWith(".ico") = False Then nombre &= ".ico"
            Dim nAss As String = ass.GetName.Name.ToString
            Dim nResource As String = IIf(nombre.StartsWith(nAss & "."), nombre, nAss & "." & nombre).ToString
            Dim s As IO.Stream = ass.GetManifestResourceStream(nResource)     ' En carpeta Resources (Como "Recurso incrustado")
            Try
                If s IsNot Nothing Then
                    resultado = New System.Drawing.Icon(s)
                    s.Close()
                End If
            Catch ex As Exception
                ' deberíamos cargarla desde el disco duro.
                If IO.File.Exists(nombre) Then
                    resultado = New System.Drawing.Icon(nombre)
                End If
            End Try
            Return resultado
        End Function
        Public Function Stream_DameIncrustado(nombre As String, Optional ass As System.Reflection.Assembly = Nothing) As Byte()
            Dim resultado As Byte() = Nothing
            If ass Is Nothing Then ass = System.Reflection.Assembly.GetExecutingAssembly
            ' Para leer un fichero externo:
            ' en vez de nombre (new Uri(Path.Combine( _imageFolder, imageName ))
            '
            Dim nAss As String = ass.GetName.Name.ToString
            Dim nResource As String = IIf(nombre.StartsWith(nAss & "."), nombre, nAss & "." & nombre).ToString
            Dim s As IO.Stream = ass.GetManifestResourceStream(nResource)     ' En carpeta Resources (Como "Recurso incrustado")
            Try
                If s IsNot Nothing Then
                    ReDim resultado(CInt(s.Length))
                    s.Read(resultado, 0, CInt(s.Length))
                    s.Close()
                End If
            Catch ex As Exception
                resultado = IO.File.ReadAllBytes(nombre)
            End Try
            Return resultado
        End Function
    End Class
End Namespace
