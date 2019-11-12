Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Module UtilesIP
    '
    ' Devuelve la IP privada (Intranet, por defecto). La que tiene este equipo en la Red interna de la empresa
    ' O todas, si solointranet = False
    Public Function IPPrivada_DameLista(Optional solointranet As Boolean = True) As String
        Dim valorIp As String = ""
        Dim listAddres = Dns.GetHostEntry(My.Computer.Name).AddressList
        If solointranet Then
            valorIp = listAddres.FirstOrDefault(
                Function(i) i.AddressFamily = Sockets.AddressFamily.InterNetwork).ToString()
        Else
            ' ** Solo InterNetwork
            'Dim varios = From i In listAddres
            '             Where i.AddressFamily = Sockets.AddressFamily.InterNetwork
            '             Select i

            'Dim ips As New List(Of String)
            'For Each ip As IPAddress In varios
            '    If ip.ToString.StartsWith("192") Then ips.Add(ip.ToString)
            'Next
            ' ** Todas
            Dim ips As New List(Of String)
            For Each ip As IPAddress In listAddres
                If ip.ToString.StartsWith("192") Then ips.Add(ip.ToString)
            Next

            If ips.Count > 0 Then valorIp = String.Join("|", ips.ToArray)
        End If
        Return valorIp
    End Function

    Public Function IPPrivada_DameCorto() As String
        Dim ip As System.Net.IPHostEntry
        ip = Dns.GetHostEntry(My.Computer.Name)
        Return ip.AddressList.Where(Function(i) i.ToString.StartsWith("192"))(0).ToString
    End Function
    Function IPPublica_Dame() As String  ' IPAddress
        Dim resultado As String = ""
        Dim lol As WebClient = New WebClient()
        Try
            Dim str As String = lol.DownloadString("http://checkip.dyndns.org/")    ' El html con el resultado
            resultado = str.Split(":"c)(1).Split("<"c)(0).Trim
        Catch ex As Exception
            ' Error de conexión, no hay internet o la página no ha dado la IP
            resultado = "¿?"
        End Try

        Return resultado
    End Function
End Module
