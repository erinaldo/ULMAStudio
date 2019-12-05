Imports crip = crip2aCAD.clsCR
Imports UCClientWebService.Models
Imports UCClientWebService.Services


Partial Public Class clsBase
    'Public Shared ReadOnly ValidIds As String() = {"WXdgcIZY", "k6Szbr16"}
    Public Shared ReadOnly keyfile As String = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), "key.dat")
    Public Shared RespID As UCClientWebService.Models.ResponseID
    '
    ' Comprobar a través de la Web id ID (Siempre debe existir key.dat)
    Public Shared Sub ID_Comprueba_OnLine()
        RespID = New ResponseID With {.id = "", .valid = False, .message = "", .RemainingDays = -1, .messagelog = ""}
        ' Si hay conexión e internet
        Dim networkinternet As String = EstadoRed_String()
        RespID.messagelog = networkinternet
        Dim srvId As New UCClientWebService.Services.AddInService
        Dim rId As New RequestID With {.id = RespID.id}
        Dim trampeado As Boolean = False
        '
        ' ***** 1.- Comprobar estructura y datos del fichero (Siempre existe key.dat)
        If cr Is Nothing Then cr = New crip("aiiao2K19")
        ' Fichero si existe
        Dim texto As String = " "
        Try
            texto = crip.Texto_Desencriptar(IO.File.ReadAllText(keyfile))
        Catch ex As Exception
        End Try
        Dim lineas As String() = texto.Split(vbCrLf)
        Dim id As String = ""
        Dim pc As String = ""
        Dim tTicksConn As String = ""
        Dim tTicksAddin As String = ""
        '
        If lineas.Count <> 7 Then
            ' No tiene las 7 lineas que debe tener.
            ' (La penultima fila es la última comprobación online en Ticks)
            ' (La ultima fila es la fecha de la ultima ejecución del AddIn en Ticks)
            RespID.id = ""
            RespID.valid = False
            RespID.message = "Registration ID is not valid"
            RespID.messagelog = "FAKE. Invalid key data in file (lineas)"
            trampeado = True
        ElseIf lineas.Count = 7 Then
            id = lineas(2).Replace(vbLf, "")
            pc = lineas(4).Replace(vbLf, "")
            tTicksConn = lineas(5).Replace(vbLf, "")
            Dim nTicksConn As Long = 0
            tTicksAddin = lineas(6).Replace(vbLf, "")
            Dim nTicksAddin As Long = 0
            Dim nDays As Integer = -1        ' Dias desde la última comprobación (>= 0 and < 90 para que sea correcto)
            Dim oDate As Date = Nothing
            RespID.id = id
            ' Comprobar linea 6 (Ticks con la fecha de la última comprobación)
            If Long.TryParse(tTicksConn, nTicksConn) = True Then
                oDate = New Date(nTicksConn)
                nDays = Date.Now.Subtract(oDate).Days
                RespID.RemainingDays = nDays
            Else
                ' tTicks no escorrecto (No es un número)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "FAKE. Invalid key data in file (Last connection for validation)"
                trampeado = True
                GoTo FINAL
                Exit Sub
            End If
            ' Comprobar linea 7 (Ticks con la fecha de la última ejecución del AddIn)
            If Long.TryParse(tTicksAddin, nTicksAddin) = False Then
                ' nTicksAddin no escorrecto (No es un número)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "FAKE. Invalid key data in file (Last execution Addin)"
                trampeado = True
                GoTo FINAL
                Exit Sub
            End If
            ' Comprobar ID y PC
            If id = "" Then
                ' ID no escorrecto
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "FAKE. Invalid key data in file (id)"
                trampeado = True
            ElseIf pc <> _appComputer Then
                'PC no es el correcto
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "FAKE. Invalid pc validation (pc)"
                trampeado = True
            ElseIf nDays < 0 OrElse nDays > 90 Then
                'Han pasado los 90 días sin conexión
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Connection needed to verify registration ID (offline period expired)"
                RespID.messagelog = "Network connection off. Expired (" & nDays & ")"
            ElseIf Date.Now.Ticks < nTicksAddin Then
                'Han cambiado lo fecha del equipo
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "FAKE. Invalid date Last Addin (" & Date.Now.ToString(formatofecha) & ")"
                trampeado = True
            ElseIf id <> "" AndAlso pc = _appComputer Then
                rId.id = id
                If networkinternet = "" Then
                    ' SI hay conexión / internet
                    ' Comprobar ID via Web y recoger resultado
                    RespID = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)
                    RespID.messagelog = RespID.message
                Else
                    ' NO hay conexión / internet
                    ' Validación temporal (máximo 90 días)
                    RespID.valid = True
                    RespID.message = networkinternet
                    RespID.messagelog = networkinternet & " (id=" & id & "|pc=" & pc & "|RemainingDays= " & RespID.RemainingDays.ToString & ")"
                End If
            End If
        End If
FINAL:
        ' Si validación incorrecta borrar key.data
        If (RespID.valid = False AndAlso networkinternet = "") OrElse trampeado = True Then
            IO.File.Delete(keyfile)
        Else
            ' Validación temporal (menos de 90 días sin conexión) Escribir key.dat manteniendo fecha ultima conexión.
            keyfile_escribe(id, tTicksConn)
        End If
    End Sub

    Public Shared Function EstadoRed_Boolean() As Boolean
        Try
            If My.Computer.Network.IsAvailable() Then
                If My.Computer.Network.Ping("www.google.com", 1000) Then 'Asignamos la pagina a consultar ejemplo www.google.com y el tiempo de espera máximo
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Shared Function EstadoRed_String() As String
        Dim resultado As String = ""
        Try
            If My.Computer.Network.IsAvailable() Then
                If My.Computer.Network.Ping("www.google.com", 1000) Then 'Asignamos la pagina a consultar ejemplo www.google.com y el tiempo de espera máximo
                    resultado = ""
                Else
                    resultado = "Offline. No internet access available"
                End If
            Else
                resultado = "Offline. No network connection available"
            End If
        Catch ex As Exception
            resultado = "Offline. No network connection available"
        End Try
        Return resultado
    End Function
    '
    Public Shared Sub keyfile_escribe(queID As String, Optional connection As String = "")
        ' WXdgcIZY k6Szbr16
        Try
            Dim newGuid As String = Guid.NewGuid.ToString("N")
            Dim lineas As String = ""
            lineas &= newGuid.Substring(0, 8) & vbCrLf
            lineas &= newGuid.Substring(10, 8) & vbCrLf
            lineas &= queID & vbCrLf
            lineas &= newGuid.Substring(20, 8) & vbCrLf
            lineas &= _appComputer & vbCrLf
            lineas &= IIf(connection = "", Date.Now.Ticks, connection) & vbCrLf    ' Ticks de la última conexion para comprobar
            lineas &= Date.Now.Ticks                    ' Ticks de la ultima ejecución del Addin
            '
            IO.File.WriteAllText(keyfile, crip.Texto_Encripta(lineas))
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    Public Shared Function getIdRegistrado() As String
        Dim id As String = ""
        If System.IO.File.Exists(keyfile) Then
            ' ***** 1.- Comprobar estructura y datos del fichero (Siempre existe key.dat)
            If cr Is Nothing Then cr = New crip("aiiao2K19")
            ' Fichero si existe
            Dim texto As String = " "
            Try
                texto = crip.Texto_Desencriptar(IO.File.ReadAllText(keyfile))
            Catch ex As Exception
            End Try
            Dim lineas As String() = texto.Split(vbCrLf)
            If lineas.Count = 7 Then
                id = lineas(2).Replace(vbLf, "")
            End If
        End If
        clsBase.idRegistrado = id
        Return id
    End Function

End Class

Public Enum validacion
    Code_and_PC_correct               ' id y pc correctos. Fichero existe y tiene las 5 lineas
    File_not_exist            ' fichero no existe
    File_incorrect_data    ' Fichero existe. Pero no tiene las 5 lineas que debe llevar (fake)
    Invalid_code               ' id no es correcto. Fichero existe y tiene 5 lineas
    PC_name_incorrect               ' pc no es correcto. Fichero existe y tiene 5 lineas
End Enum
