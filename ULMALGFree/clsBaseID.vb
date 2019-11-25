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
        '
        ' ***** 1.- Comprobar estructura y datos del fichero (Siempre existe key.dat)
        cr = New crip("aiiao2K19")
        ' Fichero si existe
        Dim texto As String = " "
        Try
            texto = crip.Texto_Desencriptar(IO.File.ReadAllText(keyfile))
        Catch ex As Exception
        End Try
        Dim lineas As String() = texto.Split(vbCrLf)
        If lineas.Count <> 6 Then
            ' No tiene las 6 lineas que debe tener. (La ultima fila es la última comprobación online en Ticks)
            RespID.id = ""
            RespID.valid = False
            RespID.message = "Registration ID is not valid"
            RespID.messagelog = "Invalid key data in file (lineas)"
        ElseIf lineas.Count = 6 Then
            Dim id As String = lineas(2).Replace(vbLf, "")
            Dim pc As String = lineas(4).Replace(vbLf, "")
            Dim tTicks As String = lineas(5).Replace(vbLf, "")
            Dim nTicks As Long = 0
            Dim nDays As Integer = -1        ' Dias desde la última comprobación (>= 0 and < 90 para que sea correcto)
            Dim oDate As Date = Nothing
            RespID.id = id
            ' Comprobar linea 6 (Ticks con la fecha de la última comprobación)
            If Long.TryParse(tTicks, nTicks) = True Then
                oDate = New Date(nTicks)
                nDays = Date.Now.Subtract(oDate).Days
                RespID.RemainingDays = nDays
            Else
                ' tTicks no escorrecto (No es un número)
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "Invalid key data in file (date)"
                Exit Sub
            End If
            ' Comprobar ID y PC
            If id = "" Then
                ' ID no escorrecto
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "Invalid key data in file (id)"
            ElseIf pc <> _appComputer Then
                'PC no es el correcto
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "Registration ID is not valid"
                RespID.messagelog = "Invalid pc validation (pc)"
            ElseIf nDays < 0 OrElse nDays > 90 Then
                'Han pasado los 90 días sin conexión
                'IO.File.Delete(keyfile)
                RespID.valid = False
                RespID.message = "ULMA Studio can not be launched without network connection"
                RespID.messagelog = "Network connection off (" & nDays & ")"
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
                    RespID.messagelog = networkinternet & " (id=" & id & "|pc=" & pc & "|ReminingDays= " & RespID.RemainingDays.ToString & ")"
                End If
            End If
        End If
        ' Si validación incorrecta borrar key.data
        If RespID.valid = False Then
            IO.File.Delete(keyfile)
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
    Public Shared Sub keyfile_escribe(queID As String)
        ' WXdgcIZY k6Szbr16
        Dim newGuid As String = Guid.NewGuid.ToString("N")
        Dim lineas As String = ""
        lineas &= newGuid.Substring(0, 8) & vbCrLf
        lineas &= newGuid.Substring(10, 8) & vbCrLf
        lineas &= queID & vbCrLf
        lineas &= newGuid.Substring(20, 8) & vbCrLf
        lineas &= _appComputer & vbCrLf
        lineas &= Date.Now.Ticks
        '
        IO.File.WriteAllText(keyfile, crip.Texto_Encripta(lineas))
    End Sub
End Class

Public Enum validacion
    Code_and_PC_correct               ' id y pc correctos. Fichero existe y tiene las 5 lineas
    File_not_exist            ' fichero no existe
    File_incorrect_data    ' Fichero existe. Pero no tiene las 5 lineas que debe llevar (fake)
    Invalid_code               ' id no es correcto. Fichero existe y tiene 5 lineas
    PC_name_incorrect               ' pc no es correcto. Fichero existe y tiene 5 lineas
End Enum
