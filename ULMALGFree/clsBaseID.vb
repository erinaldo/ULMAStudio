Imports crip = crip2aCAD.clsCR
Imports UCClientWebService.Models
Imports UCClientWebService.Services


Partial Public Class clsBase
    'Public Shared ReadOnly ValidIds As String() = {"WXdgcIZY", "k6Szbr16"}
    Public Shared idform As String = ""
    Public Shared ReadOnly keyfile As String = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), "key.dat")
    Public Shared resultado As UCClientWebService.Models.ResponseID
    '
    'Public Shared Function ID_Comprueba_OffLine() As validacion
    '    Dim resultado As validacion = validacion.Code_and_PC_correct
    '    ' Fichero no existe
    '    If IO.File.Exists(keyfile) = False AndAlso idform = "" Then
    '        ' Entramos la primera vez. Aun no hay idform
    '        Return validacion.File_not_exist
    '        Exit Function
    '    ElseIf IO.File.Exists(keyfile) = False AndAlso idform <> "" Then
    '        ' Entramos la priemra vez. Y hemos aceptado el Code. Hay idform
    '        If ValidIds.Contains(idform) = False Then
    '            Return validacion.Invalid_code
    '            Exit Function
    '        ElseIf ValidIds.Contains(idform) = True Then
    '            ' Escribimos por primera vez key.dat
    '            keyfile_escribe()
    '            Return validacion.Code_and_PC_correct
    '            Exit Function
    '        End If
    '    End If
    '    '
    '    ' Fichero si existe
    '    If IO.File.Exists(keyfile) = True Then
    '        Dim texto As String = crip.Texto_Desencriptar(IO.File.ReadAllText(keyfile))
    '        Dim lineas As String() = texto.Split(vbCrLf)
    '        '
    '        ' Fichero existe. No tiene las 5 lineas que debe tener.
    '        If lineas.Count <> 5 Then
    '            Return validacion.File_incorrect_data
    '            Exit Function
    '        End If
    '        ' 
    '        Dim id As String = lineas(2).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(2))
    '        Dim pc As String = lineas(4).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(4))
    '        ' Fichero existe. ID no es el correcto
    '        If ValidIds.Contains(id) = False Then
    '            Return validacion.id_incorrecto
    '            Exit Function
    '        End If
    '        ' Fichero existe. PC no es el correcto
    '        If pc <> _appComputer Then
    '            Return validacion.PC_incorrect
    '            Exit Function
    '        End If
    '        '
    '        ' ** Correcto **
    '        If ValidIds.Contains(id) = True AndAlso pc = _appComputer Then
    '            Return validacion.Code_PC__correct
    '        End If
    '    End If
    '    Return resultado
    'End Function
    '
    Public Shared Function ID_Comprueba_OnLine() As UCClientWebService.Models.ResponseID
        Dim resultado As ResponseID
        Dim srvId As New UCClientWebService.Services.AddInService
        Dim rId As New RequestID With {.id = idform}
        resultado = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)
        '
        cr = New crip("aiiao2K19")
        ' Fichero no existe
        If IO.File.Exists(keyfile) = False AndAlso resultado.valid = False Then
            ' Entramos la primera vez. Y hemos aceptado el Code. Hay idform
            ' Escribimos por primera vez key.dat
            resultado.id = ""
            resultado.message = "Invalid code"  ' validacion.Invalid_code & "|" & validacion.Invalid_code.ToString.Replace("_", "")
        ElseIf IO.File.Exists(keyfile) = False AndAlso resultado.valid = True Then
            ' Entramos la primera vez. Y hemos aceptado el Code. Hay idform
            ' Escribimos por primera vez key.dat
            keyfile_escribe()
        ElseIf IO.File.Exists(keyfile) = True Then
            ' Fichero si existe
            Dim texto As String = " " & vbCrLf & " " & vbCrLf
            Try
                texto = crip.Texto_Desencriptar(IO.File.ReadAllText(keyfile))
            Catch ex As Exception

            End Try
            Dim lineas As String() = texto.Split(vbCrLf)
            '
            ' Fichero existe. No tiene las 5 lineas que debe tener.
            If lineas.Count <> 5 Then
                IO.File.Delete(keyfile)
                resultado.id = ""
                resultado.valid = False
                resultado.message = "Invalid key data in file"  ' validacion.File_incorrect_data & "|" & validacion.File_incorrect_data.ToString.Replace("_", " ")
            ElseIf lineas.Count = 5 Then
                Dim id As String = lineas(2).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(2))
                Dim pc As String = lineas(4).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(4))
                rId.id = id
                resultado = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)
                ' Fichero existe. ID no es el correcto
                If resultado.valid = False Then
                    resultado.id = ""
                    resultado.message = "Invalid code"  ' validacion.Invalid_code & "|" & validacion.Invalid_code.ToString.Replace("_", "")
                ElseIf resultado.valid = True Then
                    ' Fichero existe, ID correcto. PC no es el correcto
                    If pc <> _appComputer Then
                        resultado.id = ""
                        resultado.valid = False
                        resultado.message = "Invalid pc registration"   ' validacion.PC_name_incorrect & "|" & validacion.PC_name_incorrect.ToString.Replace("_", "")
                    Else
                        resultado.message = "Correct pc and code validation"   ' validacion.Code_and_PC_correct & "|" & validacion.Code_and_PC_correct.ToString.Replace("_", "")
                    End If
                End If
            End If
        End If
        Return resultado
    End Function

    Public Shared Function EstadoRed_Boolean() As Boolean
        Try
            If My.Computer.Network.IsAvailable() Then
                If My.Computer.Network.Ping("www.google.com", 1000) Then 'Asignamos la pagina a consultar ejemplo www.google.com y el tiempo de espera máximo
                    EstadoRed_Boolean = True
                Else
                    EstadoRed_Boolean = False
                End If
            Else
                EstadoRed_Boolean = False
            End If
        Catch ex As Exception
            EstadoRed_Boolean = False
        End Try
    End Function
    Public Shared Function EstadoRed_String() As String
        Try
            If My.Computer.Network.IsAvailable() Then
                If My.Computer.Network.Ping("www.google.com", 1000) Then 'Asignamos la pagina a consultar ejemplo www.google.com y el tiempo de espera máximo
                    EstadoRed_String = ""
                Else
                    EstadoRed_String = "Not internet available"
                End If
            Else
                EstadoRed_String = "Not network available"
            End If
        Catch ex As Exception
            EstadoRed_String = "Not network available"
        End Try
    End Function
    '
    Public Shared Sub keyfile_escribe()
        ' WXdgcIZY k6Szbr16
        Dim lineas As String = ""
        lineas &= "sRiiYuuT" & vbCrLf
        lineas &= "O9Ip06Ty" & vbCrLf
        lineas &= idform & vbCrLf
        lineas &= "7YYuIIop" & vbCrLf
        lineas &= _appComputer
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
