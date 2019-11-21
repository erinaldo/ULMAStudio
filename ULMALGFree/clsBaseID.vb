Imports crip = crip2aCAD.clsCR
Imports UCClientWebService.Models
Imports UCClientWebService.Services


Partial Public Class clsBase
    'Public Shared ReadOnly ValidIds As String() = {"WXdgcIZY", "k6Szbr16"}
    Public Shared idform As String = ""
    Public Shared ReadOnly keyfile As String = IO.Path.Combine(IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location), "key.dat")
    Public Shared resultado As UCClientWebService.Models.ResponseID
    '
    ' Comprobar datos del fichero cuando ya existe key.dat
    Public Shared Function ID_Comprueba_OffLine() As UCClientWebService.Models.ResponseID
        Dim resultado As New ResponseID With {.id = "", .valid = False, .message = ""}
        '
        ' ** Nueva lógica. Solo llamamos a la comprobación Web, si no existe el fichero.
        ' Si ya existe key.dat. Comprobar sintaxis, id <> "" y pc. No comprobación Web, para no aumentar el nº de instalaciones.
        cr = New crip("aiiao2K19")
        If IO.File.Exists(keyfile) = True Then
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
                resultado.message = "Invalid key data in file."  ' validacion.File_incorrect_data & "|" & validacion.File_incorrect_data.ToString.Replace("_", " ")
            ElseIf lineas.Count = 5 Then
                Dim id As String = lineas(2).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(2))
                Dim pc As String = lineas(4).Replace(vbLf, "")    ' crip2aCAD.clsCR.Texto_Desencriptar(lineas(4))
                resultado.id = id
                '
                If id = "" Then
                    ' ID no escorrecto
                    IO.File.Delete(keyfile)
                    resultado.valid = False
                    resultado.message = "Registration ID is not valid (" & "''" & ")"
                ElseIf pc <> _appComputer Then
                    'PC no es el correcto
                    IO.File.Delete(keyfile)
                    resultado.valid = False
                    resultado.message = "Invalid pc validation (" & pc & ")"
                ElseIf id <> "" AndAlso pc = _appComputer Then
                    ' ID con valor y PC correcto
                    resultado.valid = True
                    resultado.message = "Correct code and pc validation (" & id & "/" & pc & ")"
                End If
            End If
        End If
        Return resultado
    End Function
    '
    ' Comprobar a través de la Web id ID (No existe key.dat)
    Public Shared Function ID_Comprueba_OnLine() As UCClientWebService.Models.ResponseID
        Dim resultado As New ResponseID With {.id = "", .valid = False, .message = ""}
        Dim srvId As New UCClientWebService.Services.AddInService
        Dim rId As New RequestID With {.id = idform}
        resultado = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)
        '
        ' ** Nueva lógica. Solo llamamos a la comprobación Web, si no existe el fichero.
        cr = New crip("aiiao2K19")
        If IO.File.Exists(keyfile) = False AndAlso resultado.valid = False Then
            ' Entramos la primera vez. Y hemos aceptado el Code. Hay idform
            ' Escribimos por primera vez key.dat
            resultado.id = ""
            resultado.message = "Registration ID is not valid."  ' validacion.Invalid_code & "|" & validacion.Invalid_code.ToString.Replace("_", "")
        ElseIf IO.File.Exists(keyfile) = False AndAlso resultado.valid = True Then
            ' Entramos la primera vez. Y hemos aceptado el Code. Hay idform
            ' Escribimos por primera vez key.dat
            keyfile_escribe()
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
                    EstadoRed_String = "No internet access available"
                End If
            Else
                EstadoRed_String = "No network connection available"
            End If
        Catch ex As Exception
            EstadoRed_String = "No network connection available"
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
