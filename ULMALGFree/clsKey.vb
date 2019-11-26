Imports crip = crip2aCAD.clsCR
Imports uf = ULMALGFree.clsBase

Friend Class clsKey
    Public l1 As String = ""
    Public l2 As String = ""
    Public l3Id As String = ""
    Public l4 As String = ""
    Public l5Pc As String = ULMALGFree.clsBase._appComputer
    Public l6Connection As String = ""
    Public l7AddIn As String = ""
    Public errorfichero As Boolean = False

    '
    Public Sub New()
        'WXdgcIZY k6Szbr16
        If uf.cr Is Nothing Then uf.cr = New crip("aiiao2K19")
        fiKey_PonDatos()
        If errorfichero = False Then
            Dim newGuid As String = Guid.NewGuid.ToString("N")
            l1 = newGuid.Substring(0, 8)
            l2 = newGuid.Substring(10, 8)
            l4 = newGuid.Substring(20, 8)
            l5Pc = uf._appComputer
            l7AddIn = Date.Now.Ticks.ToString
        End If
    End Sub
    Public Sub keyfile_escribe()
        ' WXdgcIZY k6Szbr16
        Try
            Dim lineas As String = ""
            lineas &= l1 & vbCrLf
            lineas &= l2 & vbCrLf
            lineas &= l3Id & vbCrLf
            lineas &= l4 & vbCrLf
            lineas &= l5Pc & vbCrLf
            lineas &= l6Connection & vbCrLf     ' Ticks de la última conexion para comprobar
            lineas &= l7AddIn                   ' Ticks de la ultima ejecución del Addin
            '

            IO.File.WriteAllText(ULMALGFree.clsBase.keyfile, crip.Texto_Encripta(lineas))
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
    '
    Private Sub fiKey_PonDatos()
        ' Si no existe, salir con los valores por defecto. Que darán error de validación.
        If IO.File.Exists(uf.keyfile) = False Then
            errorfichero = True
            Return
        End If
        '
        '        Dim texto As String = " "
        '        Try
        '            texto = crip.Texto_Desencriptar(IO.File.ReadAllText(uf.keyfile))
        '        Catch ex As Exception
        '        End Try
        '        Dim lineas As String() = texto.Split(vbCrLf)
        '        If lineas.Count <> 7 Then
        '            ' No tiene las 7 lineas que debe tener.
        '            ' (La penultima fila es la última comprobación online en Ticks)
        '            ' (La ultima fila es la fecha de la ultima ejecución del AddIn en Ticks)
        '            RespID.id = ""
        '            RespID.valid = False
        '            RespID.message = "Registration ID is not valid"
        '            RespID.messagelog = "Invalid key data in file (lineas)"
        '            trampeado = True
        '        ElseIf lineas.Count = 7 Then
        '            Dim id As String = lineas(2).Replace(vbLf, "")
        '            Dim pc As String = lineas(4).Replace(vbLf, "")
        '            Dim tTicks As String = lineas(5).Replace(vbLf, "")
        '            Dim nTicks As Long = 0
        '            Dim tTicksAddin As String = lineas(6).Replace(vbLf, "")
        '            Dim nTicksAddin As Long = 0
        '            Dim nDays As Integer = -1        ' Dias desde la última comprobación (>= 0 and < 90 para que sea correcto)
        '            Dim oDate As Date = Nothing
        '            RespID.id = id
        '            ' Comprobar linea 6 (Ticks con la fecha de la última comprobación)
        '            If Long.TryParse(tTicks, nTicks) = True Then
        '                oDate = New Date(nTicks)
        '                nDays = Date.Now.Subtract(oDate).Days
        '                RespID.RemainingDays = nDays
        '            Else
        '                ' tTicks no escorrecto (No es un número)
        '                RespID.valid = False
        '                RespID.message = "Registration ID is not valid"
        '                RespID.messagelog = "Invalid key data in file (Last connection)"
        '                trampeado = True
        '                GoTo FINAL
        '                Exit Sub
        '            End If
        '            ' Comprobar linea 7 (Ticks con la fecha de la última ejecución del AddIn)
        '            If Long.TryParse(tTicksAddin, nTicksAddin) = True Then
        '                oDate = New Date(nTicks)
        '                nDays = Date.Now.Subtract(oDate).Days
        '                RespID.RemainingDays = nDays
        '            Else
        '                ' nTicksAddin no escorrecto (No es un número)
        '                RespID.valid = False
        '                RespID.message = "Registration ID is not valid"
        '                RespID.messagelog = "Invalid key data in file (Last execution Addin)"
        '                trampeado = True
        '                GoTo FINAL
        '                Exit Sub
        '            End If
        '            ' Comprobar ID y PC
        '            If id = "" Then
        '                ' ID no escorrecto
        '                'IO.File.Delete(keyfile)
        '                RespID.valid = False
        '                RespID.message = "Registration ID is not valid"
        '                RespID.messagelog = "Invalid key data in file (id)"
        '                trampeado = True
        '            ElseIf pc <> _appComputer Then
        '                'PC no es el correcto
        '                'IO.File.Delete(keyfile)
        '                RespID.valid = False
        '                RespID.message = "Registration ID is not valid"
        '                RespID.messagelog = "Invalid pc validation (pc)"
        '                trampeado = True
        '            ElseIf nDays < 0 OrElse nDays > 90 Then
        '                'Han pasado los 90 días sin conexión
        '                'IO.File.Delete(keyfile)
        '                RespID.valid = False
        '                RespID.message = "Connection needed to verify registration ID (offline period expired)"
        '                RespID.messagelog = "Network connection off. Expired (" & nDays & ")"
        '            ElseIf Date.Now.Ticks < nTicksAddin Then
        '                'Han cambiado lo fecha del equipo
        '                RespID.valid = False
        '                RespID.message = "Connection needed to verify registration ID (offline period expired)"
        '                RespID.messagelog = "Network connection off. Expired (" & nDays & ")"
        '                trampeado = True
        '            ElseIf id <> "" AndAlso pc = _appComputer Then
        '                rId.id = id
        '                If networkinternet = "" Then
        '                    ' SI hay conexión / internet
        '                    ' Comprobar ID via Web y recoger resultado
        '                    RespID = srvId.IsValidAsync("https://www.ulmaconstruction.com/@@bim_form_api", rId)
        '                    RespID.messagelog = RespID.message
        '                Else
        '                    ' NO hay conexión / internet
        '                    ' Validación temporal (máximo 90 días)
        '                    RespID.valid = True
        '                    RespID.message = networkinternet
        '                    RespID.messagelog = networkinternet & " (id=" & id & "|pc=" & pc & "|RemainingDays= " & RespID.RemainingDays.ToString & ")"
        '                End If
        '            End If
        '        End If
        'FINAL:
        '        ' Si validación incorrecta borrar key.data
        '        If (RespID.valid = False AndAlso networkinternet = "") OrElse trampeado = True Then
        '            IO.File.Delete(keyfile)
        '        End If
    End Sub
End Class
