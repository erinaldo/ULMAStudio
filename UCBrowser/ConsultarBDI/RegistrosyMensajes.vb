Imports System.Configuration

    ''// Ejemplo de uso:
    '//
    '    Dim mensajero As RegistrosyMensajes
    '    mensajero = New RegistrosyMensajes("nombredelmoduloquelovaautilizar")
    '//
    '    mensajero.getMensajero.DarAviso(Nivel:=nivel.n3_NOTICE, _
    '                     Mensaje:=ex.Message, _
    '                     Procedencia:=ex.StackTrace)
    '//
    '/*
    '/***********************************************************************************
    '* Se contemplan cinco niveles:
    '// EMERG → peligro grave, error irrecuperable, sistema ha quedado inutilizable; abortar.
    '// WARNING → peligro leve, advertencia, error recuperable, se puede seguir trabajando; pero ¡ojo!.
    '// NOTICE → informativo, situación normal, significativa|importante.
    '// INFO → informativo, situación normal, detalles para usuarios.
    '// DEBUG → detalles finos, para técnicos.
    '*/
    '/***********************************************************************************
    '* La configuracion de como actuar, se lee del archivo app.config
    '* (nota: Si no se pueden leer de ahí, se pondran unos ciertos valores por defecto que estan hardcoded en el constructor.)
    '*
    '* En la seccion '<appSettings>' de 'app.config' de nuestra aplicacion, hay que poner las siguientes lineas:
    '*
    '  <appSettings>
        '<!--==== INICIO RegistrosYMensajes BEGIN logs&messages ====-->
        '<!--Los niveles disponibles son: "NONE" , "DEBUG" , "INFO" , "NOTICE" , "WARNING" o "EMERG"-->
        '<!--Allowed levels are:  "NONE" , "DEBUG" , "INFO" , "NOTICE" , "WARNING" o "EMERG"-->
        '<!--      show to user (selected and above levels)-->
        '<add key="NivelAPartirDelCualMostrarMensajeAlUsuario" value="WARNING" />
        '<!--      display into console (selected and above levels)-->
        '<add key="NivelAPartirDelCualEscribirEnLaConsola" value="INFO" />
        '<!--      save on file (selected and above levels)-->
        '<add key="NivelAPartirDelCualGuardarRegistroLog" value="INFO" />
        '<!--Si no queremos indicar un path concreto (terminado en \), podemos poner "USER" para indicar "...MyDocuments\_registrosLOG_\"-->
        '<!--If you don't want to supply a path (\ ended), you can put "USER" and logs will be writen to "...MyDocuments\_registrosLOG_\   folder"-->
        '<add key="PathDondeGuardarRegistroLog" value="USER" />
        '<!--Max size (in bytes) for each log file (x3 files)-->
        '<add key="TamañoMaximoBytesParaRegistroLog" value="524288" />
        '<!--==== FIN RegistrosYMensajes END logs&messages ====-->
    '  </appSettings>
    '*
    '*************************************************************************************/

Public Class RegistrosyMensajes

    Const SUFIJOPARAIDENTIFICARELARCHIVOLOG As String = "ConsultarBDI"

    Private gNivelParaMensaje As nivel
    Private gNivelParaConsola As nivel
    Private gNivelParaRegistro As nivel
    Private gPathArchivoLog As String
    Private gSufijoArchivoLog As String
    Private gTamañoMaximoArchivoLog As Integer

    ''// Normalmente se procesaran todos los avisos,
    ''// pero habra ocasiones en que una determinada operacion puede provocar una cascada de avisos o de errores y nos pueda interesar omitir algunos de ellos.
    Private gProcesarAvisosRepetitivos As Boolean = True



    Private Shared m_mensajero As RegistrosyMensajes
    Public Shared Function getMensajero() As RegistrosyMensajes
        If IsNothing(m_mensajero) Then
            m_mensajero = New RegistrosyMensajes(sufijoParaIdentificarElArchivoLog:=SUFIJOPARAIDENTIFICARELARCHIVOLOG)
        End If
        Return m_mensajero
    End Function
    Private Sub New(sufijoParaIdentificarElArchivoLog As String)
        LeerConfiguracionYAjustarParametros(sufijoParaIdentificarElArchivoLog)
    End Sub
    Private Sub LeerConfiguracionYAjustarParametros(sufijoParaIdentificarElArchivoLog As String)

        gSufijoArchivoLog = sufijoParaIdentificarElArchivoLog

        Try
            gNivelParaMensaje = getElNivelCorrespondienteA(ConfigurationManager.AppSettings.Get("NivelAPartirDelCualMostrarMensajeAlUsuario"))
        Catch e As Exception
            gNivelParaMensaje = nivel.n3_NOTICE
        End Try
        Try
            gNivelParaConsola = getElNivelCorrespondienteA(ConfigurationManager.AppSettings.Get("NivelAPartirDelCualEscribirEnLaConsola"))
        Catch e As Exception
            gNivelParaConsola = nivel.n2_INFO
        End Try
        Try
            gNivelParaRegistro = getElNivelCorrespondienteA(ConfigurationManager.AppSettings.Get("NivelAPartirDelCualGuardarRegistroLog"))
        Catch e As Exception
            gNivelParaRegistro = nivel.n2_INFO
        End Try

        Try
            gPathArchivoLog = ConfigurationManager.AppSettings.Get("PathDondeGuardarRegistroLog")
            If IsNothing(gPathArchivoLog) Or gPathArchivoLog = "USER" Then
                gPathArchivoLog = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & _
                                  IO.Path.DirectorySeparatorChar & "_registrosLOG_" & IO.Path.DirectorySeparatorChar
            End If
        Catch ex As Exception
            gPathArchivoLog = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & _
                              IO.Path.DirectorySeparatorChar & "_registrosLOG_" & IO.Path.DirectorySeparatorChar
        End Try
        If Not IO.Directory.Exists(gPathArchivoLog) Then IO.Directory.CreateDirectory(gPathArchivoLog)

        Try
            gTamañoMaximoArchivoLog = CInt(ConfigurationManager.AppSettings.Get("TamañoMaximoBytesParaRegistroLog"))
        Catch ex As Exception
            gTamañoMaximoArchivoLog = 524288
        End Try
        If gTamañoMaximoArchivoLog < 1 Then gTamañoMaximoArchivoLog = 524288

    End Sub




    Public Sub DarAviso(ByVal Nivel As nivel, _
                        ByVal Mensaje As String, _
                        Optional ByVal Procedencia As String = "", _
                        Optional ByVal EsAvisoRepetitivo As Boolean = False)

        If Nivel >= gNivelParaMensaje Then
            If EsAvisoRepetitivo = False Or (EsAvisoRepetitivo And gProcesarAvisosRepetitivos) Then
                Dim estilo As MsgBoxStyle
                Select Case Nivel
                    Case RegistrosyMensajes.nivel.n5_EMERG
                        estilo = MsgBoxStyle.Critical
                    Case RegistrosyMensajes.nivel.n4_WARNING
                        estilo = MsgBoxStyle.Exclamation
                    Case Else
                        estilo = MsgBoxStyle.Information
                End Select
                MsgBox(Title:="[" & getElTextoCorrespondienteA(Nivel) & "] " _
                     , Prompt:=Mensaje & Environment.NewLine & Environment.NewLine _
                               & Procedencia _
                     , Buttons:=estilo)
            End If
        End If


        If Nivel >= gNivelParaConsola Then
            If Nivel >= RegistrosyMensajes.nivel.n3_NOTICE Then
                System.Console.WriteLine("[" & getElTextoCorrespondienteA(Nivel) & "] " & Mensaje)
            Else
                System.Console.WriteLine(Mensaje)
            End If
        End If


        If Nivel >= gNivelParaRegistro Then
            Dim registro As String
            registro = Environment.NewLine & Environment.NewLine _
                       & "------------begin-------------" & Now.ToString("yyyy/MM/dd--HH:mm:ss:ffff--") & Environment.NewLine
            If Nivel >= RegistrosyMensajes.nivel.n3_NOTICE Then
                registro = registro _
                           & "[" & getElTextoCorrespondienteA(Nivel) & "]" & Environment.NewLine
            End If
            registro = registro _
                       & Mensaje & Environment.NewLine _
                       & "----" & Environment.NewLine _
                       & Procedencia & Environment.NewLine _
                       & "-------------end--------------" & Environment.NewLine & Environment.NewLine
            GuardarEnArchivoLog(registro)
        End If

    End Sub



    ''// Los prefijos nX_ son para que los niveles aparezcan ordenados al escribirlos en el codigo.
    Public Enum nivel
        n99_none = 99
        n1_DEBUG = 1
        n2_INFO = 2
        n3_NOTICE = 3
        n4_WARNING = 4
        n5_EMERG = 5
    End Enum

    Private Function getElNivelCorrespondienteA(nivel As String) As nivel
        Select Case nivel
            Case "NONE"
                Return RegistrosyMensajes.nivel.n99_none
            Case "DEBUG"
                Return RegistrosyMensajes.nivel.n1_DEBUG
            Case "INFO"
                Return RegistrosyMensajes.nivel.n2_INFO
            Case "NOTICE"
                Return RegistrosyMensajes.nivel.n3_NOTICE
            Case "WARNING"
                Return RegistrosyMensajes.nivel.n4_WARNING
            Case "EMERG"
                Return RegistrosyMensajes.nivel.n5_EMERG
            Case Else
                Throw New ArgumentOutOfRangeException(paramName:="nivel", actualValue:=nivel, message:="Ha de ser: NONE, DEBUG, INFO, NOTICE, WARNING o EMERG")
        End Select
    End Function

    Private Function getElTextoCorrespondienteA(nivel As nivel) As String
        Select Case nivel
            Case RegistrosyMensajes.nivel.n99_none
                Return "NONE"
            Case RegistrosyMensajes.nivel.n1_DEBUG
                Return "DEBUG"
            Case RegistrosyMensajes.nivel.n2_INFO
                Return "INFO"
            Case RegistrosyMensajes.nivel.n3_NOTICE
                Return "NOTICE"
            Case RegistrosyMensajes.nivel.n4_WARNING
                Return "WARNING"
            Case RegistrosyMensajes.nivel.n5_EMERG
                Return "EMERG"
            Case Else
                Throw New ArgumentOutOfRangeException(paramName:="nivel", actualValue:=nivel.ToString(), message:="")
        End Select
    End Function


    '/**
    '* El registro en archivo es circular; para conservar como minimo los ultimos 2x bytes de datos;
    '* pero asegurandonos de que todo el registro nunca superara los 3x bytes de datos.
    '*/
    Private Sub GuardarEnArchivoLog(ByVal mensaje As String)
    Try
        Dim registro As IO.StreamWriter

        Dim archivo1, archivo2, archivo3 As IO.FileInfo
        archivo1 = New IO.FileInfo(gPathArchivoLog & "LOG-" & gSufijoArchivoLog & "-1.txt")
        archivo2 = New IO.FileInfo(gPathArchivoLog & "LOG-" & gSufijoArchivoLog & "-2.txt")
        archivo3 = New IO.FileInfo(gPathArchivoLog & "LOG-" & gSufijoArchivoLog & "-3.txt")

        If archivo1.Exists Then
            If archivo1.Length < gTamañoMaximoArchivoLog Then
                registro = archivo1.AppendText()
                If archivo2.Exists Then
                    If archivo2.Length > 0 Then
                        archivo2.Delete()
                        archivo2.Create().Dispose()
                    End If
                Else
                    archivo2.Create().Dispose()
                End If
            Else
                If archivo2.Exists Then
                    If archivo2.Length < gTamañoMaximoArchivoLog Then
                        registro = archivo2.AppendText()
                        If archivo3.Exists Then
                            If archivo3.Length > 0 Then
                                archivo3.Delete()
                                archivo3.Create().Dispose()
                            End If
                        Else
                            archivo3.Create().Dispose()
                        End If
                    Else
                        If archivo3.Exists Then
                            If archivo3.Length < gTamañoMaximoArchivoLog Then
                                registro = archivo3.AppendText()
                                If archivo1.Exists Then
                                    If archivo1.Length > 0 Then
                                        archivo1.Delete()
                                        archivo1.Create().Dispose()
                                    End If
                                Else
                                    archivo1.Create().Dispose()
                                End If
                            Else
                                archivo1.Delete()
                                registro = archivo1.CreateText()
                            End If
                        Else
                            registro = archivo3.CreateText()
                        End If
                    End If
                Else
                    registro = archivo2.CreateText()
                End If
            End If
        Else
            registro = archivo1.CreateText()
        End If

        registro.Write(mensaje)
        registro.Flush()
        registro.Dispose()

    Catch ex As IO.IOException
        MsgBox("Problema guardando registro . Problem saving to log file" & Environment.NewLine & _
                gPathArchivoLog & Environment.NewLine & gSufijoArchivoLog & Environment.NewLine & Environment.NewLine & _
                ex.StackTrace, MsgBoxStyle.Exclamation)
    End Try
    End Sub



End Class
