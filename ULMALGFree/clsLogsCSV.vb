Imports System.Net

Public Class clsLogsCSV
    '
    ' 15 Campos para el csv.
    Private campos() As String = {
        "ACTION", "FILE", "FAMILY", "MARKET", "LANGUAGE", "YEAR", "MONTH", "DAY", "TIME", "COMPUTER_DOMAIN", "COMPUTER_NAME",
        "INTERNAL_IP", "EXTERNAL_IP", "USER_DOMAIN", "USER_NAME", "REVIT_VERSION", "UCREVIT_VERSION", "NOTES"}
    '
    ' Clases
    Public Shared cIni As clsINI = Nothing
    ' Constantes
    Private Const _IniName As String = "UCRevitFree.ini"
    Private Const _IniUpdaterName As String = "ULMALGFreeUpdater.ini"
    ' Variable compartidas por todas las aplicaciones
    Private ReadOnly _folderTemp As String = IO.Path.GetTempPath
    Private ReadOnly _folderGUID As String = "00FABC7172E648958CD6577C1E124CCC"
    Public Shared _appFolderLogs As String = ""    'IO.Path.Combine(_folderTemp, "00FABC7172E648958CD6577C1E124B5A"\")  'IO.Path.Combine(clsLogsCSV._appFolderApp, "logs")            ' Directorio _dirApp & "\logs"
    Public Shared _appFolderLogsBase As String = ""
    Public Shared _appFicherologUlma As String = ""
    Public Shared _appFicherologBASE As String = ""
    Public Shared _ultimaApp As queApp = queApp.UCREVITFREE
    Public Shared _ultimaAccion As String = ACTION.LOAD_FAMILY.ToString
    Public Shared _IniFull As String = ""
    Public Shared _IniPath As String = ""
    Public Shared _IniUpdaterFull As String = ""
    ' Properties
    Private Shared _Market1 As String = ""
    Private Shared _path_offlineBDIdata As String = ""
    Private Shared _path_offlineBDIdataFree As String = ""
    Private Shared _path_families_base As String = ""
    Private Shared _path_families_base_images As String = ""
    Private Shared _path_families_custom As String = ""
    Private Shared _path_families_custom_images As String = ""
    '
    ' Variables personales de cada aplicación
    Public _appFull As String = ""  'System.Reflection.Assembly.GetExecutingAssembly.Location
    Public _appName As String = ""      'Sólo nombre, sin extensión.
    Public _appVersion As String = ""
    Public _appVersionSinYear As String = ""
    Public _appNameVersion As String = ""
    Private _appyear As String = ""
    '
    Public Shared _appComputerDomain As String = Environment.UserDomainName  'System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain.ToString
    Public Shared _appComputer As String = Environment.MachineName
    Public Shared _appIPPrivate As String = UtilesIP.IPPrivada_DameLista(False) ' True=Primera, False=Todas
    Public Shared _appIPPublic As String = UtilesIP.IPPublica_Dame       ' Internet
    Public Shared _appUserDomain As String = Environment.UserDomainName
    Public Shared _appUser As String = Environment.UserName
    Public _appRevitVersion As String = ""
    Public Shared _ActualizarAddIns As Boolean = False
    Public Shared _ActualizarFamilies As Boolean = False
    Public Shared _ActualizarXMLs As Boolean = False
    '
    Friend cFtp As clsFTP = Nothing
    ' ***** Servidor Primario FTP1 (Con valores por defecto) Se vuelven a rellenar desde el .ini
    Friend Const FTP1_host As String = "ftp://ttup.ulmaconstruction.com:21"
    Friend Const FTP1_dirData As String = "ftp://ttup.ulmaconstruction.com/data/"
    Friend FTP1_dirAddins As String = FTP1_dirData & Appyear & "/addins"
    Friend FTP1_dirFamilies As String = FTP1_dirData & Appyear & "/families"
    Friend FTP1_dirXml As String = FTP1_dirData & Appyear & "/XML"
    Friend Const FTP1_dirLog As String = "ftp://ttup.ulmaconstruction.com/log/"
    Friend Const FTP1_user As String = "ftp_revitpub_user"
    Friend Const FTP1_pass As String = "Preh682ht"
    Private instanciado As Boolean = False
    '
    Public terminado As Boolean = False
    Public Shared txtResponse As String = ""
    '
    ' Datos ULMALGFreeUpdater.ini
    Public Shared addins_last As String = ""
    Public Shared families_last As List(Of String) = Nothing
    Public Shared XML_last As String = ""
    '
    Public Shared Property path_offlineBDIdata As String
        Get
            Return _path_offlineBDIdata
        End Get
        Set(value As String)
            _path_offlineBDIdata = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_offlineBDIdata", value)
        End Set
    End Property

    Public Shared Property path_offlineBDIdataFree As String
        Get
            Return _path_offlineBDIdataFree
        End Get
        Set(value As String)
            _path_offlineBDIdataFree = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_offlineBDIdataFree", value)
        End Set
    End Property

    Public Shared Property path_families_base As String
        Get
            Return _path_families_base
        End Get
        Set(value As String)
            _path_families_base = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_families_base", value)
        End Set
    End Property

    Public Shared Property path_families_custom As String
        Get
            Return _path_families_custom
        End Get
        Set(value As String)
            _path_families_custom = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_families_custom", value)
        End Set
    End Property

    Public Shared Property path_families_base_images As String
        Get
            Return _path_families_base_images
        End Get
        Set(value As String)
            _path_families_base_images = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_families_base_images", value)
        End Set
    End Property

    Public Shared Property path_families_custom_images As String
        Get
            Return _path_families_custom_images
        End Get
        Set(value As String)
            _path_families_custom_images = value
            If _appComputer.ToUpper = "ALBERTO-HP" Then
                value = Path_PonRelativo(value)
            End If
            cIni.IniWrite(_IniFull, "PATHS", "path_families_custom_images", value)
        End Set
    End Property

    Public Shared Property Market1 As String
        Get
            Return _Market1
        End Get
        Set(value As String)
            _Market1 = value
        End Set
    End Property

    Public Property Appyear As String
        Get
            Return _appyear
        End Get
        Set(value As String)
            If value = "" Then
                value = "R2019"
            Else
                If IsNumeric(value) And value.Length = 4 Then value = "R" & value
            End If
            _appyear = value
            FTP1_dirAddins = FTP1_dirData & _appyear & "/addins"
            FTP1_dirFamilies = FTP1_dirData & _appyear & "/families"
            FTP1_dirXml = FTP1_dirData & _appyear & "/XML"
        End Set
    End Property
    ' Crear instancia siempre con New(System.Reflection.Assembly.GetExecutingAssembly)
    Public Sub New(assembly As System.Reflection.Assembly)
        ' ***** Datos del ensamblaje que utilizara esta DLL
        If assembly Is Nothing Then
            MsgBox("Not exist this Assembly...")
            Exit Sub
        End If
        If Process.GetProcessesByName("REVIT").Count > 0 Then
            _appRevitVersion = Process_DameVersion("REVIT").Split("("c)(1).Substring(0, 4)
        End If
        '
        If cIni Is Nothing Then cIni = New clsINI()
        _appFull = assembly.Location
        Dim oAinf As New Microsoft.VisualBasic.ApplicationServices.AssemblyInfo(assembly)
        '_appVersion = oAinf.AssemblyName & "_v" & oAinf.Version.ToString
        _appName = oAinf.AssemblyName
        _appVersion = oAinf.Version.ToString
        _appVersionSinYear = QuitaYear()
        _appNameVersion = _appName & " " & _appVersion
        If _appRevitVersion <> "" AndAlso _appRevitVersion.Length >= 4 Then
            Appyear = _appRevitVersion.Substring(0, 4)
        Else
            Appyear = "2019"
        End If
        '
        If _appName.Contains("UCRevitFree") AndAlso _IniFull = "" Then
            _IniFull = IO.Path.Combine(IO.Path.GetDirectoryName(_appFull), _IniName)
            _IniPath = IO.Path.GetDirectoryName(_IniFull)
            _IniUpdaterFull = IO.Path.Combine(IO.Path.GetDirectoryName(_appFull), _IniUpdaterName)
        End If
        '
        ' ***** Verificar si existe fichero log y, si no existe, crearlo para todos los desarrollos
        If _appFolderLogs = "" Then
            _appFolderLogs = IO.Path.Combine(_folderTemp, _folderGUID) 'IO.Path.Combine(_appFolderApp, "logs")
        End If
        '
        If IO.Directory.Exists(_appFolderLogs) = False Then
            Try
                IO.Directory.CreateDirectory(_appFolderLogs)
            Catch ex As Exception
                MsgBox("Folder " & _appFolderLogs & " not exist..." & vbCrLf & vbCrLf & "** Verify permissions or create by hand **")
                Exit Sub
            End Try
        End If
        ' ***** Verificar si existe directorio LOGS base, lo creamos si no existe.
        If _appFolderLogsBase = "" Then
            _appFolderLogsBase = IO.Path.Combine(IO.Path.GetDirectoryName(_appFull), "logs")
        End If
        '
        If IO.Directory.Exists(_appFolderLogsBase) = False Then
            Try
                IO.Directory.CreateDirectory(_appFolderLogsBase)
            Catch ex As Exception
                MsgBox("Folder " & _appFolderLogsBase & " not exist..." & vbCrLf & vbCrLf & "** Verify permissions or create by hand **")
                Exit Sub
            End Try
        End If
        '
        ' fullPath del fichero log (.csv) de Ulma
        If _appFicherologUlma = "" Then
            _appFicherologUlma = LogCsvUlma_DameNombre(_appFolderLogs)
            instanciado = False
        Else
            instanciado = True
        End If

        ' La configuración de los FTPs está ya predefinida en la variables Friend Const de esta clase.
        Try
            Dim resultado As Boolean = CompruebaConexionFTPUlma(SubirBorrar:=True)
            If resultado = True AndAlso instanciado = False Then
                ' Solo al LOG CSV si da error
                'PonLog_ULMA(ACTION.CHECK_FTP, "* OK * (" & cFtp.host & ")")
            ElseIf ULMALGFree.clsLogsCSV.txtResponse <> "" Then
                Throw New Exception()
            End If
        Catch ex As Exception
            Dim cadenaerror As String = txtResponse & " | " & ex.ToString & " | * ERROR * (" & FTP1_host & ")"
            cadenaerror = cadenaerror.Replace(":", "|")
            cadenaerror = cadenaerror.Replace(";", "|")
            cadenaerror = cadenaerror.Replace(vbCrLf, "|")
            cadenaerror = cadenaerror.Replace(Chr(10), "|")
            cadenaerror = cadenaerror.Replace(Chr(13), "|")
            PonLog_ULMA(ACTION.CHECK_FTP_ERROR, NOTES:=cadenaerror)
        End Try
        INIUPDATER_Lee()
    End Sub
    '
    Public Sub INIUPDATER_Lee()

    End Sub
    Public Function CompruebaConexionFTPUlma(Optional SubirBorrar As Boolean = False, Optional creanuevo As Boolean = False, Optional saliendo As Boolean = False) As Boolean
        Dim resultado As Boolean = False
        ' Crear objeto cFtp y comprobar conexiones.
        ' Primero a FTP1 y, si no hay conexión, después a FTP2
        Dim aa As String = ""
        cFtp = New clsFTP(FTP1_host, FTP1_dirLog, FTP1_user, FTP1_pass)
        '
        If cFtp.conexionOk = False OrElse cFtp Is Nothing Then
            cFtp.conexionOk = False
        Else
            cFtp.conexionOk = True
        End If
        '
        If cFtp.conexionOk = True AndAlso SubirBorrar = True Then
            ' borrar otros .csv que existan en el fichero
            ' * No debería haber ninguno, ya que se habrán enviado por FTP y borrado.
            ' * Si hubiera alguno (Ha podido fallar el envio/borrado anterior), enviarlos y borrarlos.
            If IO.Directory.Exists(_appFolderLogs) Then
                ' Subir los ficheros al FTP
                For Each fiCsv As String In IO.Directory.GetFiles(_appFolderLogs, "*.csv", IO.SearchOption.TopDirectoryOnly)
                    ' No subimos ni borramos el fichero actual de LOG.
                    If _appFicherologUlma <> "" And fiCsv = _appFicherologUlma And saliendo = False Then
                        Continue For
                    End If
                    '
                    Dim nombre As String = IO.Path.GetFileName(fiCsv)
                    Dim fiFtp As String = cFtp.dir & nombre
                    Dim resultadoTxt As String = cFtp.FTP_Upload(fiCsv, fiFtp).ToString
                    If resultadoTxt.StartsWith("CORRECTO") Then
                        IO.File.Delete(fiCsv)
                    End If
                Next
            End If
        End If
        If creanuevo = True Then
            Try
                IO.File.AppendAllText(_appFicherologUlma, Join(campos, ";"c), System.Text.Encoding.UTF8)
            Catch ex As Exception
                Process_Close("excel")
                Try
                    IO.File.AppendAllText(_appFicherologUlma, Join(campos, ";"c), System.Text.Encoding.UTF8)
                Catch ex1 As Exception
                End Try
            End Try
        End If
        '
        ' Comprobar los ficheros log Base. Borrar los que sean de más de 10 días.
        If IO.Directory.Exists(_appFolderLogsBase) Then
            For Each fiLog As String In IO.Directory.GetFiles(_appFolderLogs, "*.log", IO.SearchOption.TopDirectoryOnly)
                Dim finf As New IO.FileInfo(fiLog)
                ' No borrar el fichero actual de LOG.
                If _appFicherologBASE <> "" And fiLog = _appFicherologBASE Then
                    Continue For
                End If
                ' No borrar los que sean de menos de 10 días.
                If finf.LastWriteTime.AddDays(10) > Date.Now Then
                    Continue For
                End If
                '
                IO.File.Delete(fiLog)
            Next
        End If
        '
        terminado = True
        resultado = cFtp.conexionOk
        Return resultado
    End Function

    Public Sub PonLog_ULMA(
                          ByVal ACTION As String,
                          Optional FILENAME As String = "",
                          Optional FAMILY As String = "",
                          Optional MARKET As String() = Nothing,
                          Optional LANGUAGE As String() = Nothing,
                          Optional NOTES As String = "")
        '
        Dim _Market As String = ""
        Dim _Language As String = ""
        If MARKET IsNot Nothing AndAlso MARKET.Count > 0 Then _Market = Join(MARKET, "|")
        If LANGUAGE IsNot Nothing AndAlso LANGUAGE.Count > 0 Then _Language = Join(LANGUAGE, "|")
        If NOTES.Contains(vbCrLf) Then NOTES = NOTES.Replace(vbCrLf, ". ")
        If NOTES.Contains(";") Then NOTES = NOTES.Replace(";", "|")
        ULMALGFree.clsLogsCSV._ultimaAccion = ACTION

        '
        'Private campos() As String = {
        '"ACTION", "FILENAME", "FAMILY", "MARKET", "LANGUAGE", "DATE.YEAR", "DATE.MONTH", "DATE.DAY", "TIME", 
        '"COMPUTER_DOMAIN", "COMPUTER_NAME", "INTERNAL_IP", "EXTERNAL_IP", 
        '"USER_DOMAIN", "USER_NAME", "REVIT_VERSION", "UCREVIT_VERSION", "NOTES"}
        Dim valores As String() = {
            ACTION.ToUpper, FILENAME, FAMILY, _Market, _Language, Date.Now.Year, Date.Now.Month, Date.Now.Day, Date.Now.ToString("HH:mm").Split(" "c)(0),
            _appComputerDomain, _appComputer, "IP: " & _appIPPrivate, "IP: " & _appIPPublic,
            _appUserDomain, _appUser, _appRevitVersion, _appNameVersion, NOTES
        }
        '
        ' ***** Comprobar antes que exista el directorio y el fichero
        If IO.Directory.Exists(_appFolderLogs) = False Then
            Try
                IO.Directory.CreateDirectory(_appFolderLogs)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        If IO.File.Exists(_appFicherologUlma) = False Then
            Try
                Call CompruebaConexionFTPUlma(False, creanuevo:=True)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        ' ******************************************************
        Try
            IO.File.AppendAllText(_appFicherologUlma, vbCrLf & Join(valores, ";"c), System.Text.Encoding.UTF8)
        Catch ex As Exception
            Process_Close("excel")
            Try
                IO.File.AppendAllText(_appFicherologUlma, vbCrLf & Join(valores, ";"c), System.Text.Encoding.UTF8)
            Catch ex1 As Exception
            End Try
        End Try
        _ultimaAccion = ACTION
    End Sub

    Public Sub PonLog_ULMA(
                          ByVal ACTION As ACTION,
                          Optional FILENAME As String = "",
                          Optional FAMILY As String = "",
                          Optional MARKET As String() = Nothing,
                          Optional LANGUAGE As String() = Nothing,
                          Optional NOTES As String = "")
        '
        PonLog_ULMA(ACTION.ToString, FILENAME, FAMILY, MARKET, LANGUAGE, NOTES)
    End Sub

#Region "FUNCIONES FECHA"
    Friend Function DateTime_DameFormateado(oDateTime As DateTime, Optional txtFormato As String = "yyyyMMddTHHmmss") As String
        ' Original      =   18/12/2018 13:50:43
        ' Formateado    =   20181218T135043         Con ToString("yyyyMMddTHHmmss")
        Return oDateTime.ToString(txtFormato)
    End Function
    '
    Friend Function LogCsvUlma_DameNombre(queDir As String) As String
        'If queDir = "" Then queDir = _appFolderLogs
        Dim resultado As String = ""
        Dim solodatetime As String = DateTime_DameFormateado(DateTime.Now)
        Dim nombreFi As String = "logR_" &
            solodatetime & "_" &
            Environment.UserDomainName & "_" &
            Environment.MachineName &
            ".csv"
        '
        If queDir <> "" AndAlso IO.Directory.Exists(queDir) Then
            resultado = IO.Path.Combine(queDir, nombreFi)
        Else
            resultado = nombreFi
        End If
        If _appFicherologBASE = "" Then
            _appFicherologBASE = IO.Path.Combine(ULMALGFree.clsLogsCSV._appFolderLogsBase, (solodatetime & ".log"))
        End If
        Return resultado
    End Function
    '
    Public Function LogCsvUlma_DameUltimo() As String
        Dim resultado As String = ""
        Dim filesCSV As String() = IO.Directory.GetFiles(_appFolderLogs, "*.csv")
        If filesCSV Is Nothing OrElse filesCSV.Count = 0 Then
            resultado = ""
        Else
            ' Filtrar los que sean de hoy
            Dim lstHoy = From fi In filesCSV
                         Where IO.File.GetLastWriteTime(fi).Day = Date.Now.Day
                         Select fi

            If lstHoy Is Nothing OrElse lstHoy.Count = 0 Then
                resultado = ""
            Else
                Dim fInf As New IO.FileInfo(lstHoy(0))
                If lstHoy.Count > 1 Then
                    For x As Integer = 1 To lstHoy.Count - 1
                        Dim fInf1 = New IO.FileInfo(lstHoy(x))
                        If fInf1.LastWriteTime > fInf.LastWriteTime Then
                            fInf = New IO.FileInfo(lstHoy(x))
                        End If
                    Next
                End If
                resultado = fInf.FullName
            End If
        End If
        '
        Return resultado
    End Function

    Friend Function Process_DameVersion(queProcess As String) As String
        Dim proc As Process = Process.GetProcessesByName(queProcess).FirstOrDefault
        If proc Is Nothing Then
            Return ""
        Else
            Dim oPInf As System.Diagnostics.ProcessStartInfo = proc.StartInfo
            'Return proc.ProcessName & " (" & proc.MainWindowTitle.Split("-"c)(0).Replace("Autodesk", "").Trim & ")"
            Return proc.ProcessName & " (" & proc.MainModule.FileVersionInfo.ProductVersion & ")"
            'Return proc.MainModule.FileVersionInfo.ProductVersion      ' Falla sin se llama desde 32bits, al ser Revit de 64bis.
        End If
    End Function
    '
    Public Function QuitaYear() As String
        Dim resultado As String = ""
        For Each ch As Char In Me._appVersion
            If IsNumeric(ch) = False Then
                resultado &= ch.ToString
            ElseIf IsNumeric(ch) Then
                Exit For
            End If
        Next
        Return resultado
    End Function
    '
    Public Shared Function Path_PonRelativo(quePath As String) As String
        Return quePath.Replace(_IniPath, ".")
    End Function
    Public Shared Sub Process_Close(name As String)   'Optional queProceso As String = "EXCEL"
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(name)
                Try
                    proc.Kill()
                Catch ex As Exception
                    ' No hacemos nada
                End Try
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
    End Sub
    Public Shared Async Sub Process_Run(quePath As String, Optional argumentos As String = "")
        Dim p As Process = IIf(argumentos = "", Process.Start(quePath), Process.Start(quePath, argumentos))
        Await Task.Run(Sub() p.WaitForExit(15000))
    End Sub
End Class
#End Region
'
#Region "ENUMERACIONES"
Public Enum queApp
    UCREVIT
    UCREVITFREE
    UCBROWSER
    REPCON
    GRAFSYSTEM
    ULMAUPDATERADDIN
End Enum
Public Enum ACTION
    CHECK_FTP
    CHECK_FTP_ERROR
    CHECK_AD
    CHECK_AD_ERROR
    OPEN
    CLOSE
    OPEN_REVIT
    CLOSE_REVIT
    OPEN_FILELINK
    NEW_PROJECT
    NEW_PROJECT_TEMPLATE
    NEW_FAMILY
    NEW_FAMILY_TEMPLATE
    OPEN_PROJECT
    OPEN_PROJECT_TEMPLATE
    OPEN_FAMILY
    OPEN_FAMILY_TEMPLATE
    CLOSE_PROJECT
    CLOSE_PROJECT_TEMPLATE
    CLOSE_FAMILY
    CLOSE_FAMILY_TEMPLATE
    UCBROWSER_LOAD_FAMILY
    GRAFSYSTEM_LOAD_FAMILY
    UCBROWSER_CHANGE_FAMILY_FOLDER
    LOAD_FAMILY
    UCBROWSER_INSERT_FAMILY
    GRAFSYSTEM_INSERT_FAMILY
    INSERT_FAMILY
    SAVE
    SAVEAS
    SAVEAS_LIBRARY_FAMILY
    PRINT_DOCUMENT
    PRINT_VIEW
    SYNCHRONIZE_DOCUMENT
    UCR_OPTIONS
    UCR_CODIFY
    UCR_EXPLODE_FAMILY
    UCR_EXPLODE_DIRECTSHAPE
    UCR_OVERKILL
    UCR_EXPORT_BOM
    UCR_COMPLETE_BOMs
    UCR_REPCON_COMPLETE_BOMs
    UCR_ROTATE
    UCR_ROTATE_POINTS
    UCR_TRANSLATE
    UCR_HELP
    UCR_ABOUT
    UCR_BROWSER
End Enum
Public Enum regLog
    inicio
    registro
    final
End Enum
#End Region
