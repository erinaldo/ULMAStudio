Imports System.Net
Imports System.Xml
Imports System.Linq
Imports System.Collections

Public Class clsBase
    Public Shared cr As crip2aCAD.clsCR
    Public Shared frmUFam As frmUpdater
    'Public Shared cLcsv As clsBase = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)
    ' 15 Campos para el csv.
    Private campos() As String = {
        "ACTION", "FILE", "FAMILY", "MARKET", "LANGUAGE", "YEAR", "MONTH", "DAY", "TIME", "COMPUTER_DOMAIN", "COMPUTER_NAME",
        "INTERNAL_IP", "EXTERNAL_IP", "USER_DOMAIN", "USER_NAME", "REVIT_VERSION", "UCREVIT_VERSION", "SYSTEM_DATE",
        "UPDATE_GROUP", "UPDATE_FILES", "TYPE", "KEYCODE", "NOTES"}
    'UPDATE_GROUP    'escribir el nombre del grupo a descargar/actualizar/eliminar
    'UPDATE_FILES    'escribir el nombre del .zip que se va a descargar/actualizar
    '
    ' OBJETOS QUE GENERAN EVENTOS
    Public Shared evAppUIC As Autodesk.Revit.UI.UIControlledApplication
    Public Shared evAppC As Autodesk.Revit.ApplicationServices.ControlledApplication
    Public Shared yo As clsBase = Nothing
    ' Clases
    Public Shared cIni As clsINI = Nothing
    ' Constantes
    Public Const comillas As String = Chr(34)
    Public Const formatofecha As String = "yyyy/MM/dd HH:mm:ss"
    Public Const _IniName As String = "ULMAStudio.ini"
    Public Const _AddInName As String = "ULMAStudio"
    ' Variable compartidas por todas las aplicaciones
    Public Shared arrM As String() = Array.Empty(Of String)
    Public Shared arrL As String() = Array.Empty(Of String)
    Public Shared colIdiomas As SortedList         '' Coleccion de idiomas para codigos/descripciones (key=idioma, Value=Pais)
    Public Shared arrIdiomas As ArrayList          '' Array de idiomas
    Public Shared colMercadosDes As SortedList        '' Coleccion de Mercados (key=descripcion, value=codigo,descripcion,idioma,pais)
    Public Shared colMercadosCod As SortedList        '' Coleccion de Mercados (key=codigo, value=codigo,descripcion,idioma,pais)
    Public Shared arrMercadosDes As ArrayList         '' array de Mercados (Con descripcion)
    Public Shared arrMercadosCod As ArrayList         '' array de Mercados (Con codigo)

    Public Shared unidPe As String = ""                ' unidPeso (para cabeceras ficheros CSV)
    Public Shared unidAr As String = ""                ' unidArea (para cabeceras ficheros CSV)
    Public Shared unidLo As String = ""                ' unidLength (para cabeceras ficheros CSV)
    Public Shared ReadOnly s As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
    Public Shared ReadOnly _sepdecimal As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyDecimalSeparator
    Public Shared ReadOnly _sepmiles As String = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.CurrencyGroupSeparator
    Public Shared ReadOnly _LgFull As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    Public Shared ReadOnly _LgFullFolder As String = IO.Path.GetDirectoryName(_LgFull)
    Public Shared ReadOnly _BrowserFull As String = IO.Path.Combine(_LgFullFolder, "UCBrowser", "UCBrowser.dll")
    Public Shared ReadOnly _BrowserFolder As String = IO.Path.GetDirectoryName(_BrowserFull)
    Public Shared ReadOnly _folderTemp As String = IO.Path.GetTempPath
    Public Shared ReadOnly _xmlFull As String = IO.Path.Combine(_LgFullFolder, "offlineBDIdata", "120_publicStructure.xml")
    Public Shared ReadOnly _imgFolder As String = IO.Path.Combine(_LgFullFolder, "images\")
    Public Shared ReadOnly _ULMAStudioReport As String = IO.Path.Combine(_LgFullFolder, "ULMAStudioReport.txt")
    Public Shared _imgBasePath As String = IO.Path.Combine(_imgFolder, "render_X.png")
    Public Shared _imgBase As System.Drawing.Image
    '
    Public Const _folderGUID As String = "00FABC7172E648958CD6577C1E124CCC"
    Public Shared ReadOnly _appLogCSVFolder As String = IO.Path.Combine(_folderTemp, _folderGUID)
    Public Shared _appLogCSVFichero As String = ""
    Public Shared ReadOnly _appLogBaseFolder As String = IO.Path.Combine(_LgFullFolder, "logs")
    Public Shared _appLogBaseFichero As String = ""
    Public Shared _ultimaApp As queApp = queApp.ULMASTUDIO
    Public Shared _ultimaAccion As String = action.LOAD_FAMILY.ToString
    Public Shared ReadOnly _IniFull As String = IO.Path.Combine(_LgFullFolder, _IniName)
    ' Properties
    Private Shared _DEFAULT_PROGRAM_MARKET As String = "120"
    Private Shared _DEFAULT_PROGRAM_LANGUAGE As String = "en"
    Private Shared _path_offlineBDIdata As String = ""
    Private Shared _path_families_base As String = ""
    Private Shared _path_families_base_images As String = ""
    Private Shared _path_families_custom As String = ""
    Private Shared _path_families_custom_images As String = ""

    'Recarga Browser y cosas del Browser
    Public Shared _recargarBrowser As Boolean = False
    Public Shared _registraLoadInsert As Boolean = True
    'cambios en Grupos para enviar llamada de cierre al Browser
    Public Shared cambiosEnGrupos As Boolean = False
    '
    ' Variables personales de cada Assembly
    Public _asmFull As String = ""  'System.Reflection.Assembly.GetExecutingAssembly.Location
    Public _asmName As String = ""      'Sólo nombre, sin extensión.
    Public _asmVersion As String = ""
    Public _asmVersionSinYear As String = ""
    Public _asmNameVersion As String = ""
    '
    ' Variable generales del equipo y Revit
    Public Shared ReadOnly _appComputerDomain As String = Environment.UserDomainName  'System.DirectoryServices.ActiveDirectory.Domain.GetComputerDomain.ToString
    Public Shared ReadOnly _appComputer As String = Environment.MachineName
    Public Shared ReadOnly _appIPPrivate As String = UtilesIP.IPPrivada_DameLista(False) ' True=Primera, False=Todas
    Public Shared ReadOnly _appIPPublic As String = UtilesIP.IPPublica_Dame       ' Internet
    Public Shared ReadOnly _appUserDomain As String = Environment.UserDomainName
    Public Shared ReadOnly _appUser As String = Environment.UserName
    'Public Shared _appSubVersionNumber As String = "2018.3"
    'Public Shared _appVersionBuild As String = "20190510_1515(X64)"
    'Public Shared _appVersionName As String = "Autodesk Revit 2018"
    'Public Shared _appVersionNumber As String = "2018"
    'Public Shared _appVersionNumberR As String = "R2018"
    'Public Shared _appVersionLogText As String = "Revit 2018.3 (20190510_1515(X64))"
    Public Shared oVersions As clsVersions = Nothing        ' Clase con Dictorionary de todas las clsVersion
    Public Shared oVersion As clsVersion = Nothing          ' Clase con todos los datos de la versión actual de Revit
    Public Shared _ActualizarAddIns As Boolean = False
    Public Shared _ActualizarFamilies As Boolean = False
    Public Shared _ActualizarXMLs As Boolean = False
    Public Shared _ULMAStudioVersion As String = "ULMAStudio 2018.0.0.16"
    Public Shared _UCBrowserVersion As String = "UCBrowser 2018.0.0.8"
    '
    Friend Shared cFtp As clsFTP = Nothing
    ' ***** Servidor Primario FTP1 (Con valores por defecto) Se vuelven a rellenar desde el .ini
    Public Const tipo As String = "develop"
    Friend Const FTP1_host As String = "ftp://ttup.ulmaconstruction.com:21"
    Friend Const FTP1_dirData As String = "ftp://ttup.ulmaconstruction.com/" & tipo & "/"
    Friend Const FTP1_dirLog As String = "ftp://ttup.ulmaconstruction.com/logs/Internal/"
    Friend Const FTP1_admin As String = "ftp_revitpub_admin"
    Friend Const FTP1_adminpass As String = "PReA685P"
    Friend Const FTP1_user As String = "ftp_revitpub_user"
    Friend Const FTP1_userpass As String = "Preh682ht"
    Private Const _FTP1_dirAddins As String = "/addins"
    Private Const _FTP1_dirFamilies As String = "/families"
    Private Const _FTP1_dirXml As String = "/XML"
    Private instanciado As Boolean = False
    '
    Public terminado As Boolean = False
    Public Shared txtResponse As String = ""
    '
    Public Shared onlywithimage As Boolean = False
    ' Datos ULMALGFreeUpdater.ini
    Public Shared addins_last As String = ""
    Public Shared families_last As List(Of String) = Nothing
    Public Shared XML_last As String = ""
    ' Colecciones de datos XML

    'Public Shared cFC As clsFamCode
    'Public Shared cFam As clsFamilia
    Public Shared cComs As clsCompanies
    '
    Public Shared colArticulos As Dictionary(Of String, clsArticulos)        '' Dictionary de Articulos leidos del XML de la compañia seleccionada (Key=artcode (ITEM_CODE), Value=clsArticulos)
    Public Shared colGroups As Dictionary(Of String, clsGroup)        '' Dictionary de Grupos leidos del XML de la compañia seleccionada (Key=groupCode, Value=clsGroups)
    '

    Public Shared Property path_offlineBDIdata As String
        Get
            Return _path_offlineBDIdata
        End Get
        Set(value As String)
            If cIni Is Nothing Then cIni = New clsINI()
            If value <> _path_offlineBDIdata Then
                _path_offlineBDIdata = value
                If _appComputer.ToUpper = "ALBERTO-HP" Then
                    value = Path_PonRelativo(value)
                End If
                cIni.IniWrite(_IniFull, "PATHS", "path_offlineBDIdata", value)
            End If
        End Set
    End Property

    Public Shared Property path_families_base As String
        Get
            Return _path_families_base
        End Get
        Set(value As String)
            If cIni Is Nothing Then cIni = New clsINI()
            If value <> _path_families_base Then
                _path_families_base = value
                If _appComputer.ToUpper = "ALBERTO-HP" Then
                    value = Path_PonRelativo(value)
                End If
                cIni.IniWrite(_IniFull, "PATHS", "path_families_base", value)
            End If
        End Set
    End Property

    Public Shared Property path_families_custom As String
        Get
            Return _path_families_custom
        End Get
        Set(value As String)
            If cIni Is Nothing Then cIni = New clsINI()
            If value <> _path_families_custom Then
                _path_families_custom = value
                If _appComputer.ToUpper = "ALBERTO-HP" Then
                    value = Path_PonRelativo(value)
                End If
                cIni.IniWrite(_IniFull, "PATHS", "path_families_custom", value)
            End If
        End Set
    End Property

    Public Shared Property path_families_base_images As String
        Get
            Return _path_families_base_images
        End Get
        Set(value As String)
            If cIni Is Nothing Then cIni = New clsINI()
            If value <> _path_families_base_images Then
                _path_families_base_images = value
                If _appComputer.ToUpper = "ALBERTO-HP" Then
                    value = Path_PonRelativo(value)
                End If
                cIni.IniWrite(_IniFull, "PATHS", "path_families_base_images", value)
            End If
        End Set
    End Property

    Public Shared Property path_families_custom_images As String
        Get
            Return _path_families_custom_images
        End Get
        Set(value As String)
            If cIni Is Nothing Then cIni = New clsINI()
            If value <> _path_families_custom_images Then
                _path_families_custom_images = value
                If _appComputer.ToUpper = "ALBERTO-HP" Then
                    value = Path_PonRelativo(value)
                End If
                cIni.IniWrite(_IniFull, "PATHS", "path_families_custom_images", value)
            End If
        End Set
    End Property

    Public Shared Property DEFAULT_PROGRAM_MARKET As String
        Get
            If _DEFAULT_PROGRAM_MARKET.Trim = "" OrElse IsNumeric(_DEFAULT_PROGRAM_MARKET.Trim) = False Then _DEFAULT_PROGRAM_MARKET = "120"
            Return _DEFAULT_PROGRAM_MARKET
        End Get
        Set(value As String)
            _DEFAULT_PROGRAM_MARKET = value.Trim
        End Set
    End Property

    Public Shared Property DEFAULT_PROGRAM_LANGUAGE As String
        Get
            If _DEFAULT_PROGRAM_LANGUAGE.Trim = "" Then _DEFAULT_PROGRAM_LANGUAGE = "en"
            Return _DEFAULT_PROGRAM_LANGUAGE
        End Get
        Set(value As String)
            _DEFAULT_PROGRAM_LANGUAGE = value.Trim
        End Set
    End Property
    'Public Sub New()
    '    ' Activar DLL de encriptar/desencriptar
    '    cr = New crip2aCAD.clsCR("aiiao2k19")
    '    ' *************************************
    'End Sub

    ' Crear instancia siempre con New(System.Reflection.Assembly.GetExecutingAssembly)
    Public Sub New(assembly As System.Reflection.Assembly)
        ' Activar DLL de encriptar/desencriptar
        cr = New crip2aCAD.clsCR("aiiao2K19")
        ' *************************************
        'System.Windows.Forms.Application.EnableVisualStyles()
        ' ***** Datos del ensamblaje que utilizara esta DLL
        If assembly Is Nothing Then Exit Sub
        '
        If cIni Is Nothing Then cIni = New clsINI()
        '
        _asmFull = assembly.Location
        Dim oAinf As New Microsoft.VisualBasic.ApplicationServices.AssemblyInfo(assembly)
        '_appVersion = oAinf.AssemblyName & "_v" & oAinf.Version.ToString
        _asmName = oAinf.AssemblyName
        _asmVersion = oAinf.Version.ToString
        _asmVersionSinYear = QuitaYear()
        _asmNameVersion = _asmName & " " & _asmVersion
        Version_Put()
        '
        ' ESTA CONFIGURACIÓN GENERAL SÓLO SE CARGA UNA VEZ. Al instanciar Assembly UCRevitFree
        If _asmName.Contains("ULMAStudio") OrElse _asmName.Contains("ULMAUpdaterAddIn") Then
            Logs_BaseYCSV_PonNombres()
            ' ***** Verificar si existe directorio LOGS base, lo creamos si no existe.
            If IO.Directory.Exists(_appLogBaseFolder) = False Then
                Try
                    IO.Directory.CreateDirectory(_appLogBaseFolder)
                Catch ex As Exception
                    'MsgBox("Folder " & _appFolderLogs & " not exist..." & vbCrLf & vbCrLf & "** Verify permissions or create by hand **")
                    Exit Sub
                End Try
            End If
            If IO.Directory.Exists(_appLogCSVFolder) = False Then
                Try
                    IO.Directory.CreateDirectory(_appLogCSVFolder)
                Catch ex As Exception
                    'MsgBox("Folder " & _appFolderLogsBase & " not exist..." & vbCrLf & vbCrLf & "** Verify permissions or create by hand **")
                    Exit Sub
                End Try
            End If
            ' La configuración de los FTPs está ya predefinida en la variables Friend Const de esta clase.
            Try
                Dim resultado As Boolean = CompruebaConexionFTPUlma(SubirBorrar:=True)
                If resultado = True Then
                    ' Solo al LOG CSV si da error
                    'PonLog_ULMA(ACTION.CHECK_FTP, "* OK * (" & cFtp.host & ")")
                ElseIf ULMALGFree.clsBase.txtResponse <> "" Then
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
            'Call INI_UpdatesLee()
            Dim partes() As String = _asmVersion.Split("."c)
            partes(0) = oVersion.RevitVersionNumber
            If _asmName.Contains("ULMAStudio") Then
                _ULMAStudioVersion = _asmName & " " & Join(partes, "."c)
            ElseIf _asmName.Contains("UCBrowser") Then
                _ULMAStudioVersion = _asmName & " " & Join(partes, "."c)
            End If
        End If
    End Sub
    '
    Public Shared Sub Version_Put()
        If evAppC IsNot Nothing Then
            oVersion = clsVersions.Version_Dame(evAppC, Process.GetCurrentProcess)
        Else
            oVersion = clsVersions.Version_Dame(Process.GetCurrentProcess)
        End If
    End Sub

    Public Shared Sub RenombraExeBak()
        For Each fi As String In IO.Directory.GetFiles(_LgFullFolder, "*.exe.bak", IO.SearchOption.AllDirectories)
            Dim fiExe As String = fi.Replace(".exe.bak", ".exe")
            'If IO.File.Exists(fiExe) Then
            IO.File.Copy(fi, fiExe, True)
            IO.File.Delete(fi)
            'End If
        Next
        If IO.File.Exists(_BatUpdateFull) Then IO.File.Delete(_BatUpdateFull)
    End Sub
    Public Shared Function RevitFull() As String
        Version_Put()
        'Public Shared ReadOnly _Revit2018Full As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\Revit\2018\REVIT-05:040A\", "InstallationLocation", "C:\Program Files\Autodesk\Revit 2018\").ToString & "Revit.exe"
        'Public Shared ReadOnly _Revit2019Full As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\Revit\2019\REVIT-05:040A\", "InstallationLocation", "C:\Program Files\Autodesk\Revit 2019\").ToString & "Revit.exe"
        'Public Shared ReadOnly _Revit2020Full As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\Revit\2020\REVIT-05:040A\", "InstallationLocation", "C:\Program Files\Autodesk\Revit 2020\").ToString & "Revit.exe"
        Dim version As String = oVersion.RevitVersionNumber
        Dim folder As String = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\SOFTWARE\Autodesk\Revit\" & version & "\REVIT-05:040A\", "InstallationLocation", "C:\Program Files\Autodesk\Revit " & version & "\").ToString
        Return folder & "Revit.exe"
    End Function
    Public Shared Function FTP1_dirAddins() As String
        Version_Put()
        Return FTP1_dirData & oVersion.RevitVersionNumberR & _FTP1_dirAddins
    End Function

    Public Shared Function FTP1_dirFamilies() As String
        Version_Put()
        Return FTP1_dirData & oVersion.RevitVersionNumberR & _FTP1_dirFamilies
    End Function

    Public Shared Function FTP1_dirXml() As String
        Version_Put()
        Return FTP1_dirData & oVersion.RevitVersionNumberR & _FTP1_dirXml
    End Function

    'Public Shared Function INI_BaseLee() As String
    '    If cIni Is Nothing Then cIni = New clsINI
    '    Dim resultado As String = ""
    '    '
    '    '
    '    ' LDatos
    '    Dim updates() As String = cIni.IniGetSection(_IniUpdaterFull, "UPDATES")
    '    LDatos = New List(Of Datos)
    '    For i As Integer = 0 To UBound(updates) - 1 Step 2
    '        Dim Clave As String = updates(i)
    '        Dim Valores() As String = updates(i + 1).Split("·")
    '        Dim d As New Datos(Valores(0), Valores(1), Valores(2))
    '        If LDatos.Contains(d) = False Then LDatos.Add(d)
    '        d = Nothing
    '    Next
    '    LDatos.Sort()
    '    ' CLast
    '    Dim last() As String = cIni.IniGetSection(_IniFull, "LAST")
    '    CLast = New Dictionary(Of String, String)
    '    For i As Integer = 0 To UBound(last) - 1 Step 2
    '        Dim Clave As String = last(i)
    '        Dim Valor As String = last(i + 1)   ' Es una fecha 20191004
    '        If Valor.Contains(".") Then         ' Si fuese un nº de versión 2019.0.0.35
    '            Dim partes As String() = Valor.Split("."c)
    '            For x As Integer = 0 To partes.Count - 1
    '                If partes(x).Length < 2 Then
    '                    partes(x) = partes(x).PadLeft(2, "0")
    '                End If
    '            Next
    '            Valor = String.Join("", partes)
    '        End If
    '        If CLast.ContainsKey(Clave) = False Then
    '            CLast.Add(Clave, Valor)
    '        Else
    '            CLast(Clave) = Valor
    '        End If
    '    Next
    '    Return resultado
    'End Function
    Public Function CompruebaConexionFTPUlma(Optional SubirBorrar As Boolean = False, Optional creanuevo As Boolean = False, Optional saliendo As Boolean = False) As Boolean
        Dim resultado As Boolean = False
        ' Crear objeto cFtp y comprobar conexiones.
        ' Primero a FTP1 y, si no hay conexión, después a FTP2
        Dim aa As String = ""
        'cFtp = New clsFTP(FTP1_host, FTP1_dirData, FTP1_user, FTP1_pass)
        cFtp = New clsFTP(FTP1_host, FTP1_dirLog, FTP1_user, FTP1_userpass)
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
            If IO.Directory.Exists(_appLogCSVFolder) Then
                ' Subir los ficheros al FTP
                For Each fiCsv As String In IO.Directory.GetFiles(_appLogCSVFolder, "*.csv", IO.SearchOption.TopDirectoryOnly)
                    ' No subimos ni borramos el fichero actual de LOG.
                    If _appLogCSVFichero <> "" And fiCsv = _appLogCSVFichero And saliendo = False Then
                        Continue For
                    End If
                    '
                    Dim nombre As String = IO.Path.GetFileName(fiCsv)
                    Dim fiFtp As String = FTP1_dirLog & nombre
                    Dim resultadoTxt As String = cFtp.FTP_Upload(fiCsv, fiFtp).ToString
                    If resultadoTxt.StartsWith("CORRECTO") Then
                        IO.File.Delete(fiCsv)
                    End If
                Next
            End If
        End If
        If creanuevo = True Then
            Try
                IO.File.AppendAllText(_appLogCSVFichero, Join(campos, ";"c), System.Text.Encoding.UTF8)
            Catch ex As Exception
                Process_Close("excel")
                Try
                    IO.File.AppendAllText(_appLogCSVFichero, Join(campos, ";"c), System.Text.Encoding.UTF8)
                Catch ex1 As Exception
                End Try
            End Try
        End If
        '
        ' Comprobar los ficheros log Base. Borrar los que sean de más de 10 días.
        If IO.Directory.Exists(_appLogBaseFolder) Then
            For Each fiLog As String In IO.Directory.GetFiles(_appLogBaseFolder, "*.log", IO.SearchOption.TopDirectoryOnly)
                Dim finf As New IO.FileInfo(fiLog)
                ' No borrar el fichero actual de LOG.
                If fiLog = _appLogBaseFichero Then
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
                          Optional UPDATE_GROUP As String = "",
                          Optional UPDATE_FILES As String = "",
                          Optional TYPE As String = tipo,
                          Optional KEYCODE As String = "",
                          Optional NOTES As String = "",
                          Optional EApp As ULMALGFree.queApp = ULMALGFree.queApp.ULMASTUDIO)
        '
        If yo Is Nothing Then yo = New clsBase(Reflection.Assembly.GetExecutingAssembly)
        Dim _Market As String = ""
        Dim _Language As String = ""
        If MARKET IsNot Nothing AndAlso MARKET.Count > 0 Then _Market = Join(MARKET, "|")
        If LANGUAGE IsNot Nothing AndAlso LANGUAGE.Count > 0 Then _Language = Join(LANGUAGE, "|")
        If NOTES Is Nothing Then NOTES = ""
        If NOTES.Contains(vbCrLf) Then NOTES = NOTES.Replace(vbCrLf, ". ")
        If NOTES.Contains(";") Then NOTES = NOTES.Replace(";", "|")
        If NOTES.Contains("_") Then NOTES = NOTES.Replace("_", " ")
        If NOTES <> "" AndAlso IsDate(NOTES) Then
            NOTES = "=" & comillas & CDate(NOTES).ToString(formatofecha) & comillas
        End If
        If KEYCODE = "" AndAlso RespID.id <> "" Then
            KEYCODE = RespID.id
        End If
        If KEYCODE = "" AndAlso RespID.id <> "" Then
            KEYCODE = RespID.id
        End If
        ULMALGFree.clsBase._ultimaAccion = ACTION
        '
        'Private campos() As String = {
        '"ACTION", "FILENAME", "FAMILY", "MARKET", "LANGUAGE", "DATE.YEAR", "DATE.MONTH", "DATE.DAY", "TIME", 
        '"COMPUTER_DOMAIN", "COMPUTER_NAME", "INTERNAL_IP", "EXTERNAL_IP", 
        '"USER_DOMAIN", "USER_NAME", "REVIT_VERSION", "UCREVIT_VERSION","UPDATE_GROUP", "UPDATE_FILES", "TYPE", "KEYCODE", "NOTES"}
        Dim ahora As String = Date.Now.ToString(formatofecha)
        ' Registrar acciones de UCBrowser o el resto ULMAStudio
        Dim valores As String() = {
            ACTION.ToUpper, FILENAME, FAMILY, _Market, _Language, Date.Now.Year, Date.Now.Month, Date.Now.Day, Date.Now.ToString("HH:mm").Split(" "c)(0),
            _appComputerDomain, _appComputer, "IP: " & _appIPPrivate, "IP: " & _appIPPublic,
            _appUserDomain, _appUser, oVersion.RevitVersionLogText,
            IIf(EApp = queApp.UCBROWSER, _UCBrowserVersion, _ULMAStudioVersion),
            "=" & comillas & ahora & comillas,
            UPDATE_GROUP, UPDATE_FILES, TYPE.ToUpper, KEYCODE, NOTES
        }
        '
        ' ***** Comprobar antes que exista el directorio y el fichero
        If IO.Directory.Exists(_appLogCSVFolder) = False Then
            Try
                IO.Directory.CreateDirectory(_appLogCSVFolder)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        If IO.File.Exists(_appLogCSVFichero) = False Then
            Try
                Call CompruebaConexionFTPUlma(False, creanuevo:=True)
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        ' ******************************************************
        Try
            IO.File.AppendAllText(_appLogCSVFichero, vbCrLf & Join(valores, ";"c), System.Text.Encoding.UTF8)
        Catch ex As Exception
            Process_Close("excel")
            Try
                IO.File.AppendAllText(_appLogCSVFichero, vbCrLf & Join(valores, ";"c), System.Text.Encoding.UTF8)
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
                          Optional UPDATE_GROUP As String = "",
                          Optional UPDATE_FILES As String = "",
                          Optional TYPE As String = tipo,
                          Optional KEYCODE As String = "",
                          Optional NOTES As String = "",
                          Optional EApp As ULMALGFree.queApp = ULMALGFree.queApp.ULMASTUDIO)
        '
        PonLog_ULMA(ACTION.ToString, FILENAME, FAMILY, MARKET, LANGUAGE, UPDATE_GROUP, UPDATE_FILES, TYPE, KEYCODE, NOTES, EApp)
    End Sub

#Region "FUNCIONES FECHA"
    Friend Function DateTime_DameFormateado(oDateTime As DateTime, Optional txtFormato As String = "yyyyMMddTHHmmss") As String
        ' Original      =   18/12/2018 13:50:43
        ' Formateado    =   20181218T135043         Con ToString("yyyyMMddTHHmmss")
        Return oDateTime.ToString(txtFormato)
    End Function
    '
    Friend Sub Logs_BaseYCSV_PonNombres()
        Dim solodatetime As String = DateTime_DameFormateado(DateTime.Now)
        Dim nombreFi As String = "logR_" &
            solodatetime & "_" &
            Environment.UserDomainName & "_" &
            Environment.MachineName &
            ".csv"
        '
        If _appLogCSVFichero = "" Then _appLogCSVFichero = IO.Path.Combine(_appLogCSVFolder, nombreFi)
        If _appLogBaseFichero = "" Then _appLogBaseFichero = IO.Path.Combine(ULMALGFree.clsBase._appLogBaseFolder, solodatetime & ".log")
    End Sub
    '
    Public Function LogCsvUlma_DameUltimo() As String
        Dim resultado As String = ""
        Dim filesCSV As String() = IO.Directory.GetFiles(_appLogBaseFolder, "*.csv")
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

    Public Shared Function App_DameVersion(Optional queDato As PROCESSVERSION = PROCESSVERSION.VERSION_TEXTO) As String
        Dim resultado As String = ""
        Select Case queDato
            Case PROCESSVERSION.VERSION_SUBVERSION
                If evAppC IsNot Nothing Then
                    resultado = evAppC.SubVersionNumber
                Else
                    resultado = Process_DameVersion(, PROCESSVERSION.VERSION_SUBVERSION)
                End If
            Case PROCESSVERSION.VERSION_TEXTO
                If evAppC IsNot Nothing Then
                    resultado = "Revit " & evAppC.SubVersionNumber & " (" & evAppC.VersionBuild & ")"
                Else
                    resultado = Process_DameVersion(, PROCESSVERSION.VERSION_TEXTO)
                End If
            Case PROCESSVERSION.VERSION_YEAR
                If evAppC IsNot Nothing Then
                    resultado = evAppC.VersionNumber
                Else
                    resultado = Process_DameVersion(, PROCESSVERSION.VERSION_YEAR)
                End If
            Case PROCESSVERSION.VERSION_BUILD
                If evAppC IsNot Nothing Then
                    resultado = evAppC.VersionBuild
                Else
                    resultado = Process_DameVersion(, PROCESSVERSION.VERSION_BUILD)
                End If
        End Select
        Return resultado
    End Function
    Public Shared Function Process_DameVersion(Optional queProcess As String = "REVIT", Optional queDato As PROCESSVERSION = PROCESSVERSION.VERSION_TEXTO) As String
        Dim resultado As String = ""
        Dim proc As Process = Nothing
        Dim procS As Process() = Process.GetProcessesByName(queProcess)
        If procS Is Nothing OrElse procS.Count = 0 Then
            resultado = ""
        Else
            proc = procS.FirstOrDefault
            Dim oPInf As System.Diagnostics.ProcessStartInfo = proc.StartInfo
            Dim productversion As String = proc.MainModule.FileVersionInfo.ProductVersion
            Dim fileversion As String = proc.MainModule.FileVersionInfo.FileVersion
            Dim partes() As String = fileversion.Split("."c)
            Select Case queDato
                Case PROCESSVERSION.VERSION_SUBVERSION
                    'resultado = productversion
                    resultado = "20" & partes(0) & "." & partes(1) & "." & partes(2)
                Case PROCESSVERSION.VERSION_TEXTO
                    'resultado = proc.ProcessName & " (" & productversion & ")"
                    resultado = proc.ProcessName & " 20" & partes(0) & "." & partes(1) & "." & partes(2) & " (" & productversion & ")"
                Case PROCESSVERSION.VERSION_YEAR
                    If productversion <> "" AndAlso productversion.Length >= 4 Then
                        'resultado = productversion.Substring(0, 4)
                        resultado = "20" & partes(0)
                    Else
                        resultado = "2019"
                    End If
                Case PROCESSVERSION.VERSION_BUILD
                    resultado = productversion
            End Select
        End If
        Return resultado
    End Function

    Public Shared Function Process_DameFullPath(Optional queProcess As String = "REVIT") As String
        Dim resultado As String = ""
        Dim proc As Process = Nothing
        Dim procS As Process() = Process.GetProcessesByName(queProcess)
        If procS Is Nothing OrElse procS.Count = 0 Then
            resultado = ""
        Else
            proc = procS.FirstOrDefault
            resultado = proc.MainModule.FileName
        End If
        Return resultado
    End Function

    Public Enum PROCESSVERSION
        VERSION_SUBVERSION  ' 2019.2
        VERSION_TEXTO       ' Revit 2019.2 (20190225_1515(x64))
        VERSION_YEAR        ' 2019
        VERSION_BUILD       ' 20190225_1515(x64)
    End Enum
    '
    Public Function QuitaYear() As String
        Dim resultado As String = ""
        For Each ch As Char In Me._asmVersion
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
        Return quePath.Replace(ULMALGFree.clsBase._LgFullFolder, ".")
    End Function
    Public Shared Sub Process_Close(name As String)   'Optional queProceso As String = "EXCEL"
        Try
            Dim proc As System.Diagnostics.Process
            ''
            For Each proc In System.Diagnostics.Process.GetProcessesByName(name)
                Try
                    proc.Kill()
                    proc.WaitForExit(5000)
                Catch ex As Exception
                    Try
                        Debug.Print(ex.ToString)
                        proc.CloseMainWindow()
                        proc.Close()
                        proc.WaitForExit(5000)
                    Catch ex1 As Exception
                        Debug.Print(ex1.ToString)
                    End Try
                    ' No hacemos nada
                End Try
            Next
        Catch ex As Exception
            '    'MessageBox.Show("No hay instancias de " & queProceso & " en ejecución.")
        End Try
    End Sub
    Public Shared Async Sub Process_Run(quePath As String, Optional argumentos As String = "", Optional visible As Boolean = False, Optional esperar As Boolean = False)
        Dim pInf As New ProcessStartInfo(quePath)
        If argumentos <> "" Then pInf.Arguments = argumentos
        If visible = False Then pInf.WindowStyle = ProcessWindowStyle.Hidden
        Await Task.Run(Sub()
                           Dim p As Process = Process.Start(pInf)
                           If esperar Then p.WaitForExit(15000)
                       End Sub)
    End Sub

    Public Shared Function DameCaminoAbsolutoUsuarioFichero(quePathFi As String) As String
        Dim resultado As String = quePathFi
        ''
        'Dim appdata As String = Environment.GetEnvironmentVariable("APPDATA")
        'Dim userprofile As String = Environment.GetEnvironmentVariable("USERPROFILE")
        ''
        '' Solo cambiamos el valor si se cumplen estos criterios.
        '' Si no se cumplen, devolvemos el mismo valor que tenía. Para evaluarlo posteriormente.
        If quePathFi.ToUpper.StartsWith("%USERPROFILE%") Then
            resultado = quePathFi.Replace("%USERPROFILE%", ULMALGFree.clsBase._LgFullFolder)
        ElseIf quePathFi.ToUpper.StartsWith("%APPDATA%") Then
            resultado = quePathFi.Replace("%APPDATA%", ULMALGFree.clsBase._LgFullFolder)
        ElseIf quePathFi.StartsWith(".\") Then
            resultado = quePathFi.Replace(".\", ULMALGFree.clsBase._LgFullFolder & "\")
        ElseIf quePathFi.StartsWith("..\") Then
            resultado = quePathFi.Replace("..\", ULMALGFree.clsBase._LgFullFolder & "\")
        End If
        '' Otra comprobación, por seguridad
        'If IO.Directory.Exists(IO.Path.GetDirectoryName(resultado)) = False OrElse resultado = "" Then
        'resultado = IO.Path.Combine(_dirApp, resultado)
        'End If
        ''
        Return resultado
    End Function

    Public Shared Function LlenaDatosMercados() As String
        Dim intMercado As Integer = CInt(ULMALGFree.clsBase.DEFAULT_PROGRAM_MARKET)
        Dim errores As String = ""
        Try
            Dim queFicheroArt As String = IO.Path.Combine(path_offlineBDIdata, intMercado & fijoArt & ".xml")
            RellenaArticulosXML_Linq(queFicheroArt)
            System.Threading.Thread.Sleep(1)
            Dim queFicheroGroups As String = IO.Path.Combine(path_offlineBDIdata, intMercado & fijoGroups & ".xml")
            RellenaGruposXMLParaFree_Linq(queFicheroGroups)
            System.Threading.Thread.Sleep(1)
        Catch ex As Exception
            errores &= "LlenaDatosMercados --> " & ex.Message & vbCrLf & vbCrLf
        End Try
        Return errores
    End Function
    '
    Public Shared Function CheckDecimal_Value(value As String) As String
        If value.Contains(".") Or value.Contains(",") Then
            If (s = ",") Then
                value = Replace(value, ".", ",")
            ElseIf (s = ".") Then
                value = Replace(value, ",", ".")
            End If
        End If
        Return value
    End Function
End Class
#End Region
'
#Region "ENUMERACIONES"
Public Enum queApp
    ULMASTUDIO
    UCBROWSER
    ULMASTUDIOREPORT
    ULMAUPDATERADDIN
End Enum
Public Enum ACTION
    ABOUT
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
    LOAD_FAMILY
    INSERT_FAMILY
    SAVE
    SAVEAS
    SAVEAS_LIBRARY_FAMILY
    PRINT_DOCUMENT
    PRINT_VIEW
    SYNCHRONIZE_DOCUMENT
    ' UCBROWSER
    BROWSER_OPEN
    BROWSER_LOAD_FAMILY
    BROWSER_INSERT_FAMILY
    BROWSER_CHANGE_FAMILY_FOLDER
    BROWSER_NAVIGATE
    BROWSER_SEARCH
    BROWSER_CLOSE
    ' DOWNLOAD Y UPDATE
    DOWNLOAD_GROUP
    UPDATE_FAMILIES
    REMOVE_FAMILIES
    UPDATE_ADDIN
    UPDATE_XML
End Enum
Public Enum regLog
    inicio
    registro
    final
End Enum
Public Enum FOLDERWEB
    Addins
    Families
    XML
End Enum
Public Enum queAction
    toupdate        ' Tiene actualizacion.
    notupdated      ' No tenemos que mostrar sus actualizaciones.
    updated         ' Ya esta actualizada.
End Enum
Public Enum status
    DESCARGAR
    ACTUALIZAR
    YADESCARGADO
End Enum
#End Region
