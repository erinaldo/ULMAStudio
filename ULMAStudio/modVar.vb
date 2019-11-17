Imports System
Imports System.IO
'Imports ConsultarBDI
''
Module modVar
    '' ***** Control de tiempo para actualizaciones
    Public ultimacomprobacion As Long = 0
    Public oP As Process = Nothing
    Public actualizar As Boolean = False
    '' ***** Objetos Revit
    Public queTr As Transaction = Nothing
    Public queTrSub As SubTransaction = Nothing
    Public activeSketchPlane3D As SketchPlane = Nothing
    Public activeViewOrientation3D As ViewOrientation3D = Nothing
    Public activeOptId As ElementId = ElementId.InvalidElementId
    '
    ' ***** Ribbon, Panel y botones
    ' RibbonPanels
    Public panelTools As RibbonPanel
    Public panelToolsW As Autodesk.Windows.RibbonPanel
    Public panelAbout As RibbonPanel
    Public panelAboutW As Autodesk.Windows.RibbonPanel
    ' RibbonItems (RibbonButtons)
    Public btnBrowserBoton As PushButton     ' Browser Juan Murua
    Public btnReportBoton As PushButton
    Public btnDownloadBoton As PushButton        ' Una imagen con un Nº por cada actualización disponible.
    Public btnAboutBoton As PushButton
    '
    ' Nombres Paneles UCRevitFree
    Public Const nombrePanelTools As String = "Tools"
    Public Const nombrePanelAbout As String = "About"
    Public arrPaneles() As String = New String() {nombrePanelTools, nombrePanelAbout}
    '
    ' ***** FORMULARIOS ************
    Public frmA As frmAbout = Nothing
    Public fWait As frmGeneratingReport = Nothing
    '
    ' ***** CLASES ******************
    Public cIni As clsINI
    Public cLcsv As ULMALGFree.clsBase
    '
    ' ***** Colecciones *************
    Public colPaneles As ArrayList          '' Coleccion de paneles a cargar en UTILIDADES
    Public colBotones As ArrayList          '' Coleccion de botones a cargar en UTILIDADES
    'Public colIdiomas As SortedList         '' Coleccion de idiomas para codigos/descripciones (key=idioma, Value=Pais)
    'Public arrIdiomas As ArrayList          '' Array de idiomas
    'Public colMercadosDes As SortedList        '' Coleccion de Mercados (key=descripcion, value=codigo,descripcion,idioma,pais)
    'Public colMercadosCod As SortedList        '' Coleccion de Mercados (key=codigo, value=codigo,descripcion,idioma,pais)
    'Public arrMercadosDes As ArrayList         '' array de Mercados (Con descripcion)
    'Public arrMercadosCod As ArrayList         '' array de Mercados (Con codigo)
    Public arrM() As String = New String() {"120"}  ' Mercado M1
    Public arrL() As String = New String() {"en"}   ' Lenguaje L1
    Public L1Viejo As String = ""           '' Lenguaje 1 que estaba en configuración de Revit o en fichero .ini
    Public L1Nuevo As String = ""           '' Lenguaje 1 que ponemos en opciones (Comparar con L1Viejo para ver si ha cambiado)
    Public M1Viejo As String = ""           ' Mercado 1 que estaba en configuración de Revit.
    Public M1Nuevo As String = ""           ' Mercado 1 que ponemos en opciones
    Public colFilas As Hashtable            '' Hashtable de filas BOM-WK (key=nº fila, value=cFam)
    Public bom_union As List(Of String)      '' Las tablas de planificación a unir y exportar (1 cabecera común)
    Public bom_todas As List(Of String)     '' Tablas de planificación del desarrollo (Las parciales y #BOM-WK)
    Public bom_union_columns As List(Of String)       '' Las cabeceras que tendrá esta tabla unida
    Public beta As List(Of String)                      '' PCs/Dominio/users para ver botones de funciones beta/pruebas
    Public explode_parameters As List(Of String)        '' Colección de parámetros que hay que volver a pasar de FamOri a FamFin
    Public oIdsFamI As List(Of ElementId)
    Public oIdsAnoI As List(Of ElementId)
    Public colIdiomaCod As Dictionary(Of String, String) = Nothing
    Public colCodIdioma As Dictionary(Of String, String) = Nothing
    Public arrTemplates As String() = New String() {"UC_FAMILY_TEMPLATE.rft", "UC_SET_TEMPLATE.rft", "UC_SPECIAL_FAMILY_TEMPLATE.rft", "ULMA TEMPLATE BASIC.rte", "ULMA TEMPLATE BRIO.rte"}
    Public colFicheros As Dictionary(Of String, String)     '' Ficheros que abrimos y codificamos (Key=En Temp. Sin codificar,   Value=fullPath Real, codificado)
    Public colDejar As List(Of String)
    Public colFamCode As Hashtable            '' "[M1]_revit_dynamicFamilies_dynamicfamilies.xml leidos del XML del mercado (Key=Market&FAMILY_CODE&ITEM_LENGTH&ITEM_WIDTH&ITEM_HEIGHT, Value=clsFamCode) "RellenaFamCodeXML_Linq"
    Public dominios() As String = New String() {"ulma", "ada"}
    Public arrCompanies As String() = {"120", "ULMA C Y E S.COOP.", "es", "ES"}
    Public arrLanguages As String() = {"en", "English"}
    ''
    '' ***** Variables de la aplicación y caminos iniciales
    Public Const fabricante As String = "ULMA"
    Public _dirApp As String = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly.Location)      ' Directorio de instalación de la DLL.
    Public _dirApphelp As String = IO.Path.Combine(_dirApp, "help")            ' Directorio _dirApp & "\help"
    Public _dirAppimages As String = IO.Path.Combine(_dirApp, "images")        ' Directorio _dirApp & "\images"
    'Public _dirApplogs As String = IO.Path.Combine(_dirApp, "logs")            ' Directorio _dirApp & "\logs"
    Public _dirApptemplates As String = IO.Path.Combine(_dirApp, "templates")  ' Directorio _dirApp & "\templates"
    Public _fileIni As String = IO.Path.Combine(_dirApp, My.Application.Info.AssemblyName & ".ini")
    Public _CODIGOS_IDIOMAS As String = IO.Path.Combine(_dirApp, "COD_LANGUAGES.txt")
    Public _reportExe As String = IO.Path.Combine(_dirApp, "ULMAStudioReport.exe")
    'ULMAUpdaterAddin.exe REVIT True C:\Temp\Actualizaciones\Update1.zip %appdata%\Autodesk\Revit\Addins\2019\UCRevitF\
    '
    Public _introLabPath As String = Reflection.Assembly.GetExecutingAssembly.Location      ' Ubicación de UCRevitFree.dll
    Public _introLabName As String = IO.Path.GetFileNameWithoutExtension(_introLabPath)
    Public _introLabPathUCB As String = Path.Combine(_dirApp, "UCBrowser\UCBrowser.dll")     ' Ubicación de UCBrowser.dll
    Public _introLabNameUCB As String = IO.Path.GetFileNameWithoutExtension(_introLabPathUCB)
    '
    Public dateXML As String = ""
    ' ************************************************************
    '
    ' ***** Variables fichero .INI
    '[OPTIONS]
    Public _uiLabName As String = "ULMA Studio"          ' Nombre Ribbon "ULMAF"
    Public _uiLabNameShort As String = "ULMAStudio"
    Public _uiLabNameULMA As String = "ULMA"          ' Nombre Ribbon "ULMA"
    Public loadaddin As Boolean = False
    Public anularWARNINGS As Boolean = False         '' anularWarnings = True los anula
    Public SaveImagePreview As Boolean = False
    Public log As Boolean = True
    Public includewithoutcode As Boolean = True
    Public DEFAULT_PROGRAM_LANGUAGE As String = ""
    Public DEFAULT_PROGRAM_MARKET As String = ""
    Public DEFAULT_USER_CONFIG As String = ""    '' Configuración por defecto del .ini (Parametro DEFAULT_USER_CONFIG)
    Public contact As String = "bim@ulmaconstruction.com"               '' email de contacto de ULMA.
    Public checkupdatesinseconds As Double = 60
    Public onlyulma As Boolean = False
    Public onlyactiveview As Boolean = False
    Public onlywithimage As Boolean = False
    '
    '[PATHS]
    Public path_offlineBDIdata As String = ".\offlineBDIdata"
    Public path_families_base As String = ".\families"
    Public path_families_base_images As String = path_families_base & "_images"
    Public path_families_custom As String = ""
    Public path_families_custom_images As String = ""
    ' ****************************************************************
    '
    '' ***** Variables Generales
    Public bom_base As String = "#BOM-WK"
    Public enejecucion As Boolean = False
    Public ITEM_MARKET As Double
    Public CerrandoRevit As Boolean = False
    Public registraLog As Boolean = True
    Public asBorradoOtrosRibbonItem As Boolean = False
    Public actua As Boolean = True
    Public yacodificado As Boolean = False      '' Si ya lo hemos codificado.
    '
    ' ***** CONSTANTES
    Public Const RevitVersion As String = "2018"                ' Versión de Revit en la que va el desarrollo
    'Public Const ficherosXML = "offlineBDIdata"
    Public Const tab As Char = Chr(9)
    Public Const comi As String = """" '& Chr(34)
    '' Tipos actuales
    Public Const typeAnnotationSymbol As Autodesk.Revit.DB.BuiltInCategory = BuiltInCategory.OST_GenericAnnotation
    Public Const typeFamily As Autodesk.Revit.DB.BuiltInCategory = BuiltInCategory.OST_GenericModel 'BuiltInCategory.OST_Site
    Public Const typeSet As Autodesk.Revit.DB.BuiltInCategory = BuiltInCategory.OST_GenericModel
    Public Const typeNotes As Autodesk.Revit.DB.BuiltInCategory = BuiltInCategory.OST_GenericAnnotation
    Public Const conmensajes As Boolean = True         ' Si se mostrarán determinados mensajes de aviso
    '
    Public Function INICargar() As String()
        If cIni Is Nothing Then cIni = New clsINI
        Dim mensaje(1) As String
        ' Mensaje(0) contendrá los errores.
        ' Mensaje(1) contendrá los valores leidos del .INI
        mensaje(0) = "" : mensaje(1) = ""
        '
        '[OPTIONS]
        '_uiLabName=ULMAStudio
        'loadaddin = 1
        'anularWARNINGS = 0
        'log = 1
        'DEFAULT_PROGRAM_MARKET = 120
        'DEFAULT_PROGRAM_LANGUAGE = en
        'contact = revit@ulmaconstruction.es
        'checkupdatesinseconds = 60
        'onlyulma = 1
        'onlywithimage = 1
        ';
        '[PATHS]
        'path_offlineBDIdata = .\offlineBDIdata
        'path_families = .\UCBrowser\families
        ';path_families_images =.\UCBrowser\families_images
        '
        ' _uiLabName
        _uiLabName = cIni.IniGet(_fileIni, "OPTIONS", "_uiLabName")
        If _uiLabName = "" Then _uiLabName = "ULMA Studio"
        _uiLabNameShort = _uiLabName.Replace(" ", "").Trim
        mensaje(1) &= "_uiLabName = " & _uiLabName & vbCrLf
        mensaje(1) &= "_uiLabNameShort = " & _uiLabNameShort & vbCrLf
        '
        ' loadaddin
        Dim loadaddinTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "loadaddin")
        If loadaddinTemp = "1" Then loadaddin = True
        '******************************************************************************
        ' TEMPORAL. Siempre activo, hasta que validen en ULMA.
        loadaddin = True
        '******************************************************************************
        mensaje(1) &= "loadaddin = " & loadaddin.ToString & vbCrLf
        '
        ' anularWARNINGS
        Dim anularWARNINGSTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "anularWARNINGS")
        If anularWARNINGSTemp = "1" Then anularWARNINGS = True
        mensaje(1) &= "anularWARNINGS = " & anularWARNINGS.ToString & vbCrLf
        '
        ' log
        'Dim logTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "log")
        'If logTemp <> "1" Then log = False
        '********** Log siempre activado (Lo quitamos del .ini)
        log = True
        ' ******************************
        mensaje(1) &= "log = " & log.ToString & vbCrLf
        '
        ' includewithoutcode
        Dim includewithoutcodeTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "includewithoutcode")
        If includewithoutcodeTemp <> "1" Then includewithoutcode = False
        mensaje(1) &= "includewithoutcodeTemp = " & includewithoutcodeTemp.ToString & vbCrLf
        '
        ' DEFAULT_PROGRAM_MARKET
        DEFAULT_PROGRAM_MARKET = cIni.IniGet(_fileIni, "OPTIONS", "DEFAULT_PROGRAM_MARKET")
        arrM(0) = DEFAULT_PROGRAM_MARKET
        ULMALGFree.clsBase.DEFAULT_PROGRAM_MARKET = DEFAULT_PROGRAM_MARKET
        mensaje(1) &= "DEFAULT_PROGRAM_MARKET = " & DEFAULT_PROGRAM_MARKET & vbCrLf
        '
        ' DEFAULT_PROGRAM_LANGUAGE
        DEFAULT_PROGRAM_LANGUAGE = cIni.IniGet(_fileIni, "OPTIONS", "DEFAULT_PROGRAM_LANGUAGE")
        arrL(0) = DEFAULT_PROGRAM_LANGUAGE
        ULMALGFree.clsBase.DEFAULT_PROGRAM_LANGUAGE = DEFAULT_PROGRAM_LANGUAGE
        mensaje(1) &= "DEFAULT_PROGRAM_LANGUAGE = " & DEFAULT_PROGRAM_LANGUAGE & vbCrLf
        '
        ' contact
        contact = cIni.IniGet(_fileIni, "OPTIONS", "contact")
        '
        ' updateseconds
        Dim checkupdatesinsecondsTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "checkupdatesinseconds")
        If checkupdatesinsecondsTemp <> "" AndAlso IsNumeric(checkupdatesinsecondsTemp) Then
            checkupdatesinseconds = Convert.ToDouble(checkupdatesinsecondsTemp)
        Else
            checkupdatesinseconds = 300
        End If
        '
        ' onlyulma
        Dim onlyulmaTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "onlyulma")
        If onlyulmaTemp = "1" Then onlyulma = True
        '
        ' onlyulma
        Dim onlyactiveviewTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "onlyactiveview")
        If onlyactiveviewTemp = "1" Then onlyactiveview = True
        '
        ' onlywithimage
        Dim onlywithimageTemp As String = cIni.IniGet(_fileIni, "OPTIONS", "onlywithimage")
        If onlywithimageTemp = "1" Then ULMALGFree.clsBase.onlywithimage = True

        ' [PATHS] CARGAR DIRECTORIOS Y PLANTILLAS
        '
        ' path_offlineBDIdata
        path_offlineBDIdata = cIni.IniGet(_fileIni, "PATHS", "path_offlineBDIdata")
        path_offlineBDIdata = DameCaminoAbsolutoUsuarioFichero(path_offlineBDIdata)
        If IO.Directory.Exists(path_offlineBDIdata) Then
            mensaje(1) &= "path_offlineBDIdata = " & path_offlineBDIdata & vbCrLf
        Else
            mensaje(0) &= "The directory (path_offlineBDIdata) not exist... --> " & path_offlineBDIdata & vbCrLf
        End If
        ULMALGFree.clsBase.path_offlineBDIdata = path_offlineBDIdata
        '
        ' Directorio Librerias/Imagenes Delegation
        path_families_base = cIni.IniGet(_fileIni, "PATHS", "path_families_base")
        path_families_base = DameCaminoAbsolutoUsuarioFichero(path_families_base)
        If IO.Directory.Exists(path_families_base) Then
            mensaje(1) &= "path_families_base = " & path_families_base & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families) not exist... --> " & path_families_base & vbCrLf
        End If
        path_families_base_images = cIni.IniGet(_fileIni, "PATHS", "path_families_base_images")
        path_families_base_images = DameCaminoAbsolutoUsuarioFichero(path_families_base_images)
        If IO.Directory.Exists(path_families_base_images) Then
            mensaje(1) &= "path_families_base_images = " & path_families_base_images & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_base_images) not exist... --> " & path_families_base_images & vbCrLf
        End If
        ULMALGFree.clsBase.path_families_base = path_families_base
        ULMALGFree.clsBase.path_families_base_images = path_families_base_images
        '
        '
        ' Directorio Librerias/Imagenes CUSTOM
        path_families_custom = cIni.IniGet(_fileIni, "PATHS", "path_families_custom")
        path_families_custom = DameCaminoAbsolutoUsuarioFichero(path_families_custom)
        If IO.Directory.Exists(path_families_base) Then
            mensaje(1) &= "path_families_base = " & path_families_base & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families) not exist... --> " & path_families_base & vbCrLf
        End If
        path_families_custom_images = cIni.IniGet(_fileIni, "PATHS", "path_families_custom_images")
        path_families_custom_images = DameCaminoAbsolutoUsuarioFichero(path_families_custom_images)
        If IO.Directory.Exists(path_families_base_images) Then
            mensaje(1) &= "path_families_base_images = " & path_families_base_images & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_base_images) not exist... --> " & path_families_base_images & vbCrLf
        End If
        ULMALGFree.clsBase.path_families_custom = path_families_custom
        ULMALGFree.clsBase.path_families_custom_images = path_families_custom_images
        '
        ' Coger fecha del más reciente de los XML que hay en offlineBDIdata
        Dim lastDate As New Date(2000, 12, 31)
        For Each fiXML As String In IO.Directory.GetFiles(path_offlineBDIdata, "*.xml")
            Dim oFinf As New IO.FileInfo(fiXML)
            If oFinf.LastWriteTime > lastDate Then
                lastDate = oFinf.LastWriteTime
            End If
        Next
        dateXML = lastDate.ToShortDateString
        '
        Return mensaje
    End Function
    ''
    Public Function DameCaminoAbsolutoUsuarioFichero(quePathFi As String) As String
        Dim resultado As String = quePathFi
        ''
        'Dim appdata As String = Environment.GetEnvironmentVariable("APPDATA")
        'Dim userprofile As String = Environment.GetEnvironmentVariable("USERPROFILE")
        ''
        '' Solo cambiamos el valor si se cumplen estos criterios.
        '' Si no se cumplen, devolvemos el mismo valor que tenía. Para evaluarlo posteriormente.
        If quePathFi.ToUpper.StartsWith("%USERPROFILE%") Then
            resultado = quePathFi.Replace("%USERPROFILE%", _dirApp)
        ElseIf quePathFi.ToUpper.StartsWith("%APPDATA%") Then
            resultado = quePathFi.Replace("%APPDATA%", _dirApp)
        ElseIf quePathFi.StartsWith(".\") Then
            resultado = quePathFi.Replace(".\", _dirApp & "\")
        ElseIf quePathFi.StartsWith("..\") Then
            resultado = quePathFi.Replace("..\", _dirApp & "\")
        End If
        '' Otra comprobación, por seguridad
        'If IO.Directory.Exists(IO.Path.GetDirectoryName(resultado)) = False OrElse resultado = "" Then
        'resultado = IO.Path.Combine(_dirApp, resultado)
        'End If
        ''
        Return resultado
    End Function
    ''
    Public Function FormDameControl(ByRef queform As System.Windows.Forms.Form, queNombre As String) As System.Windows.Forms.Control
        Dim resultado As System.Windows.Forms.Control = Nothing
        ''
        Dim encontrado As Boolean = False
        For Each queC As System.Windows.Forms.Control In queform.Controls
            If queC.Name = queNombre Then
                resultado = queC
                Exit For
            End If
            ''
            For Each queC1 As System.Windows.Forms.Control In queC.Controls
                If queC1.Name = queNombre Then
                    resultado = queC1
                    encontrado = True
                    Exit For
                End If
            Next
            If encontrado = True Then Exit For
        Next
        ''
        Return resultado
    End Function
    '
    Public Function Rellenaconfig_CambiaM0L0Solo(Optional queM1 As String = "", Optional queL1 As String = "") As String
        '' Cargamos o la configuración de Markets y Languages del .ini o de PROJECT_CONFIG
        Dim queConf As String = modULMA.ParametroProyectoLee(evRevit.evAppUI.ActiveUIDocument.Document, "PROJECT_CONFIG")
        If queConf = "" Then
            queConf = DEFAULT_USER_CONFIG
        End If
        Dim resultado As String = queConf
        '
        If queM1 = "" AndAlso queL1 = "" Then
            Return resultado
            Exit Function
        End If
        '
        If queConf <> "" AndAlso
            queConf.Contains("M1=") AndAlso
            queConf.Contains("M2=") AndAlso
            queConf.Contains("M3=") AndAlso
            queConf.Contains("L1=") AndAlso
            queConf.Contains("L2=") AndAlso
            queConf.Contains("L3=") AndAlso
            queConf.Contains("|") Then
            Dim partes0() As String = queConf.Split("|"c)   '0=Markets | 1=Languages
            Dim Mark() As String = partes0(0).ToString.Split(","c)
            Dim Lang() As String = partes0(1).ToString.Split(","c)
            ''
            If queM1 <> "" AndAlso Mark(0).ToString.ToUpper.Split("="c)(1).Trim <> queM1.ToUpper Then
                Mark(0) = Mark(0).ToString.Split("="c)(0).Trim & "=" & queM1
                arrM(0) = queM1
                ITEM_MARKET = CInt(queM1)
            End If
            '
            If queL1 <> "" AndAlso Lang(0).ToString.ToUpper.Split("="c)(1).Trim <> queL1.ToUpper Then
                Lang(0) = Lang(0).ToString.Split("="c)(0).Trim & "=" & queL1
                arrL(0) = queL1
            End If
            resultado = String.Join(",", Mark) & "|" & String.Join(",", Lang)
        End If
        '
        Return resultado
    End Function
    '
    Public Sub BotonesAplicacionEstado(activar As Boolean)
        If btnBrowserBoton IsNot Nothing Then btnBrowserBoton.Enabled = activar
        If btnReportBoton IsNot Nothing Then btnReportBoton.Enabled = activar
        If btnDownloadBoton IsNot Nothing Then btnDownloadBoton.Enabled = activar
        If btnAboutBoton IsNot Nothing Then btnAboutBoton.Enabled = True
    End Sub
    Public Function FormulariosDesarrolloAbiertos() As Boolean
        '' ***** FORMULARIOS
        If frmA IsNot Nothing OrElse
            fWait IsNot Nothing Then
            Return True
        Else
            Return False
        End If
    End Function
    '
    Public Function DameTipo(eTipo As DocumentType, Optional extension As String = "") As String
        Dim dTipo As String = "_"
        Dim template As String = ""
        '
        If eTipo = DocumentType.Project Then
            dTipo &= eTipo.ToString.ToUpper
        ElseIf eTipo = DocumentType.Family Then
            dTipo &= eTipo.ToString.ToUpper
        ElseIf eTipo = DocumentType.Template Then
            If extension <> "" Then
                If extension.ToUpper.Contains("RFT") OrElse extension.ToUpper.Contains("RTE") Then
                    template = "_TEMPLATE"
                End If
            End If
        ElseIf eTipo = DocumentType.IFC Then
            dTipo &= eTipo.ToString.ToUpper
        ElseIf eTipo = DocumentType.BuildingComponent Then
            dTipo &= eTipo.ToString.ToUpper
        ElseIf eTipo = DocumentType.Other Then
            dTipo &= eTipo.ToString.ToUpper
        End If

        '
        Return dTipo & template
    End Function
    'Public Function DameTipo(oD As Document, extension As String) As String
    '    Dim ext As String = IO.Path.GetExtension(oD.PathName)
    '    Dim eTipo As DocumentType = Nothing
    '    If ext.ToUpper.Contains("RVT") Then
    '        eTipo = DocumentType.Project
    '    ElseIf ext.ToUpper.Contains("RFA") Then
    '        eTipo = DocumentType.Family
    '    ElseIf ext.ToUpper.Contains("RTE") Then
    '        eTipo = DocumentType.Template
    '    ElseIf ext.ToUpper.Contains("RFT") Then
    '        eTipo = DocumentType.Template
    '    ElseIf ext.ToUpper.Contains("IFC") Then
    '        eTipo = DocumentType.IFC
    '    ElseIf ext.ToUpper.Contains("ADSK") Then
    '        eTipo = DocumentType.BuildingComponent
    '    Else
    '        eTipo = DocumentType.Other
    '    End If
    '    '
    '    Return DameTipo(eTipo, extension)
    'End Function
End Module
