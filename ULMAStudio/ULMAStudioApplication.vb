#Region "Imported Namespaces"
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI

Imports System.Collections.Generic
Imports System.Xaml
Imports System.Diagnostics '– used for debug 
Imports System.IO '– used for reading folders 
Imports System.Windows
Imports Autodesk.Revit.DB.Events
Imports adWin = Autodesk.Windows
Imports System.Timers
Imports uf = ULMALGFree.clsBase
#End Region

Class ULMAStudioApplication
    Implements IExternalApplication
    ''' <summary>
    ''' This method is called when Revit starts up before a 
    ''' document or default template is actually loaded.
    ''' </summary>
    ''' <param name="app">An object passed to the external 
    ''' application which contains the controlled application.</param>
    ''' <returns>Return the status of the external application. 
    ''' A result of Succeeded means that the external application started successfully. 
    ''' Cancelled can be used to signify a problem. If so, Revit informs the user that 
    ''' the external application failed to load and releases the internal reference.
    ''' </returns>
    Public Function OnStartup(
      ByVal app As UIControlledApplication) _
    As Result Implements IExternalApplication.OnStartup

        'TODO: Add your code here
        evRevit.evAppUIC = app
        ULMALGFree.clsBase.evAppUIC = app
        evRevit.evAppC = app.ControlledApplication
        ULMALGFree.clsBase.evAppC = app.ControlledApplication
        uf.Version_Put()
        '
        AddHandler app.Idling, AddressOf evRevit.evAppCUI_Idling_LlenaObjectos
        'AddHandler app.Idling, AddressOf evRevit.evAppCUI_Idling
        'AddHandler Autodesk.Windows.ComponentManager.PreviewExecute, AddressOf evRevit.evCompM_PreviewExecute

        'Process.Start("mklink " &
        '              comi & IO.Path.Combine(uf._LgFullFolder, "UCBrowser", "ConsultarBDI.dll") & comi &
        '              " " &
        '              comi & IO.Path.Combine(uf._LgFullFolder, "ConsultarBDI.dll") & comi)
        '
        '
        If cIni Is Nothing Then cIni = New clsINI
        'cIni.IniDeleteSection(ULMALGFree.clsBase._IniFull, "UPDATES")
        ULMALGFree.clsBase.RenombraExeBak()
        'ULMALGFree.clsBase.Process_Close("ULMAUpdaterAddIn")
        '
        ' 1.- ***** Cargar datos del fichero .INI
        Dim resultado() As String = modVar.INICargar()
        Dim errores As String = resultado(0)
        ''PonLog("***** Iniciamos Revit *****")
        ''PonLog("***** VALORES FICHERO .INI *****" & vbCrLf & resultado(1) & vbCrLf & StrDup(50, "*"))
        If errores <> "" Then
            ''PonLog(vbCrLf & errores)
            TaskDialog.Show("ATTENTION", "Error in " & My.Application.Info.AssemblyName & ".ini" & vbCrLf & errores)
            Return Result.Cancelled
            Exit Function
        End If
        ' Crear instancia ULMALGFree.clsBase (Solo para algunas variables y funciones) El resto son Shared.
        cLcsv = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)  ', app.ControlledApplication.VersionName.Substring(0, 4))
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
        '
        ' ***** COMPROBAR SI ESTÁ AUTORIZADO Y SI TIENE QUE MOSTRAR FORMULARIO ID
        Dim msgCancel As String = "The UCRevit AddIn could not be validated. It will not load."
        ' Si NO existe el fichero key.dat. ** Siempre llamamos al Formulario
        ' Si existe el fichero key.dat. Comprobar conexión/internet, estructura y datos del fichero.

        If IO.File.Exists(uf.keyfile) = False Then
            ' No key.dat. Formulario
            Dim frmID As New ULMALGFree.frmCompruebaID
            Dim fRes As Forms.DialogResult = frmID.ShowDialog(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle))

            If fRes = Forms.DialogResult.Cancel Then
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Form Code Canceled: " & uf.RespID.message)
                Call cLcsv.CompruebaConexionFTPUlma(SubirBorrar:=True)
                Return Result.Cancelled
                Exit Function
            End If
        ElseIf IO.File.Exists(uf.keyfile) = True Then
            ' Si key.dat. Comprobaciones
            uf.ID_Comprueba_OnLine()

            If uf.RespID.valid = True Then
                If uf.RespID.message.Contains("Offline") OrElse uf.RespID.messagelog.Contains("Offline") Then
                    ' Es correcto temporalmente, ya que no tenía conexión y no podemos comprobar el ID
                    ' (Le damos hasta un maximo de 90 días sin conexión, informamos de los días sin conexión y activamos)
                    cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="OK temporal: " & uf.RespID.messagelog)
                Else
                    ' Es correcto, continuamos sin avisos, cargando el AddIn.
                    cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Check Code OK: " & uf.RespID.messagelog)
                End If
            Else
                MsgBox(uf.RespID.message, MsgBoxStyle.Critical, "Registration")
                'TaskDialog.Show("ATTENTION", msgCancel, TaskDialogCommonButtons.Close)
                cLcsv.PonLog_ULMA("CHECK CODE", KEYCODE:=uf.RespID.id, NOTES:="Check Code error: " & uf.RespID.messagelog)
                Return Result.Cancelled
                Exit Function
            End If
        End If
        '************************************************************************
        ' ***** Rellenar datos XML en un hilo aparte.
        XML_Lee()
        '
        'TaskDialog.Show("DATOS", ULMALGFree.clsBase.Articulos().Count.ToString)
        'Dim t As New System.Threading.Thread(AddressOf ULMALGFree.clsBase.CONSULTAS_TODO) : t.Start()
        'If log Then PonLog_BASICO(ULMALGFree.clsLogsCSV._appFicherologBASE, "Start REVIT")
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_REVIT,,, arrM, arrL)
        '
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("XMLs VERSION", dateXML)
        ' 4.- ***** Creamos el interface (Ribbon, Paneles y botones)
        ''PonLog("Cargar RibbonTab, RibbonPanels y RibbonButtons")
        Ribbon_ADD_ULMA(app)
        If loadaddin = False Then
            'utilesRevit.RibbonPanels_ActivarDesactivar(_uiLabName, colDejar, loadaddin)
        Else
            NombresParametrosLocalizados()
            'modULMA.CargaDatosMercadosYO()
        End If
        ultimacomprobacion = 0
        actualizar = False
        'Must return some code
        Return Result.Succeeded
    End Function

    ''' <summary>
    ''' This method is called when Revit is about to exit.
    ''' All documents are closed before this method is called.
    ''' </summary>
    ''' <param name="app">An object passed to the external 
    ''' application which contains the controlled application.</param>
    ''' <returns>Return the status of the external application. 
    ''' A result of Succeeded means that the external application successfully shutdown. 
    ''' Cancelled can be used to signify that the user cancelled the external operation 
    ''' at some point. If false is returned then the Revit user should be warned of the 
    ''' failure of the external application to shut down correctly.</returns>
    Public Function OnShutdown(
      ByVal app As UIControlledApplication) _
    As Result Implements IExternalApplication.OnShutdown

        'TODO: Add your code here
        'Must return some code
        utilesRevit.RibbonTabBorra(_uiLabName)
        If evRevit.evApp IsNot Nothing Then
            evRevit.evApp.PurgeReleasedAPIObjects()
        End If
        Try
            'evRevit.UnsubscribeAll()
        Catch ex As Exception
            ' continuar
        End Try
        Return Result.Succeeded
    End Function
    '

    Sub Ribbon_ADD_ULMA(ByVal app As UIControlledApplication)
        ' 1.- Crear el RibbonTab
        Try
            app.CreateRibbonTab(_uiLabName)
        Catch ex As Exception
            ' Por si ya existía _uiLabName = ULMAStudio
        End Try
        '
        ' 2.- Crear el RibbonPanel (Utilidades)
        panelTools = app.CreateRibbonPanel(_uiLabName, nombrePanelTools)       '' Herramientas (Explode)
        panelTools.Title = ""
        panelAbout = app.CreateRibbonPanel(_uiLabName, nombrePanelAbout)      '' Help/About (Botones Help y About)
        panelAbout.Title = ""
        '
        ' 3.- Llenar los objetos adWin.RibbonPanel
        modULMA.RibbonPanelULMALlena()
        '
        ' 4.- Añadir Botones al cada RibbonPanel
        '
        ' Botones de Tools (Browser, Report y Download)
        PushButtons_Add_BROWSER(panelTools)
        panelTools.AddSeparator()
        PushButtons_Add_REPORT(panelTools)
        panelTools.AddSeparator()
        PushButtons_Add_UPDATEFAMILIES(panelTools)
        '
        ' Botones de About (Help + About)
        PushButtons_Add_ABOUT(panelAbout)
        'panelAbout.AddSeparator()
    End Sub
    '
    Sub PushButtons_Add_BROWSER(ByVal panel As RibbonPanel)     ' Browser Juan Murua.
        Dim btnRibbonBotonData As PushButtonData
        Dim queImg As System.Windows.Media.Imaging.BitmapSource
        btnRibbonBotonData = New PushButtonData("UCBrowser", "Browser", _introLabPathUCB, "UCBrowser.UCBrowser")
        ' Add a button to the panel 
        btnBrowserBoton = CType(panel.AddItem(btnRibbonBotonData), PushButton)
        queImg = DameImagenRecurso(My.Resources.mnuBrowser32.GetHbitmap())
        btnBrowserBoton.Image = queImg
        btnBrowserBoton.LargeImage = queImg
        ' Add a tooltip
        btnBrowserBoton.ToolTip = "Family Browser ULMA"    ' & Environment.NewLine & "ULMA"
        'btnBrowserBoton.ItemText = "Browser"   ' & Environment.NewLine & "ULMA"
        btnBrowserBoton.LongDescription = "Family Browser ULMA"
        btnBrowserBoton.Enabled = True
        btnBrowserBoton.Visible = True
    End Sub
    '
    Sub PushButtons_Add_REPORT(ByVal panel As RibbonPanel)
        Dim btnRibbonBotonData As PushButtonData
        Dim queImg As System.Windows.Media.Imaging.BitmapSource
        btnRibbonBotonData = New PushButtonData("btnReport", "Report", _introLabPath, _introLabName & ".btnReport")
        ' Add a button to the panel 
        btnReportBoton = CType(panel.AddItem(btnRibbonBotonData), PushButton)
        ' Add an icon 
        queImg = DameImagenRecurso(My.Resources.mnuReport32.GetHbitmap())
        btnReportBoton.Image = queImg
        btnReportBoton.LargeImage = queImg
        ' Add a tooltip
        btnReportBoton.ToolTip = "Report ULMA"
        'btnReportBoton.ItemText = "Opciones" & Environment.NewLine & "PrefaBIM"
        'btnReportBoton.ToolTipImage = NewBitmapImage("GENPLANOToolTip.png")
        btnReportBoton.LongDescription = "Report ULMA in PDF"
        'btnReportBoton.ClassName = _introLabName + ".btnWeb2aCAD"
        'btnReportBoton.AvailabilityClassName = _introLabName
        btnReportBoton.Enabled = True    '' Hasta que pasemos por Options
        btnReportBoton.Visible = True
        ''
        'panel.AddSeparator()
        ''
    End Sub
    '
    Sub PushButtons_Add_UPDATEFAMILIES(ByVal panel As RibbonPanel)
        Dim btnRibbonBotonData As PushButtonData
        Dim queImg As System.Windows.Media.Imaging.BitmapSource
        btnRibbonBotonData = New PushButtonData("btnUpdateFamilies", "Download", _introLabPath, _introLabName & ".btnUpdateFamilies")
        ' Add a button to the panel 
        btnDownloadBoton = CType(panel.AddItem(btnRibbonBotonData), PushButton)
        ' Add an icon
        queImg = DameImagenRecurso(My.Resources.uf32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
        btnDownloadBoton.Image = queImg
        btnDownloadBoton.LargeImage = queImg
        ' Add a tooltip
        btnDownloadBoton.ToolTip = "Download Families"
        'btnDownloadBoton.ItemText = "Opciones" & Environment.NewLine & "PrefaBIM"
        'btnDownloadBoton.ToolTipImage = NewBitmapImage("GENPLANOToolTip.png")
        btnDownloadBoton.LongDescription = "If exist downloads of ULMA Families, Download Families"
        'btnDownloadBoton.ClassName = _introLabName + ".btnWeb2aCAD"
        'btnDownloadBoton.AvailabilityClassName = _introLabName
        btnDownloadBoton.Enabled = True    '' Hasta que pasemos por Options
        btnDownloadBoton.Visible = True
        'panel.AddSeparator()
    End Sub
    Sub PushButtons_Add_ABOUT(ByVal panel As RibbonPanel)
        ''
        Dim btnRibbonBotonData As PushButtonData
        Dim queImg As System.Windows.Media.Imaging.BitmapSource
        Dim sufijo As String = ""
        ''
        '' ***** BOTON LISTADOS
        btnRibbonBotonData = New PushButtonData("btnAbout", "About" + sufijo, _introLabPath, _introLabName & ".btnAbout")
        ' Add a button to the panel 
        btnAboutBoton = CType(panel.AddItem(btnRibbonBotonData), PushButton)
        ' Add an icon 
        queImg = DameImagenRecurso(My.Resources.mnuAboutU32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
        btnAboutBoton.Image = queImg
        btnAboutBoton.LargeImage = queImg
        ' Add a tooltip
        btnAboutBoton.ToolTip = "About" + sufijo
        btnAboutBoton.LongDescription = "About" + sufijo & " of " & _introLabName & " and Update AddIn"
        btnAboutBoton.Enabled = True    '' Hasta que pasemos por Options
        btnAboutBoton.Visible = True
    End Sub
    '
    '
    Public Shared Sub BotonULMAStudioEjecuta(nombrePanel As String, nombreBoton As String, Optional nombreRibbon As String = "ULMAF")
        ''
        '' ************** Ejecutar el boton Options de ULMA
        Dim comOpti As String = "CustomCtrl_%CustomCtrl_%" &
            nombreRibbon &
            "%" & nombrePanel &
            "%" & nombreBoton
        Dim queId As RevitCommandId = RevitCommandId.LookupCommandId(comOpti)
        evRevit.evAppUI.PostCommand(queId)
    End Sub

    'Public Shared Sub Botones_ActualizaEstadoActualizaciones()
    '    Dim queImg As System.Windows.Media.Imaging.BitmapSource
    '    '
    '    Call uf.INIUpdates_LeeTODO()
    '    'If uf.cUp("addins").Count > 0 OrElse uf.cUp("families").Count > 0 OrElse uf.cUp("xmls").Count > 0 Then
    '    '    ULMALGFree.clsBase.Process_Run(_ULMAUpdaterAddIn, "UPDATE REVIT", visible:=False)
    '    'End If
    '    ' BTNABOUTBOTON (Si hay actualizaciones del AddIn) 
    '    If uf.cUp("addins").Count = 0 AndAlso uf.cUp("xmls").Count = 0 Then
    '        queImg = DameImagenRecurso(My.Resources.LogoU32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
    '    Else
    '        queImg = DameImagenRecurso(My.Resources.LogoU32U.GetHbitmap())
    '    End If
    '    'If btnAboutBoton.Image.Equals(queImg) = False Then btnAboutBoton.Image = queImg
    '    'If btnAboutBoton.LargeImage.Equals(queImg) = False Then btnAboutBoton.LargeImage = queImg
    '    btnAboutBoton.Image = queImg
    '    btnAboutBoton.LargeImage = queImg

    '    ' BTNDOWNLOADBOTON (Si hay actualizaciones de las familias)
    '    If uf.cUp("families").Count = 0 Then
    '        queImg = DameImagenRecurso(My.Resources.uf32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
    '    Else
    '        Dim nombre As String = "uf32_" & uf.cUp("families").Count '& ".png"
    '        queImg = Imagen_DesdeResources(nombre & ".png")
    '        'queImg = DameImagenRecurso(My.Resources.uf32_0.GetHbitmap())
    '    End If
    '    'If btnDownloadBoton.Image.Equals(queImg) = False Then btnDownloadBoton.Image = queImg
    '    'If btnDownloadBoton.LargeImage.Equals(queImg) = False Then btnDownloadBoton.LargeImage = queImg
    '    btnDownloadBoton.Image = queImg
    '    btnDownloadBoton.LargeImage = queImg
    '    '
    '    panelToolsW.RibbonControl.UpdateLayout()
    'End Sub
    Public Shared Sub Botones_ActualizaEstadoActualizaciones(Optional esinicio As Boolean = False)
        Dim queImg As System.Windows.Media.Imaging.BitmapSource
        '
        If esinicio = False Then
            Call uf.INIUpdates_LeeTODO()
        End If
        'If uf.cUp("addins").Count > 0 OrElse uf.cUp("families").Count > 0 OrElse uf.cUp("xmls").Count > 0 Then
        '    ULMALGFree.clsBase.Process_Run(_ULMAUpdaterAddIn, "UPDATE REVIT", visible:=False)
        'End If
        ' BTNABOUTBOTON (Si hay actualizaciones del AddIn) 
        If uf.cUp("addins").Count = 0 Then  ' AndAlso uf.cUp("xmls").Count = 0 Then
            'queImg = DameImagenRecurso(My.Resources.mnuAboutU32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
            queImg = DameImagenRecurso(My.Resources.mnuAboutU32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
        Else
            queImg = DameImagenRecurso(My.Resources.mnuAboutUU32.GetHbitmap())
        End If
        'RibbonButton_ChangeImage(btnAboutBoton, CType(queImg, System.Windows.Media.Imaging.BitmapImage))
        btnAboutBoton.Image = queImg
        btnAboutBoton.LargeImage = queImg

        ' BTNDOWNLOADBOTON (Si hay actualizaciones de las familias)
        If uf.cUp("botonesRojos").Count = 0 Then
            queImg = DameImagenRecurso(My.Resources.uf32.GetHbitmap()) '' NewBitmapImage("bomUnion.png")
        Else
            Dim nombre As String = "uf32_" & uf.cUp("botonesRojos").Count '& ".png"
            queImg = MyResource.Resources.Imagen_DameIncrustada(nombre & ".png")
        End If
        If uf.cUp("descargados").Count = 0 Then
            btnBrowserBoton.Enabled = False
            btnReportBoton.Enabled = False
        Else
            btnBrowserBoton.Enabled = True
            btnReportBoton.Enabled = True
        End If
        btnDownloadBoton.Image = queImg
        btnDownloadBoton.LargeImage = queImg
        '
        If panelToolsW IsNot Nothing Then panelToolsW.RibbonControl.UpdateLayout()
    End Sub
    '
    'Sub AddPushButtonLISTARIBBONS(ByVal panel As RibbonPanel)
    '    ' Set the information about the command we will be assigning 
    '    ' to the button 
    '    Dim btnListaRibbonsBotonData As _
    '      New PushButtonData("btnListaRibbons", "LISTARIBBONS", _introLabPath, _introLabName + ".btnListaRibbons")

    '    ' Add a button to the panel 
    '    btnListaRibbonsBoton = CType(panel.AddItem(btnListaRibbonsBotonData), PushButton)
    '    ' Add an icon 
    '    ' Make sure you reference WindowsBase and PresentationCore, 
    '    ' and import System.Windows.Media.Imaging namespace.
    '    'Dim queImg As System.Windows.Media.Imaging.BitmapImage = NewBitmapImage("GENPLANO.png")
    '    Dim queimg = DameImagenRecurso(My.Resources.OpenG32.GetHbitmap())
    '    'btnWeb2aCADBoton.Image = queImg
    '    btnListaRibbonsBoton.LargeImage = queimg
    '    ' Add a tooltip
    '    btnListaRibbonsBoton.ToolTip = "Lista Ribbons"
    '    btnListaRibbonsBoton.ToolTipImage = DameImagenRecurso(My.Resources.OpenG32.GetHbitmap())    'NewBitmapImage("GENPLANOToolTip.png")
    '    btnListaRibbonsBoton.LongDescription = "Listar Ribbons Revit"
    '    'btnWeb2aCADBoton.ClassName = _introLabName + ".btnWeb2aCAD"
    '    'btnWeb2aCADData.AvailabilityClassName = _introLabName
    '    btnListaRibbonsBoton.Enabled = True ' licencia ' True
    '    btnListaRibbonsBoton.Visible = True
    'End Sub
    '
    'Public Sub OptionsEjecuta(e As Autodesk.Revit.DB.Document)
    '    ''
    '    utilesRevit.InicioAplicacionObligatorio(shared_parameter_file)
    '    '' Desactivar botones de la aplicación, hasta pasar por opciones.
    '    'BotonesAplicacionEstado(False)
    '    ''
    '    '' ************** Ejecutar el boton Options de ULMA
    '    Dim comOpti As String = "CustomCtrl_%CustomCtrl_%ULMA%Options%btnOpti"
    '    Dim queId As RevitCommandId = RevitCommandId.LookupCommandId(comOpti)
    '    evRevit.evAppUI.PostCommand(queId)
    '    '' **************
    '    '' Esto es para ejecutar un comando de Revit. De los fijos (PostableCommand)
    '    'uiapp.PostCommand(RevitCommandId.LookupCommandId(PostableCommand.Array.ToString))
    '    '
    'End Sub
    '
    'Sub PushButtons_Add_LOADFAMILY(ByVal panel As RibbonPanel)
    '    '
    '    '' ***** BOTON ID_FAMILY_LOAD (De Revit)
    '    Dim oRi As adWin.RibbonItem = modRibbons.RibbonItem_BuscaDame("Insert", "load_from_library_shr", "ID_FAMILY_LOAD")
    '    If oRi IsNot Nothing Then
    '        panelPickingW.Source.Items.Add(oRi)
    '    End If
    '    '
    'End Sub
    Public Shared Sub XML_Lee()
        ' ***** Rellenar datos XML en un hilo aparte.
        'Rellenar colArticulos
        Dim task As New System.Threading.Thread(AddressOf uf.LlenaDatosMercados) : task.Start()
    End Sub

End Class
