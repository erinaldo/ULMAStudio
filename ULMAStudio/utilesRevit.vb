#Region "Imported Namespaces"
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.DB.Architecture

Imports System.Collections.Generic
Imports System.Xaml
Imports System.Diagnostics '– used for debug 
'Imports System.Windows.Forms
Imports System.IO '– used for reading folders 
Imports System.Windows.Media.Imaging '– used for bitmap images 
Imports adWin = Autodesk.Windows
#End Region


Module utilesRevit
    ''
    '' ***** Variables generales de Revit
    'Public utAppCont As Autodesk.Revit.ApplicationServices.ControlledApplication = utAppContUI.ControlledApplication
    'Public utAppUI As UIApplication = Nothing
    'Public utApp1 As Autodesk.Revit.ApplicationServices.Application = Nothing
    Public oSpSup As SketchPlane = Nothing
    Public oSpFro As SketchPlane = Nothing
    Public oSpBase As SketchPlane = Nothing
    Public oLxProxi As Line = Nothing
    Public famProxy As FamilyInstance = Nothing
    Public utMuestraDialogos As Boolean = True
    Public ultView As String = ""
    Public amplia As Double = 5
    '
    ''
    Public OmniClassTaxonomyUser As String =
            Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile) &
            "\AppData\Roaming\Autodesk\Revit\Autodesk Revit " & RevitVersion & "\OmniClassTaxonomy.txt"
    Public RevitFolderAddInsUser As String =
            Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile) &
            "\AppData\Roaming\Autodesk\Revit\Addins\" & RevitVersion & "\"
    Public RevitIniUser As String =
            Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile) &
            "\AppData\Roaming\Autodesk\Revit\Autodesk Revit " & RevitVersion & "\Revit.ini"
    Public RevitPrincipal As String = "C:\ProgramData\Autodesk\RVT " & RevitVersion & "\UserDataCache\Revit.ini"
    ''
    Public unidLDoc As FormatOptions          '' Unidades de Longitud del documento
    Public unidADoc As FormatOptions          '' Undiades de area del documento
    Public unidDoc As Units                 '' Unidades actuales del documento (LENGTH) ** Restablecer después de leer
    Public unidDocMM As Units               '' Unidades en MM para longitudes (LENGTH 2 decimales) ** Cambiar antes de leer.
    ''
    '' ***** Colecciones
    Public arrIds As ArrayList = Nothing
    Public colGrupos As Dictionary(Of String, ICollection(Of ElementId)) = Nothing      '' String=Nombre del grupo, ElementSet=Elementos que tenía al explotar
    ''
    '' ***** Constantes
    Public Const daError As String = "*ERROR*"
    ''
    '' ***** NOMBRES DE PARAMETROS LOCALIZADOS QUE USA EL DESARROLLO
    Public nType As String = ""
    Public nSite As String = ""
    Public nManufacturer As String = ""
    Public nCategory As String = ""
    Public nCount As String = ""
    Public nFaseCrea As String = ""
    Public nDesfase As String = ""
    Public nFamily As String = ""
    '
#Region "OBJETOS APPLICATION Y DOCUMENTO ACTIVO"
    Public Function utViewUI_Dame(name As String) As UIView
        Dim resultado As UIView = Nothing
        Dim lstUIViews As IList(Of UIView) = utOpenUIViews()
        For Each oUIV As UIView In lstUIViews
            Dim oV As View = CType(evRevit.evDoc.GetElement(oUIV.ViewId), View)
            If oV.Name = name OrElse oV.Name = name Then
                resultado = oUIV
                Exit For
            End If
        Next
        '
        Return resultado
    End Function
    Public Function utOpenUIViews() As IList(Of UIView)
        Return evRevit.evAppUI.ActiveUIDocument.GetOpenUIViews
    End Function
    '
    Public Sub Zoom_Fit()
        Dim oViewUI As UIView = utViewUI_Dame(evRevit.evView.Name)
        If oViewUI IsNot Nothing Then oViewUI.ZoomToFit()
    End Sub
#End Region

    Public Sub Directorio_OcultarMostrar(queDir As String, ocultar As Boolean)
        If IO.Directory.Exists(queDir) = False Then Exit Sub
        Dim oDinf As IO.DirectoryInfo = New IO.DirectoryInfo(queDir)
        Dim attribute As System.IO.FileAttributes = Nothing
        If ocultar = True Then
            attribute = oDinf.Attributes Or IO.FileAttributes.Hidden
        Else
            attribute = IO.FileAttributes.Normal
        End If
        oDinf.Attributes = attribute
        oDinf = Nothing
    End Sub
    Public Sub Fichero_OcultarMostrar(queFic As String, ocultar As Boolean)
        If IO.File.Exists(queFic) = False Then Exit Sub
        Dim oFinf As IO.FileInfo = New IO.FileInfo(queFic)
        Dim attribute As System.IO.FileAttributes = Nothing
        If ocultar = True Then
            attribute = oFinf.Attributes Or IO.FileAttributes.Hidden
        Else
            attribute = IO.FileAttributes.Normal
        End If
        oFinf.Attributes = attribute
        oFinf = Nothing
    End Sub
    Public Sub PonLog_BASICO(quefilog As String, ByVal quetexto As String, Optional ByVal borrar As Boolean = False)
        If IO.File.Exists(quefilog) = True AndAlso borrar = True Then IO.File.Delete(quefilog)
        If quefilog = "" Then quefilog = ULMALGFree.clsBase._appLogBaseFichero
        If quetexto.EndsWith(vbCrLf) = False Then quetexto &= vbCrLf
        IO.File.AppendAllText(quefilog, Date.Now.ToString & vbTab & quetexto)
    End Sub
    Public Sub MostrarMensaje(titulo As String, mensaje As String)
        Dim td As New Autodesk.Revit.UI.TaskDialog(titulo)
        td.MainInstruction = mensaje
        td.TitleAutoPrefix = False
        td.Show()
    End Sub
    Public Sub DesignOptions_Pon(doc As Document, ByRef oEle As Element, queDoId As ElementId)
        If queDoId <> ElementId.InvalidElementId Then
            Dim doFilter As New ElementDesignOptionFilter(queDoId)
            If oEle.DesignOption Is Nothing OrElse (oEle.DesignOption IsNot Nothing AndAlso oEle.DesignOption.Id.IntegerValue <> queDoId.IntegerValue) Then
                doFilter.PassesFilter(doc, oEle.Id)
            End If
        End If
    End Sub
    '
    Public Function DesignOptions_Dame(oEle As Element) As DesignOption
        Return oEle.DesignOption
    End Function
    ''' <summary>
    ''' Starting at the given directory, search upwards for 
    ''' a subdirectory with the given target name located
    ''' in some parent directory. 
    ''' </summary>
    ''' <param name="path">Starting directory, e.g. GetDirectoryName( GetExecutingAssembly().Location ).</param>
    ''' <param name="target">Target subdirectory name, e.g. "Images".</param>
    ''' <returns>The full path of the target directory if found, else null.</returns>
    Public Function FindFolderInParents(
      ByVal path As String, ByVal target As String) As String

        Debug.Assert(Directory.Exists(path),
          "expected an existing directory to start search in")
        Do
            Dim s As String = System.IO.Path.Combine(path, target)
            If Directory.Exists(s) Then
                Return s
            End If
            path = System.IO.Path.GetDirectoryName(path)
        Loop While (path IsNot Nothing)
        Return Nothing
    End Function
    '
    Public Function FolderLibrary_BuscaDame(appUI As Autodesk.Revit.UI.UIControlledApplication) As String
        Dim resultado As String = ""
        'Dim colPaths As IDictionary(Of String, String) = appUI.ControlledApplication.GetLibraryPaths
        'For Each key As String In colPaths.Keys
        '    Dim queDir As String = colPaths(key)
        '    If IO.Directory.Exists(queDir) Then
        '        resultado = queDir
        '        Exit For
        '    End If
        'Next
        Try
            Call evRevit.evAppC.GetLibraryPaths.TryGetValue("Metric Library", resultado)
        Catch ex As Exception
            Try
                Call evRevit.evAppC.GetLibraryPaths.TryGetValue("Library", resultado)
            Catch ex1 As Exception
                Debug.Print(ex.ToString)
            End Try
        End Try
        'Call AdskApplication.appUiCont.ControlledApplication.GetLibraryPaths.TryGetValue("Metric Library", library_folder)
        'Call AdskApplication.appUiCont.ControlledApplication.GetLibraryPaths.TryGetValue("Library", library_folder)
        Return resultado
    End Function

    ''' <summary>
    ''' Load a new icon bitmap from our image folder.
    ''' </summary>
    Public Function NewBitmapImage(ByVal imageName As String) As BitmapImage
        Return New BitmapImage(New Uri(Path.Combine(_dirAppimages, imageName)))
    End Function

    ''' <summary>Cargar imagen desde fichero en disco</summary>
    ''' <param name="embeddedPath">embeddedPath as string (Camino completo a fichero .png)</param>
    ''' <remarks>Llamar con: DameImagenRecurso(ByVal embeddedPath As String)</remarks>
    Public Function DameImagenRecurso(ByVal embeddedPath As String) As System.Windows.Media.ImageSource
        'If embeddedPath.StartsWith(_introLabName) = False Then
        'embeddedPath = _introLabName & "." & embeddedPath
        'End If
        Dim stream As Stream = Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(embeddedPath)
        Dim decoder = New System.Windows.Media.Imaging.BmpBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default)
        Return decoder.Frames(0)
    End Function
    '' 
    ''' <summary>Cargar imagen desde los recursos de la aplicación</summary>
    ''' <param name="bm">bm as intPtr (Resources.red.GetHbitmap())</param>
    ''' <returns>System.Windows.Media.Imaging.BitmapSource</returns>
    ''' <remarks>Llamar con: DameImagenRecurso(Resources.red.GetHbitmap())</remarks>
    Public Function DameImagenRecurso(bm As IntPtr) As System.Windows.Media.Imaging.BitmapSource
        Dim resultado As System.Windows.Media.Imaging.BitmapSource = Nothing
        ''
        resultado = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(bm, IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions)
        Return resultado
    End Function
    ''
    '' ** Llamar siempre a este procedimiento al inicio de las aplicaciones.
    '' Pone el idioma ingles (en) en la aplicación (Revit trabaja en Ingles o US)
    '' Para la correcta conversión de unidades.
    ''
    '' Rellena las variables que guardan: Todas las unidades de Revit, unidades de Longitud y unidades de Area.
    Public Sub InicioAplicacionObligatorio(queFiSP As String)
        '' Poner la cultura ingles en la aplicación (en) Se podría probar también con (us)
        ' Sets the culture to English *** Fundamental para que convierta bien la aplicación
        System.Threading.Thread.CurrentThread.CurrentCulture = New System.Globalization.CultureInfo("en") '("es") ("en-US")
        '' Para coger los datos del documento, tiene que haber un documento abierto.
        NombresParametrosLocalizados()
        If evRevit.evAppUI.ActiveUIDocument.Document IsNot Nothing Then '' Poner el ficher correcto ULMA_PARAMETROS.txt en SharedParameters.
            ''
            utilesRevit.ParametrosSharedCambiaFichero(evRevit.evAppUI.ActiveUIDocument.Document, queFiSP)
            ''
            Dim oTr As Transaction = Nothing
            Try
                '' ** UT_Length (Unidades de longitud del documento)
                unidLDoc = UnidadesDocumento_DameFormatOptions(UnitType.UT_Length)   ''.DisplayUnits para ver el UnitType
                If unidLDoc.DisplayUnits = DisplayUnitType.DUT_MILLIMETERS AndAlso unidLDoc.Accuracy <> 1.0 Then
                    oTr = New Transaction(evRevit.evAppUI.ActiveUIDocument.Document, "CHANGE LENGTH UNITS ACCURACY")
                    oTr.Start()
                    unidLDoc.Accuracy = 1.0
                    evRevit.evAppUI.ActiveUIDocument.Document.GetUnits.SetFormatOptions(UnitType.UT_Length, unidLDoc)
                    oTr.Commit()
                ElseIf unidLDoc.Accuracy > 0.01 Then
                    oTr = New Transaction(evRevit.evAppUI.ActiveUIDocument.Document, "CHANGE LENGTH UNITS ACCURACY")
                    oTr.Start()
                    unidLDoc.Accuracy = 0.01
                    evRevit.evAppUI.ActiveUIDocument.Document.GetUnits.SetFormatOptions(UnitType.UT_Length, unidLDoc)
                    oTr.Commit()
                End If
                'unidLDoc.DisplayUnits
                '' ** UT_Area (Unidades de area del documento)
                unidADoc = UnidadesDocumento_DameFormatOptions(UnitType.UT_Area)  ''.DisplayUnits para ver el UnitType
                If unidADoc.Accuracy > 0.01 Then
                    oTr = New Transaction(evRevit.evAppUI.ActiveUIDocument.Document, "CHANGE AREA UNITS ACCURACY")
                    oTr.Start()
                    unidADoc.Accuracy = 0.01
                    evRevit.evAppUI.ActiveUIDocument.Document.GetUnits.SetFormatOptions(UnitType.UT_Area, unidADoc)
                    oTr.Commit()
                End If
            Catch ex As Exception
                oTr.RollBack()
            Finally
                If oTr IsNot Nothing Then oTr.Dispose()
            End Try
            ''
            '' Unidades en mm y con 2 decimales. Objeto preparado para cambio de unidades.
            If unidDocMM Is Nothing Then
                unidDocMM = New Units(UnitSystem.Metric)
                unidDocMM.SetFormatOptions(UnitType.UT_Length, New FormatOptions(DisplayUnitType.DUT_MILLIMETERS, 0.01))
            End If
            ''
            oTr = Nothing
        End If
    End Sub
    ''' 
    ''' Esta función quita los carácteres inválidos para un nombre de fichero y los sustituye con "_"
    ''' 
    Public Function RemoveInvalidFileNameChars(NombreFichero As String) As String
        For Each invalidChar In IO.Path.GetInvalidFileNameChars
            NombreFichero = NombreFichero.Replace(invalidChar, "_")
        Next
        Return NombreFichero
    End Function
    ''
    Public Function DameRoom(queFam As FamilyInstance, Optional queFase As Phase = Nothing) As Room
        Dim resultado As Room = Nothing

        If queFase Is Nothing Then
            resultado = queFam.Room
            If resultado Is Nothing Then resultado = queFam.FromRoom
            If resultado Is Nothing Then resultado = queFam.ToRoom
        Else
            resultado = queFam.Room(queFase)
            If resultado Is Nothing Then resultado = queFam.FromRoom(queFase)
            If resultado Is Nothing Then resultado = queFam.ToRoom(queFase)
        End If

        If resultado IsNot Nothing Then
            'TaskDialog.Show("Mensajes", resultado.Name & " (" & resultado.Number & ")")
        End If
        ''
        Return resultado
    End Function

    Public Function PonCotaLinea(doc As Autodesk.Revit.DB.Document, vista As View, linea As Line) As Dimension
        Try
            ' Use the Start and End points of our line as the references
            ' Line must come from something in Revit, such as a beam
            Dim referencias As New ReferenceArray()
            referencias.Append(linea.GetEndPointReference(0))
            referencias.Append(linea.GetEndPointReference(1))
            ' Crear linea
            Dim linea1 As Line = Line.CreateBound(linea.GetEndPoint(0), linea.GetEndPoint(1))
            ' create the new dimension
            Dim cota As Dimension = doc.Create.NewDimension(vista, linea, referencias)
            'Dim cota As Dimension = doc.Create.NewAlignment(vista, linea.GetEndPointReference(0), linea.GetEndPointReference(1))
            Return cota
        Catch ex As Exception
            MsgBox("PonCotaLinea --> " & ex.Message)
            Return Nothing
        End Try
    End Function


    Public Function PonCotaCurva(doc As Autodesk.Revit.DB.Document, vista As View, curva As Curve) As Dimension
        Try
            ' Use the Start and End points of our line as the references
            ' Line must come from something in Revit, such as a beam
            Dim references As New ReferenceArray()
            references.Append(curva.GetEndPointReference(0))
            references.Append(curva.GetEndPointReference(1))
            ' Crear linea
            Dim linea As Line = Line.CreateBound(curva.GetEndPoint(0), curva.GetEndPoint(1))
            ' create the new dimension
            Dim cota As Dimension = doc.Create.NewDimension(vista, linea, references)
            'Dim cota As Dimension = doc.Create.NewAlignment(vista, curva.GetEndPointReference(0), curva.GetEndPointReference(1))
            Return cota
        Catch ex As Exception
            MsgBox("PonCotaCurva --> " & ex.Message)
            Return Nothing
        End Try
    End Function
    ''
    Public Function FamilyInstance_DameFaceSuperior(queFa As FamilyInstance) As PlanarFace
        Dim resultado As PlanarFace = Nothing
        ''
        Dim opt As New Options
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.DetailLevel = ViewDetailLevel.Fine
        ''
        Dim oPoint As LocationPoint = CType(queFa.Location, LocationPoint)
        '' Sacar el GeometryElement
        Dim geoElement As GeometryElement = queFa.Geometry(opt)
        '' Recorrer los GeometryObjects que tenga el GeometryElement
        For Each geoObject As GeometryObject In geoElement
            If Not (TypeOf geoObject Is GeometryInstance) Then Continue For
            Dim oGeoIns As GeometryInstance = TryCast(geoObject, GeometryInstance)
            ''
            For Each gObj As GeometryObject In oGeoIns.GetSymbolGeometry
                If TypeOf gObj Is Solid Then
                    Dim oSolid As Solid = TryCast(gObj, Solid)
                    If resultado IsNot Nothing Then Exit For
                    For Each oFace As Face In oSolid.Faces
                        'Dim normal As XYZ = oFace.ComputeNormal(New UV(oPoint.Point.X, oPoint.Point.Y))
                        'If normal.Z > 0.0 Then
                        'resultado = TryCast(oFace, PlanarFace)
                        'Exit For
                        'End If
                        Dim pt As PlanarFace = TryCast(oFace, PlanarFace)
                        If pt IsNot Nothing AndAlso pt.FaceNormal.Y > 0 Then
                            resultado = pt
                            Exit For
                        End If
                    Next
                End If
            Next
        Next
        ''
        Return resultado
    End Function
    ''
    Public Sub ProxyRellenaLineaX(queFa As FamilyInstance)
        Dim opt As New Options
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.DetailLevel = ViewDetailLevel.Fine
        ''
        Dim oPoint As LocationPoint = CType(queFa.Location, LocationPoint)
        '' Sacar el GeometryElement
        Dim geoElement As GeometryElement = queFa.Geometry(opt)
        '' Recorrer los GeometryObjects que tenga el GeometryElement
        For Each geoObject As GeometryObject In geoElement
            If Not (TypeOf geoObject Is GeometryInstance) Then Continue For
            Dim oGeoIns As GeometryInstance = TryCast(geoObject, GeometryInstance)
            ''
            For Each gObj As GeometryObject In oGeoIns.SymbolGeometry
                If TypeOf gObj Is Line Then
                    oLxProxi = CType(gObj, Line)
                    Exit For
                End If
            Next
            If oLxProxi IsNot Nothing Then Exit For
        Next
    End Sub
    ''
    '' nFace = 2 (3ª cara. Face Superior)
    '' nFace = 1 (2ª cara. Face Frontal)
    Public Function FamilyInstance_DamePlanarFaceProxi(queFa As FamilyInstance, nFace As Integer) As PlanarFace
        Dim resultado As PlanarFace = Nothing
        ''
        Dim opt As New Options
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.DetailLevel = ViewDetailLevel.Fine
        ''
        Dim oPoint As LocationPoint = CType(queFa.Location, LocationPoint)
        '' Sacar el GeometryElement
        Dim geoElement As GeometryElement = queFa.Geometry(opt)
        '' Recorrer los GeometryObjects que tenga el GeometryElement
        For Each geoObject As GeometryObject In geoElement
            If Not (TypeOf geoObject Is GeometryInstance) Then Continue For
            Dim oGeoIns As GeometryInstance = TryCast(geoObject, GeometryInstance)
            ''
            For Each gObj As GeometryObject In oGeoIns.SymbolGeometry
                If TypeOf gObj Is Solid Then
                    Dim oSolid As Solid = TryCast(gObj, Solid)
                    '' Si no tiene volumen o no tiene caras, continuar.
                    If oSolid.Volume = 0 Then Continue For
                    If oSolid.Faces Is Nothing OrElse (oSolid.Faces IsNot Nothing AndAlso oSolid.Faces.Size < 6) Then Continue For
                    '' La cara superior es la 3ª (Item(2))
                    '' La cara frontal es la 2ª (Item(1))
                    resultado = CType(oSolid.Faces.Item(nFace), PlanarFace)
                    Exit For
                End If
            Next
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameFaceSuperiorProxi(queFa As FamilyInstance) As Face
        Dim resultado As Face = Nothing
        ''
        Dim opt As New Options
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.DetailLevel = ViewDetailLevel.Fine
        ''
        Dim oPoint As LocationPoint = CType(queFa.Location, LocationPoint)
        '' Sacar el GeometryElement
        Dim geoElement As GeometryElement = queFa.Geometry(opt)
        '' Recorrer los GeometryObjects que tenga el GeometryElement
        For Each geoObject As GeometryObject In geoElement
            If Not (TypeOf geoObject Is GeometryInstance) Then Continue For
            Dim oGeoIns As GeometryInstance = TryCast(geoObject, GeometryInstance)
            ''
            For Each gObj As GeometryObject In oGeoIns.SymbolGeometry
                If TypeOf gObj Is Solid Then
                    Dim oSolid As Solid = TryCast(gObj, Solid)
                    '' Si no tiene volumen o no tiene caras, continuar.
                    If oSolid.Volume = 0 Then Continue For
                    If oSolid.Faces Is Nothing OrElse (oSolid.Faces IsNot Nothing AndAlso oSolid.Faces.Size < 6) Then Continue For
                    '' La cara es la 3ª (Item(2))
                    resultado = oSolid.Faces.Item(2)
                    Exit For
                End If
            Next
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameLineaReferenciaProxi(queFa As FamilyInstance) As Line
        Dim resultado As Line = Nothing
        ''
        Dim opt As New Options
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True
        opt.DetailLevel = ViewDetailLevel.Fine
        ''
        Dim oPoint As LocationPoint = CType(queFa.Location, LocationPoint)
        '' Sacar el GeometryElement
        Dim geoElement As GeometryElement = queFa.Geometry(opt)
        '' Recorrer los GeometryObjects que tenga el GeometryElement
        For Each geoObject As GeometryObject In geoElement
            If Not (TypeOf geoObject Is GeometryInstance) Then Continue For
            Dim oGeoIns As GeometryInstance = TryCast(geoObject, GeometryInstance)
            ''
            For Each gObj As GeometryObject In oGeoIns.SymbolGeometry
                If TypeOf gObj Is Line Then
                    Dim oLine As Line = TryCast(gObj, Line)
                    '' Si no tiene volumen o no tiene caras, continuar.
                    resultado = oLine
                    Exit For
                End If
            Next
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameRotacion(queFa As FamilyInstance) As Double
        Dim resultado As Double = Nothing
        ' Get the Location property and judge whether it exists
        Dim position As Autodesk.Revit.DB.Location = queFa.Location
        ''
        If position Is Nothing Then
            Return resultado
            Exit Function
        Else
            ' If the location is a point location, give the user information
            Dim positionPoint As Autodesk.Revit.DB.LocationPoint = TryCast(position, Autodesk.Revit.DB.LocationPoint)
            If positionPoint IsNot Nothing Then
                resultado = positionPoint.Rotation
            Else
                ' If the location is a curve location, give the user information
                Dim positionCurve As Autodesk.Revit.DB.LocationCurve = TryCast(position, Autodesk.Revit.DB.LocationCurve)
                If positionCurve IsNot Nothing Then
                    Dim pt1 As XYZ = positionCurve.Curve.GetEndPoint(0)
                    Dim pt2 As XYZ = positionCurve.Curve.GetEndPoint(1)
                    Dim lin As Line = CType(positionCurve.Curve, Line)
                    If lin IsNot Nothing Then
                        resultado = lin.Direction.AngleTo(New XYZ(1, 0, 0))
                    End If
                End If
            End If
        End If
        ''
        Return resultado
    End Function
    Public Function FamilyInstance_DamePuntoInsercionBase(queFa As FamilyInstance) As XYZ
        Dim resultado As XYZ = Nothing
        ' Get the Location property and judge whether it exists
        Dim position As Autodesk.Revit.DB.Location = queFa.Location
        ''
        If position Is Nothing Then
            Return resultado
            Exit Function
        Else
            ' If the location is a point location, give the user information
            Dim positionPoint As Autodesk.Revit.DB.LocationPoint = TryCast(position, Autodesk.Revit.DB.LocationPoint)
            If positionPoint IsNot Nothing Then
                resultado = positionPoint.Point
            Else
                ' If the location is a curve location, give the user information
                Dim positionCurve As Autodesk.Revit.DB.LocationCurve = TryCast(position, Autodesk.Revit.DB.LocationCurve)
                If positionCurve IsNot Nothing Then
                    resultado = positionCurve.Curve.GetEndPoint(0)
                End If
            End If
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DamePuntoInsercionBaseElev(queFa As FamilyInstance) As XYZ
        Dim resultado As XYZ = Nothing
        ' Get the Location property and judge whether it exists
        Dim position As Autodesk.Revit.DB.Location = queFa.Location
        ''
        If position Is Nothing Then
            Return resultado
            Exit Function
        Else
            ' If the location is a point location, give the user information
            Dim positionPoint As Autodesk.Revit.DB.LocationPoint = TryCast(position, Autodesk.Revit.DB.LocationPoint)
            If positionPoint IsNot Nothing Then
                resultado = positionPoint.Point
            Else
                ' If the location is a curve location, give the user information
                Dim positionCurve As Autodesk.Revit.DB.LocationCurve = TryCast(position, Autodesk.Revit.DB.LocationCurve)
                If positionCurve IsNot Nothing Then
                    resultado = positionCurve.Curve.GetEndPoint(0)
                End If
            End If
        End If
        ''
        'Dim elev As Double = ParametroElementLeeObjeto(oDoc, queFa,
        Return resultado
    End Function
    '
    Public Function Rotate_Curve(c As Curve, queGrad As Double, Optional ptRot As XYZ = Nothing) As Curve
        Dim resultado As Curve = Nothing
        '
        Dim inicio As XYZ = c.GetEndPoint(0)
        Dim fin As XYZ = c.GetEndPoint(1)
        Dim vec As XYZ = (fin - inicio).Normalize
        Dim eje As XYZ = Nothing
        If ptRot Is Nothing Then
            eje = New XYZ(inicio.X, inicio.Y, inicio.Z + 10)
        Else
            eje = New XYZ(ptRot.X, ptRot.Y, ptRot.Z + 10)
        End If
        '
        Dim rot As Transform = Transform.CreateRotation(eje, U_DameRadianes_DesdeGrados(queGrad))
        resultado = c.CreateTransformed(rot)
        '
        Dim inicioF As XYZ = resultado.GetEndPoint(0)      ' Punto inicial transformada (Será igual que inicio, si la hemos rotado sobre el. ptRot=Nothing)
        Dim finF As XYZ = resultado.GetEndPoint(1)      ' Punto final transformada
        Dim vecF As XYZ = (inicioF - finF).Normalize       ' Vector de dirección transformada.
        Return resultado
    End Function
    Public Function FamilyInstance_RotateGradosLocation(ByRef queFa As FamilyInstance, grados As Double) As Boolean
        Dim rotated As Boolean = False
        Dim location As Location = queFa.Location

        If location IsNot Nothing Then
            Dim aa As XYZ = FamilyInstance_DamePuntoInsercionBase(queFa)
            Dim cc As New XYZ(aa.X, aa.Y, aa.Z + 10)
            Dim axis As Line = Line.CreateBound(aa, cc)
            '' Radianes = [grados] * PI / 180
            rotated = location.Rotate(axis, (grados * Math.PI) / 180)
        End If
        ''
        Return rotated
    End Function
    ''
    Public Function FamilyInstance_RotateGrados(ByRef queFa As FamilyInstance, grados As Double) As Boolean
        Dim rotated As Boolean = False
        Dim location As LocationPoint = TryCast(queFa.Location, LocationPoint)

        If location IsNot Nothing Then
            Dim aa As XYZ = location.Point
            Dim cc As New XYZ(aa.X, aa.Y, aa.Z + 10)
            Dim axis As Line = Line.CreateBound(aa, cc)
            '' Radianes = [grados] * PI / 180
            rotated = location.Rotate(axis, (grados * Math.PI) / 180)
        Else
            rotated = FamilyInstance_RotateGrados_Curve(queFa, grados)
        End If
        ''
        Return rotated
    End Function
    ''
    Public Function FamilyInstance_RotateGrados_Curve(ByRef queFa As FamilyInstance, grados As Double) As Boolean
        Dim rotated As Boolean = False
        Dim location As LocationCurve = TryCast(queFa.Location, LocationCurve)

        If location IsNot Nothing Then
            Dim aa As XYZ = location.Curve.GetEndPoint(0)
            Dim cc As New XYZ(aa.X, aa.Y, aa.Z + 10)
            Dim axis As Line = Line.CreateBound(aa, cc)
            '' Radianes = [grados] * PI / 180
            rotated = location.Rotate(axis, (grados * Math.PI) / 180)
        End If
        ''
        Return rotated
    End Function
    ''
    Public Function FamilyInstance_MoveLocationPoint(ByRef queFa As FamilyInstance, ptFin As XYZ) As Boolean
        Dim moved As Boolean = False
        Dim location As LocationPoint = TryCast(queFa.Location, LocationPoint)
        ''
        If location IsNot Nothing Then
            location.Point = ptFin
            moved = True
        End If
        ''
        Return moved
    End Function
    ''
    Public Function FamilyInstance_Move(ByRef queFa As FamilyInstance, ptFin As XYZ, Optional tVector As XYZ = Nothing) As Boolean
        Dim moved As Boolean = False
        Dim location As LocationPoint = TryCast(queFa.Location, LocationPoint)
        Dim translationVector As XYZ = Nothing
        ''
        If tVector IsNot Nothing Then
            translationVector = tVector
        ElseIf ptFin IsNot Nothing AndAlso location IsNot Nothing Then
            translationVector = ptFin.Subtract(location.Point)
            'Else
            'Return False
            'Exit Function
        End If
        ''
        If location IsNot Nothing AndAlso translationVector IsNot Nothing Then
            moved = location.Move(translationVector)
        Else
            moved = FamilyInstance_Move_Curve(queFa, ptFin)
        End If
        Return moved
    End Function
    ''
    Public Function FamilyInstance_Move_Curve(ByRef queFa As FamilyInstance, ptFin As XYZ, Optional tVector As XYZ = Nothing) As Boolean
        Dim moved As Boolean = False
        Dim location As LocationCurve = TryCast(queFa.Location, LocationCurve)
        Dim translationVector As XYZ
        ''
        If tVector IsNot Nothing Then
            translationVector = tVector
        ElseIf ptFin IsNot Nothing Then
            translationVector = ptFin.Subtract(location.Curve.GetEndPoint(0))
        Else
            Return False
            Exit Function
        End If
        ''
        If location IsNot Nothing Then
            moved = location.Move(translationVector)
        End If
        Return moved
    End Function
    ''
    Public Function TitleBlockDame(doc As Autodesk.Revit.DB.Document, Optional queNombre As String = "") As FamilySymbol
        Dim resultado As FamilySymbol = Nothing   ' ElementId = Nothing
        Dim collector As New FilteredElementCollector(doc)
        '' Seleccionamos la colección de OST_TitleBlocks
        collector.OfClass(GetType(FamilySymbol)).OfCategory(BuiltInCategory.OST_TitleBlocks)
        ''
        If queNombre <> "" Then
            For Each queTitle As Autodesk.Revit.DB.FamilySymbol In collector
                If queTitle.Name = queNombre Then
                    resultado = queTitle
                    Exit For
                End If
            Next
        Else
            resultado = CType(collector.ToList.First, FamilySymbol)
        End If
        ''
        If resultado Is Nothing Then
            TaskDialog.Show("Error buscando TitleBlock", "No se ha encontado" & queNombre)
            Return Nothing
        Else
            Return resultado
        End If
        ''
    End Function
    ''
    Public Function ViewportCrea(doc As Autodesk.Revit.DB.Document, quePlano As ViewSheet, queVista As View, quepunto As XYZ) As Autodesk.Revit.DB.Viewport
        Dim viewportcreado As Autodesk.Revit.DB.Viewport = Nothing
        Try
            viewportcreado = Autodesk.Revit.DB.Viewport.Create(doc, quePlano.Id, queVista.Id, quepunto)
        Catch ex As Exception
            '' Error y devolverá Nothing.
        End Try
        Return viewportcreado
    End Function
    ''
    Public Sub View_Orientation3DPon(ByRef oV3d As View3D, eyeDir As XYZ, upDir As XYZ, forwardDir As XYZ, oFip As FamilyInstance, Optional factor As Double = 1)
        ' By default, the 3D view uses a default orientation.
        ' Change the orientation by creating and setting a ViewOrientation3D 
        'Dim eye As New XYZ(0, -100, 10)
        'Dim up As New XYZ(0, 0, 1)
        'Dim forward As New XYZ(0, 1, 0)
        oV3d.SetOrientation(New ViewOrientation3D(eyeDir, upDir, forwardDir))
        Dim oUvi As UIView = GetActiveUiView(evRevit.evAppUI.ActiveUIDocument)
        If oFip IsNot Nothing Then
            Dim oBb As BoundingBoxXYZ = oFip.BoundingBox(evRevit.evAppUI.ActiveUIDocument.Document.ActiveView)
            oUvi.ZoomAndCenterRectangle(oBb.Min.Subtract(New XYZ(amplia, 0, 0)), oBb.Max.Add(New XYZ(amplia * 2, 0, 0)))
            'oUvi.Zoom(amplia)
        Else
            oUvi.Zoom(factor)
        End If
        ''
        Try
            ' turn off the far clip plane with standard parameter API
            Dim farClip As Parameter = oV3d.LookupParameter("Far Clip Active")
            farClip.[Set](0)
        Catch ex As Exception
            ''
        End Try
    End Sub
    ''
    Public Sub ZoomObjecto(uiD As UIDocument, queObj As Element, Optional factor As Double = 1)
        '' Zoom ajustar sobre objecto
        uiD.ShowElements(queObj)
        '' Localizar vista activa y aplicar un Zoom (Valor superires a 1 = amplia, Valores inferiores a 1=aleja)
        For Each uiV As UIView In uiD.GetOpenUIViews
            If uiV.ViewId.Equals(uiD.ActiveView.Id) Then
                uiV.Zoom(factor)
                Exit For
            End If
        Next
    End Sub
    Public Sub UIView3D_Activa(docUI As Autodesk.Revit.UI.UIDocument)
        Dim v3D As View3D = View3D_Dame(docUI.Document)
        ''
        If v3D Is Nothing Then Exit Sub
        '' Activamos la vista 3D. si no era la activa
        If docUI.Document.ActiveView.Equals(v3D) = False Then
            docUI.ActiveView = v3D
        End If
    End Sub
    '
    ' ***** Guardar la vista actual, para recuperarla al terminar.
    Public Sub UIView_GuardaRestaura(Optional guardar As Boolean = True)
        If evRevit.evAppUI Is Nothing Then Exit Sub
        If evRevit.evAppUI.ActiveUIDocument Is Nothing Then Exit Sub
        If evRevit.evAppUI.ActiveUIDocument.Document Is Nothing Then Exit Sub
        '
        Dim queDocUI As UIDocument = evRevit.evAppUI.ActiveUIDocument
        'Dim queDoc As Document = evRevit.evAppUI.ActiveUIDocument.Document
        'Dim vActive As View = queDocUI.ActiveView
        '
        If guardar = True Then
            '' GUARDA la vista actual
            Try
                ultView = queDocUI.ActiveView.Name
            Catch ex As Exception
            End Try
        ElseIf guardar = False AndAlso ultView <> "" Then
            '' RESTAURAR la vista a la última guardada (ultView)
            If queDocUI.ActiveView.Name = ultView Then
                ' Ya es la activa
                Exit Sub
            End If

            Try
                ' Poner activar la ultima vista (ultView)
                evRevit.evAppUI.ActiveUIDocument.ActiveView = utilesRevit.View_Dame(evRevit.evAppUI.ActiveUIDocument.Document, ultView)
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            ultView = ""
        End If
    End Sub
    '
    Public Sub UIView_Cierra(viewname As String, Optional uidoc As UIDocument = Nothing)    ' Autodesk.Revit.DB.Document = Nothing)
        If uidoc Is Nothing Then uidoc = evRevit.evAppUI.ActiveUIDocument
        Dim openViews As IList(Of UIView) = uidoc.GetOpenUIViews
        For Each openView As Autodesk.Revit.UI.UIView In openViews
            If openView.ViewId = utilesRevit.View_Dame(uidoc.Document, viewname).Id Then
                Try
                    openView.Close()
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End If
        Next
    End Sub
    '
    Public Sub UIView_CierraOtras(Optional queDoc As Document = Nothing)
        Dim queDocUI As UIDocument = Nothing
        If queDoc Is Nothing Then
            queDoc = evRevit.evAppUI.ActiveUIDocument.Document
            queDocUI = evRevit.evAppUI.ActiveUIDocument
        Else
            queDocUI = New UIDocument(queDoc)
        End If
        '
        Dim vActive As View = queDocUI.ActiveView
        '' Cerrar otras vistas
        Try
            Dim openViews As IList(Of UIView) = queDocUI.GetOpenUIViews
            For Each openView As Autodesk.Revit.UI.UIView In openViews
                If openView.ViewId <> vActive.Id Then
                    openView.Close()
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub
    '
    Public Sub UIView3D_ActivaCierraOtras(doc As Autodesk.Revit.DB.Document)
        Dim v3D As View3D = View3D_Dame(doc)
        ''
        If v3D Is Nothing Then Exit Sub
        '' Activamos la vista 3D
        evRevit.evAppUI.ActiveUIDocument.ActiveView = v3D
        '' Cerrar otras vistas
        'Dim actView As Autodesk.Revit.DB.View = uiapp.ActiveUIDocument.ActiveView
        'Dim actViewID = actView.Id
        Dim openViews As IList(Of UIView) = evRevit.evAppUI.ActiveUIDocument.GetOpenUIViews
        For Each openView As Autodesk.Revit.UI.UIView In openViews
            If openView.ViewId <> evRevit.evAppUI.ActiveUIDocument.ActiveView.Id Then
                openView.Close()
            End If
        Next
        If evRevit.evAppUI.ActiveUIDocument.ActiveView.Name <> v3D.Name Then
            '' Activamos la vista 3D
            evRevit.evAppUI.ActiveUIDocument.ActiveView = v3D
        End If
    End Sub
    Public Function View3D_Dame(doc As Autodesk.Revit.DB.Document) As View3D
        Dim resultado As View3D = Nothing
        ''
        Dim collector As New FilteredElementCollector(doc)
        collector.OfClass(GetType(View3D))
        ''
        For Each v As View3D In collector
            'If v.IsTemplate = False Then
            If v.Name = "{3D}" Then
                resultado = v
                Exit For
            End If
        Next
        '
        If resultado Is Nothing AndAlso (collector IsNot Nothing AndAlso collector.Count > 0) Then
            resultado = CType(collector.FirstOrDefault, View3D)
        End If
        Return resultado
    End Function
    ''
    Public Function View_Dame(doc As Autodesk.Revit.DB.Document, viewname As String) As View
        Dim resultado As View = Nothing
        ''
        Dim collector As New FilteredElementCollector(doc)
        collector.OfClass(GetType(View))
        ''
        For Each v As View In collector
            If v.Name = viewname Then
                resultado = v
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    '
    ''
    ''' <summary>
    ''' Return currently active UIView or null.
    ''' </summary>
    Public Function GetActiveUiView(uidoc As UIDocument) As UIView
        Dim doc As Document = uidoc.Document
        Dim view As View = doc.ActiveView
        Dim uiviews As IList(Of UIView) = uidoc.GetOpenUIViews()
        Dim uiview As UIView = Nothing

        For Each uv As UIView In uiviews
            If uv.ViewId.Equals(view.Id) Then
                uiview = uv
                Exit For
            End If
        Next
        Return uiview
    End Function
    ''
    Public Function ProyectoTransfiereCosas(
                                           docOrigen As Autodesk.Revit.DB.Document,
                                            queIds As List(Of ElementId),
                                           docDestino As Autodesk.Revit.DB.Document) As Boolean

        '' docOrigen = Plantilla
        '' docDestino = Documento actual donde traemos todo
        Dim resultado As Boolean = True
        ''

        Using transaction As New Transaction(docDestino, "Copy Resources from template")
            transaction.Start()
            Try
                '' Creamos las opciones de copiar-pegar (Para que utilice los Ids que existan) Sin más avisos.
                Dim copyOptions As CopyPasteOptions = New CopyPasteOptions
                copyOptions.SetDuplicateTypeNamesHandler(New DuplicateTypeNamesHandler)
                '' Copiar a docDestino
                ElementTransformUtils.CopyElements(docOrigen, queIds, docDestino, Transform.Identity, copyOptions)
                ''
                '' Quitar el error de elementos duplicados
                'setup a failure handler to supress warnings when the transaction is commited
                Dim failureOptions As FailureHandlingOptions = transaction.GetFailureHandlingOptions
                failureOptions.SetFailuresPreprocessor(New FailuresPreprocessor)
                ''
                resultado = True
                '' Regeneramos antes de terminar la transacción.
                Try
                    docDestino.Regenerate()
                Catch ex As Exception
                    '' No hacemos nada
                End Try
                '' *********************************************
                transaction.Commit(failureOptions)
            Catch ex As Exception
                resultado = False
                transaction.RollBack()
            End Try
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleBorraVarias(ByRef oDoc As Autodesk.Revit.DB.Document, queSh As String()) As Boolean
        Dim resultado As Boolean = False
        ''
        '' ***** Comprobar si hay Tablas de Planificación.
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(oDoc).OfClass(GetType(ViewSchedule))
        '' Iterar con los informes (Tablas de planificación)
        Dim colBorra As New List(Of ElementId)
        For Each vs As ViewSchedule In collection
            'Dim viewName As String = RemoveInvalidFileNameChars(vs.Name)
            ''Try Loop to avoid any export exceptions
            If queSh.Contains(vs.Name.ToUpper) Then
                colBorra.Add(vs.Id)
            End If
        Next
        ''
        If colBorra IsNot Nothing AndAlso colBorra.Count > 0 Then
            Using oT As New Transaction(oDoc, "Delete Schedules --> " & colBorra.Count)
                oT.Start()
                ''
                Try
                    Call oDoc.Delete(colBorra)
                    oDoc.Regenerate()
                    If oT.GetStatus = TransactionStatus.Started Then oT.Commit()
                    resultado = True
                Catch ex As Exception
                    oT.RollBack()
                    resultado = False
                End Try
            End Using
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleBorraTodas(ByRef oDoc As Autodesk.Revit.DB.Document) As Boolean
        Dim resultado As Boolean = True
        ''
        '' ***** Comprobar si hay Tablas de Planificación.
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(oDoc).OfClass(GetType(ViewSchedule))
        '' Iterar con los informes (Tablas de planificación)
        Dim colBorra As New List(Of ElementId)
        For Each vs As ViewSchedule In collection
            'Dim viewName As String = RemoveInvalidFileNameChars(vs.Name)
            ''Try Loop to avoid any export exceptions
            colBorra.Add(vs.Id)
        Next
        ''
        If colBorra IsNot Nothing AndAlso colBorra.Count > 0 Then
            Using oT As New Transaction(oDoc, "Delete Schedules --> " & colBorra.Count)
                oT.Start()
                For Each queId As ElementId In colBorra
                    Dim queIdBorro As New List(Of ElementId)
                    queIdBorro.Add(queId)
                    ''
                    Try
                        Call oDoc.Delete(queIdBorro)
                    Catch ex As Exception
                        Continue For
                    End Try
                Next
                oDoc.Regenerate()
                If oT.GetStatus = TransactionStatus.Started Then oT.Commit()
            End Using
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleDameUna(oDoc As Autodesk.Revit.DB.Document, queSh As String) As ViewSchedule
        Dim resultado As ViewSchedule = Nothing
        ''
        '' ***** Comprobar si hay Tablas de Planificación.
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(oDoc).OfClass(GetType(ViewSchedule))
        '' Iterar con los informes (Tablas de planificación)
        For Each vs As ViewSchedule In collection
            'Dim viewName As String = RemoveInvalidFileNameChars(vs.Name)
            ''Try Loop to avoid any export exceptions
            If (vs.Name.ToUpper = queSh.ToUpper) Then
                resultado = vs
                Exit For
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleDameNombreIdVarias(oDoc As Autodesk.Revit.DB.Document, queSh() As String, Optional noCopiar As IList(Of String) = Nothing) As Hashtable
        Dim resultado As New Hashtable
        ''
        '' ***** Comprobar si hay Tablas de Planificación.
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(oDoc).OfClass(GetType(ViewSchedule))
        '' Iterar con los informes (Tablas de planificación)
        For Each vs As ViewSchedule In collection
            If noCopiar Is Nothing Then
                If queSh.Contains(vs.Name) Then
                    resultado.Add(vs.Name, vs.Id)
                End If
            Else
                If queSh.Contains(vs.Name) AndAlso noCopiar.Contains(vs.Name) = False Then
                    resultado.Add(vs.Name, vs.Id)
                End If
            End If
        Next
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleDameTodas(doc As Autodesk.Revit.DB.Document, queSh As String) As FilteredElementCollector
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(GetType(ViewSchedule))
        ''
        Return collection
    End Function
    Public Function ScheduleDameTodasEmpiezanPor(doc As Autodesk.Revit.DB.Document, queSh As String) As List(Of ViewSchedule)
        '' Filtro para informes y opciones de exportacion
        Dim collection As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(GetType(ViewSchedule))
        ''
        Dim resultado As New List(Of ViewSchedule)
        For Each vs As ViewSchedule In collection
            If vs.Name.ToUpper.StartsWith(queSh.ToUpper) OrElse vs.Name.ToUpper.Contains(queSh.ToUpper) Then
                If resultado.Contains(vs) = False Then resultado.Add(vs)
            End If
        Next
        Return resultado
    End Function
    ''
    Public Function ScheduleCreaConfigura(doc As Autodesk.Revit.DB.Document, queSh As String, queCampos() As String) As Boolean
        Dim resultado As Boolean = True
        ''
        Dim vs As ViewSchedule = ScheduleDameUna(doc, queSh)
        '' Si no existe, salimos.
        Using transaction As New Transaction(doc, "Creating Schedule " & queSh)
            transaction.Start()
            Try
                If vs Is Nothing Then
                    vs = Autodesk.Revit.DB.ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
                Else
                    vs.Definition.ClearFields()
                End If
                ''
                '' Ahora asignaremos y configuraremos los campos.
                ''        Dim camposobligatorios As String() = New String() { _
                ''    "ITEM_CODE", "ITEM_DESCRIPTION", "ITEM_WEIGHT", nCount, "W_MARKET", _
                ''    nType, "FAMILY_CODE", "ITEM_LENGTH", "ITEM_WIDTH", "ITEM_HEIGHT", _
                ''    "ITEM_GENERIC", "FILTER_ID"}
                ''
                '' 1.- Añadir los campos que necesitamos a queSh
                For Each queField As String In queCampos
                    queField = queField.ToUpper
                    For Each sF As SchedulableField In vs.Definition.GetSchedulableFields
                        Dim nombre As String = sF.GetName(doc).ToUpper
                        ''
                        If nombre = "TYPE" OrElse nombre = "TIPO" OrElse nombre = nType.ToUpper OrElse nombre = nType Then
                            nombre = nType.ToUpper
                        ElseIf nombre = "COUNT" OrElse nombre = "RECUENTO" OrElse nombre = nCount.ToUpper OrElse nombre = nType Then
                            nombre = nCount.ToUpper
                        End If
                        ''
                        If nombre = queField Then
                            Dim sF1 As ScheduleField = vs.Definition.AddField(sF)
                            Exit For
                        End If
                    Next
                Next
                ''
                resultado = True
                doc.Regenerate()
                transaction.Commit()
            Catch ex As Exception
                resultado = False
                transaction.RollBack()
            End Try
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleCreaConfiguraBomLineal(doc As Autodesk.Revit.DB.Document,
                                               queSh As String,
                                               queCampos() As String) As Boolean
        Dim resultado As Boolean = True
        ''
        Dim vs As ViewSchedule = ScheduleDameUna(doc, queSh)
        '' Si no existe, salimos.
        Using transaction As New Transaction(doc, "Creating Schedule " & queSh)
            transaction.Start()
            Try
                If vs Is Nothing Then
                    vs = Autodesk.Revit.DB.ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
                    vs.Name = queSh
                Else
                    vs.Definition.ClearFields()
                    vs.Definition.ClearFilters()
                    vs.Definition.ClearSortGroupFields()
                End If
                ''
                '' Ahora asignaremos y configuraremos los campos.
                '' "ITEM_CODE", "ITEM_LENGTH", "ITEM_DESCRIPTION", _
                '"ITEM_WEIGHT", "TOTAL WEIGHT", "ITEM_GENERIC", nFaseCrea
                ''
                'Dim filterId As ScheduleFieldId = Nothing
                'Dim groupId As ScheduleFieldId = Nothing
                Dim iFil As New List(Of ScheduleFilter)
                Dim iGro As New List(Of ScheduleSortGroupField)

                '' 1.- Añadir los campos que necesitamos a queSh
                For Each queField As String In queCampos
                    queField = queField.ToUpper
                    For Each sF As SchedulableField In vs.Definition.GetSchedulableFields
                        Dim nombre As String = sF.GetName(doc).ToUpper
                        If nombre = queField Then
                            Dim sF1 As ScheduleField = vs.Definition.AddField(sF)
                            '' Poner Filtro y renombrar
                            If nombre = "ITEM_CODE" Then
                                iGro.Add(New ScheduleSortGroupField(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleSortOrder.Ascending))
                                sF1.ColumnHeading = "CODE"
                            ElseIf nombre = "ITEM_LENGTH" Then
                                sF1.ColumnHeading = "QUANTITY"
                            ElseIf nombre = "ITEM_DESCRIPTION" Then
                                sF1.ColumnHeading = "DESCRIPTION"
                            ElseIf nombre = "ITEM_WEIGHT" Then
                                sF1.ColumnHeading = "WEIGHT"
                                '' Crear campo calculado TOTAL WEIGHT

                                Dim shT As ScheduleFieldType = ScheduleFieldType.Formula
                                Dim tw As ScheduleField = vs.Definition.AddField(shT)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                                tw.ColumnHeading = "TOTAL WEIGHT"
                                Dim formatOpts As New FormatOptions()
                                formatOpts.UseDefault = False
                                formatOpts.DisplayUnits = DisplayUnitType.DUT_KILOGRAMS_MASS
                                formatOpts.UnitSymbol = UnitSymbolType.UST_KGM
                                formatOpts.Accuracy = 2
                                tw.SetFormatOptions(formatOpts)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                            ElseIf nombre = "ITEM_GENERIC" Then
                                iFil.Add(New ScheduleFilter(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleFilterType.Equal, "1"))
                                sF1.IsHidden = True
                            ElseIf nombre = nFaseCrea.ToUpper Then
                                sF1.ColumnHeading = "FASE"
                                sF1.IsHidden = True
                            End If
                            ''
                            Exit For
                        End If
                    Next
                Next
                ''
                '' 2.- Poner los filtros y el grupo
                vs.Definition.SetFilters(iFil)
                vs.Definition.SetSortGroupFields(iGro)
                ''
                resultado = True
                transaction.Commit()
            Catch ex As Exception
                resultado = False
                transaction.RollBack()
            End Try
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleCreaConfiguraBomParts(doc As Autodesk.Revit.DB.Document,
                                               queSh As String,
                                               queCampos() As String) As Boolean
        Dim resultado As Boolean = True
        ''
        Dim vs As ViewSchedule = ScheduleDameUna(doc, queSh)
        '' Si no existe, salimos.
        Using transaction As New Transaction(doc, "Creating Schedule " & queSh)
            transaction.Start()
            Try
                If vs Is Nothing Then
                    vs = Autodesk.Revit.DB.ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
                    vs.Name = queSh
                Else
                    vs.Definition.ClearFields()
                    vs.Definition.ClearFilters()
                    vs.Definition.ClearSortGroupFields()
                End If
                ''
                '' Ahora asignaremos y configuraremos los campos.
                '' "ITEM_CODE", nCount, "ITEM_DESCRIPTION", _
                '"ITEM_WEIGHT", "TOTAL WEIGHT", "ITEM_GENERIC", nFaseCrea
                ''
                'Dim filterId As ScheduleFieldId = Nothing
                'Dim groupId As ScheduleFieldId = Nothing
                Dim iFil As New List(Of ScheduleFilter)
                Dim iGro As New List(Of ScheduleSortGroupField)

                '' 1.- Añadir los campos que necesitamos a queSh
                For Each queField As String In queCampos
                    queField = queField.ToUpper
                    For Each sF As SchedulableField In vs.Definition.GetSchedulableFields
                        Dim nombre As String = sF.GetName(doc).ToUpper
                        If nombre = queField Then
                            Dim sF1 As ScheduleField = vs.Definition.AddField(sF)
                            '' Poner Filtro y renombrar
                            If nombre = "ITEM_CODE" Then
                                iGro.Add(New ScheduleSortGroupField(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleSortOrder.Ascending))
                                sF1.ColumnHeading = "CODE"
                            ElseIf nombre = nCount.ToUpper OrElse nombre = nCount Then
                                sF1.ColumnHeading = "QUANTITY"
                            ElseIf nombre = "ITEM_DESCRIPTION" Then
                                sF1.ColumnHeading = "DESCRIPTION"
                            ElseIf nombre = "ITEM_WEIGHT" Then
                                sF1.ColumnHeading = "WEIGHT"
                                '' Crear campo calculado TOTAL WEIGHT
                                Dim tw As ScheduleField = vs.Definition.AddField(sF)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                                tw.ColumnHeading = "TOTAL WEIGHT"
                                Dim formatOpts As New FormatOptions()
                                formatOpts.UseDefault = False
                                formatOpts.DisplayUnits = DisplayUnitType.DUT_KILOGRAMS_MASS
                                formatOpts.UnitSymbol = UnitSymbolType.UST_KGM
                                formatOpts.Accuracy = 2
                                tw.SetFormatOptions(formatOpts)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                            ElseIf nombre = "ITEM_GENERIC" Then
                                iFil.Add(New ScheduleFilter(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleFilterType.LessThan, "1"))
                                sF1.IsHidden = True
                            ElseIf nombre = nFaseCrea.ToUpper Then
                                sF1.ColumnHeading = "FASE"
                                sF1.IsHidden = True
                            End If
                            ''
                            Exit For
                        End If
                    Next
                Next
                '' 2.- Poner los filtros y el grupo
                vs.Definition.SetFilters(iFil)
                vs.Definition.SetSortGroupFields(iGro)
                ''
                resultado = True
                transaction.Commit()
            Catch ex As Exception
                resultado = False
                transaction.RollBack()
            End Try
        End Using
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function ScheduleCreaConfiguraBomArea(doc As Autodesk.Revit.DB.Document,
                                               queSh As String,
                                               queCampos() As String) As Boolean
        Dim resultado As Boolean = True
        ''
        Dim vs As ViewSchedule = ScheduleDameUna(doc, queSh)
        '' Si no existe, salimos.
        Using transaction As New Transaction(doc, "Creating Schedule " & queSh)
            transaction.Start()
            Try
                If vs Is Nothing Then
                    vs = Autodesk.Revit.DB.ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)
                    vs.Name = queSh
                Else
                    vs.Definition.ClearFields()
                    vs.Definition.ClearFilters()
                    vs.Definition.ClearSortGroupFields()
                End If
                ''
                '' Ahora asignaremos y configuraremos los campos.
                '' "ITEM_CODE", "ITEM_AREA", "ITEM_DESCRIPTION", _
                '"ITEM_WEIGHT", "TOTAL WEIGHT", "ITEM_GENERIC", nFaseCrea
                ''
                'Dim filterId As ScheduleFieldId = Nothing
                'Dim groupId As ScheduleFieldId = Nothing
                Dim iFil As New List(Of ScheduleFilter)
                Dim iGro As New List(Of ScheduleSortGroupField)

                '' 1.- Añadir los campos que necesitamos a queSh
                For Each queField As String In queCampos
                    queField = queField.ToUpper
                    For Each sF As SchedulableField In vs.Definition.GetSchedulableFields
                        Dim nombre As String = sF.GetName(doc).ToUpper
                        If nombre = queField Then
                            Dim sF1 As ScheduleField = vs.Definition.AddField(sF)
                            '' Poner Filtro y renombrar
                            If nombre = "ITEM_CODE" Then
                                iGro.Add(New ScheduleSortGroupField(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleSortOrder.Ascending))
                                sF1.ColumnHeading = "CODE"
                            ElseIf nombre = "ITEM_AREA" Then
                                sF1.ColumnHeading = "QUANTITY"
                            ElseIf nombre = "ITEM_DESCRIPTION" Then
                                sF1.ColumnHeading = "DESCRIPTION"
                            ElseIf nombre = "ITEM_WEIGHT" Then
                                sF1.ColumnHeading = "WEIGHT"
                                '' Crear campo calculado TOTAL WEIGHT
                                '' Crear campo calculado TOTAL WEIGHT
                                Dim tw As ScheduleField = vs.Definition.AddField(sF)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                                tw.ColumnHeading = "TOTAL WEIGHT"
                                Dim formatOpts As New FormatOptions()
                                formatOpts.UseDefault = False
                                formatOpts.DisplayUnits = DisplayUnitType.DUT_KILOGRAMS_MASS
                                formatOpts.UnitSymbol = UnitSymbolType.UST_KGM
                                formatOpts.Accuracy = 2
                                tw.SetFormatOptions(formatOpts)
                                tw.DisplayType = ScheduleFieldDisplayType.Totals
                            ElseIf nombre = "ITEM_GENERIC" Then
                                iFil.Add(New ScheduleFilter(vs.Definition.GetFieldId(sF1.FieldIndex), ScheduleFilterType.Equal, "2"))
                                sF1.IsHidden = True
                            ElseIf nombre = nFaseCrea.ToUpper Then
                                sF1.ColumnHeading = "FASE"
                                sF1.IsHidden = True
                            End If
                            ''
                            Exit For
                        End If
                    Next
                Next
                '' 2.- Poner los filtros y el grupo
                vs.Definition.SetFilters(iFil)
                vs.Definition.SetSortGroupFields(iGro)
                ''
                resultado = True
                transaction.Commit()
            Catch ex As Exception
                resultado = False
                transaction.RollBack()
            End Try
        End Using
        ''
        Return resultado
    End Function
    ''
    Public Function ScheduleTieneFields(document As Document, vs As ViewSchedule, listaFields As String()) As Boolean
        Dim resultado As Boolean = False
        'Get all fields from view schedule definition.
        Dim schedulableFields As IList(Of ScheduleFieldId) = vs.Definition.GetFieldOrder
        ''
        Dim contador As Integer = 0
        For Each queField As String In listaFields
            For Each sF As ScheduleFieldId In schedulableFields
                Dim nombre As String = vs.Definition.GetField(sF).GetName.ToUpper
                Dim nombre1 As String = ""
                ''
                If nombre = queField.ToUpper Then   ' OrElse nombre1 = queField.ToUpper Then
                    contador += 1
                    Exit For
                End If
            Next
        Next
        ''
        If listaFields.Count = contador Then
            resultado = True
        Else
            resultado = False
        End If
        Return resultado
    End Function
    ''
    Public Sub ScheduleAgrupaFields(document As Document, ByRef vs As ViewSchedule, listaFields As String())
        Using transaction As New Transaction(document, "Adding GroupField to schedule")
            transaction.Start()
            ''
            '' Borrar las agrupaciones que ya hubiera.
            vs.Definition.ClearSortGroupFields()
            ''
            '' Todos los campos de ViewSchedule, en orden.
            Dim schedulableFields As IList(Of ScheduleFieldId) = vs.Definition.GetFieldOrder
            ''
            Dim contador As Integer = 0
            Dim idsG As New List(Of ScheduleSortGroupField)
            ''
            For x As Integer = 0 To listaFields.Length - 1
                Dim queField As String = listaFields(x).ToUpper
                For Each sF As ScheduleFieldId In schedulableFields
                    Dim nombre As String = vs.Definition.GetField(sF).GetName.ToUpper
                    ''
                    'If queField = nType.ToUpper Then
                    '    If sF.IntegerValue.Equals(CInt(BuiltInParameter.ELEM_TYPE_PARAM)) Then
                    '        nombre = nType.ToUpper
                    '    End If
                    'End If
                    ''
                    If nombre = "TYPE" OrElse nombre = "TIPO" OrElse nombre = nType.ToUpper OrElse nombre = nType Then
                        nombre = nType.ToUpper
                    End If
                    '' 
                    If nombre = queField Then
                        idsG.Add(New ScheduleSortGroupField(sF, ScheduleSortOrder.Ascending))
                        Exit For
                    End If
                Next
            Next
            '' Que se agrupen y junten las filas con la agrupación igual.
            vs.Definition.SetSortGroupFields(idsG)
            vs.Definition.IsItemized = False
            '' Totales
            'vs.Definition.GrandTotalTitle = "Grand total"
            'vs.Definition.ShowGrandTotal = True
            'vs.Definition.ShowGrandTotalCount = True
            'vs.Definition.ShowGrandTotalTitle = True
            ''
            document.Regenerate()
            transaction.Commit()
        End Using
    End Sub
    ''
    ' <summary>
    ' Add fields to view schedule.
    ' </summary>
    ' <param name="schedules">List of view schedule.</param>
    Public Sub ScheduleAddField(document As Document, schedules As List(Of ViewSchedule))
        Using transaction As New Transaction(document, "Adding fields to schedule")
            transaction.Start()

            For Each vs As ViewSchedule In schedules
                'Get all schedulable fields from view schedule definition.
                Dim schedulableFields As IList(Of SchedulableField) = vs.Definition.GetSchedulableFields()

                For Each sf As SchedulableField In schedulableFields
                    Dim fieldAlreadyAdded As Boolean = False
                    'Get all schedule field ids
                    Dim ids As IList(Of ScheduleFieldId) = vs.Definition.GetFieldOrder()
                    For Each id As ScheduleFieldId In ids
                        ' If the GetSchedulableField() method of gotten schedule field returns same
                        ' schedulable field, it means the field is already added to the view schedule.
                        If vs.Definition.GetField(id).GetSchedulableField() = sf Then
                            fieldAlreadyAdded = True
                            Exit For
                        End If
                    Next

                    'If schedulable field doesn't exist in view schedule, add it.
                    If fieldAlreadyAdded = False Then
                        vs.Definition.AddField(sf)
                    End If
                Next
            Next

            transaction.Commit()
        End Using
    End Sub
    ''
    Public Function DameColeccionSegunSystemType(doc As Autodesk.Revit.DB.Document,
                                            queTipo As Type,
                                            Optional queCategoria As BuiltInCategory = BuiltInCategory.INVALID) As FilteredElementCollector
        '' Filtro para queTipo as Type
        '' GetType(ViewSchedule)
        ''
        Dim collection As New FilteredElementCollector(doc)
        collection.OfClass(queTipo)
        If queCategoria <> BuiltInCategory.INVALID Then
            collection.OfCategory(queCategoria)
        End If
        ''
        Return collection
    End Function
    ''
    Public Function DameListaIDSegunSystemType(doc As Autodesk.Revit.DB.Document,
                                            queCategoria As BuiltInCategory) As ICollection(Of Element)
        '' Filtro para informes y opciones de exportacion
        '' GetType(ViewSchedule)
        '' Crear filtro de FamilyInstance
        Dim collector As New FilteredElementCollector(doc)
        collector.OfCategory(queCategoria)  '.OfClass(queTipo)

        Dim anno As ICollection(Of Element) = collector.ToElements
        ''
        Return anno
    End Function
    ''
    Public Function DameColeccionSegunClassCategory(doc As Autodesk.Revit.DB.Document,
                                            queTipo As Type,
                                            Optional categoria As BuiltInCategory = BuiltInCategory.INVALID) As IList(Of Element)
        '' Filtro para informes y opciones de exportacion
        '' GetType(ViewSchedule)
        '' Crear filtro de FamilyInstance
        Dim collector As FilteredElementCollector = Nothing
        Dim anno As IList(Of Element)
        Dim FamilyInstanceFilter As New ElementClassFilter(GetType(FamilyInstance))
        ''
        Dim filtrototal As LogicalAndFilter = Nothing
        ''
        collector = New FilteredElementCollector(doc)
        ''
        If categoria = BuiltInCategory.INVALID Then
            anno = collector.WherePasses(FamilyInstanceFilter).ToElements
        Else
            '' Crear filtro de AnnotationCrop
            Dim FiltroCategoria As New ElementCategoryFilter(categoria)
            filtrototal = New LogicalAndFilter(FamilyInstanceFilter, FiltroCategoria)
            anno = collector.WherePasses(filtrototal).ToElements
        End If
        ''
        'Dim collection As FilteredElementCollector = New FilteredElementCollector(doc).OfClass(queTipo)
        ''
        Return anno
    End Function
    ''
    Public Enum queTipo1
        TODO
        ANNOTATIONSYMBOL
    End Enum
    ''
    'Public Sub BotonesCambiaImagen()
    '    '' iconos=0 (Color Gris), iconos=1 (Color Amarillo)
    '    If icons = 0 Then
    '        btnOptiBoton.Image = DameImagenRecurso(My.Resources.OptionsG32.GetHbitmap())
    '        btnOptiBoton.LargeImage = DameImagenRecurso(My.Resources.OptionsG32.GetHbitmap())
    '        ''
    '        btnCodiBoton.Image = DameImagenRecurso(My.Resources.CodifyG32.GetHbitmap())
    '        btnCodiBoton.LargeImage = DameImagenRecurso(My.Resources.CodifyG32.GetHbitmap())
    '        ''
    '        btnTranBoton.Image = DameImagenRecurso(My.Resources.TranslateG32.GetHbitmap())
    '        btnTranBoton.LargeImage = DameImagenRecurso(My.Resources.TranslateG32.GetHbitmap())
    '        ''
    '        btnExpoBoton.Image = DameImagenRecurso(My.Resources.ExportBOMG32.GetHbitmap())
    '        btnExpoBoton.LargeImage = DameImagenRecurso(My.Resources.ExportBOMG32.GetHbitmap())
    '        ''
    '        btnBomuBoton.Image = DameImagenRecurso(My.Resources.CompleteBOMG32.GetHbitmap())
    '        btnBomuBoton.LargeImage = DameImagenRecurso(My.Resources.CompleteBOMG32.GetHbitmap())
    '        ''
    '        btnHelpBoton.Image = DameImagenRecurso(My.Resources.HelpG32.GetHbitmap())
    '        btnHelpBoton.LargeImage = DameImagenRecurso(My.Resources.HelpG32.GetHbitmap())
    '        ''
    '        btnAboutBoton.Image = DameImagenRecurso(My.Resources.AboutG32.GetHbitmap())
    '        btnAboutBoton.LargeImage = DameImagenRecurso(My.Resources.AboutG32.GetHbitmap())
    '    ElseIf icons = 1 Then
    '        btnOptiBoton.Image = DameImagenRecurso(My.Resources.Options32.GetHbitmap())
    '        btnOptiBoton.LargeImage = DameImagenRecurso(My.Resources.Options32.GetHbitmap())
    '        ''
    '        btnCodiBoton.Image = DameImagenRecurso(My.Resources.Codify32.GetHbitmap())
    '        btnCodiBoton.LargeImage = DameImagenRecurso(My.Resources.Codify32.GetHbitmap())
    '        ''
    '        btnTranBoton.Image = DameImagenRecurso(My.Resources.Translate32.GetHbitmap())
    '        btnTranBoton.LargeImage = DameImagenRecurso(My.Resources.Translate32.GetHbitmap())
    '        ''
    '        btnExpoBoton.Image = DameImagenRecurso(My.Resources.ExportBOM32.GetHbitmap())
    '        btnExpoBoton.LargeImage = DameImagenRecurso(My.Resources.ExportBOMG32.GetHbitmap())
    '        ''
    '        btnBomuBoton.Image = DameImagenRecurso(My.Resources.CompleteBOM32.GetHbitmap())
    '        btnBomuBoton.LargeImage = DameImagenRecurso(My.Resources.CompleteBOM32.GetHbitmap())
    '        ''
    '        btnHelpBoton.Image = DameImagenRecurso(My.Resources.Help32.GetHbitmap())
    '        btnHelpBoton.LargeImage = DameImagenRecurso(My.Resources.Help32.GetHbitmap())
    '        ''
    '        btnAboutBoton.Image = DameImagenRecurso(My.Resources.About32.GetHbitmap())
    '        btnAboutBoton.LargeImage = DameImagenRecurso(My.Resources.About32.GetHbitmap())
    '    End If
    'End Sub

    '
    'Dim arrB As New ArrayList(New String() {"Alinear", "Copiar", "Matriz", "Mover", "Rotar"})
    'PonBotonesEnMiRibbonPanel(panelModiW, "Modificar", "Modificar", arrB, False)
    'Dim arrB1 As New ArrayList(New String() {"Igualar tipo"})
    'PonBotonesEnMiRibbonPanel(panelModiW, "Modificar", "Portapapeles", arrB1, False)
    Public Sub Botones_PonBotonesEnMiRibbonPanel(ByRef RibbonPanelDestino As adWin.RibbonPanel,
                              enRibbon As String,
                              enPanel As String,
                              queBotones As ArrayList,
                              Optional contexto As Boolean = True)
        Dim cuantos As Integer = 0
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        '' *** Primero cogemos RibbonPanelDestino como adWin.RibbonPanel
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            If tab.AutomationName <> enRibbon Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                If oPanel.Source.AutomationName <> enPanel Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    ''
                    If TypeOf oRi Is adWin.RibbonButton AndAlso queBotones.Contains(oRi.AutomationName) Then
                        ritem = oRi.Clone
                        ritem.LargeImage = oRi.LargeImage
                        ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        RibbonPanelDestino.Source.Items.Add(ritem)
                        cuantos += 1
                    ElseIf TypeOf oRi Is adWin.RibbonSplitButton Then
                        '' RibbonItem de cada RibbonSplitButton
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonSplitButton).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonRowPanel Then
                        '' RibbonItem de cada RibbonRowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonRowPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFlowPanel Then
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFlowPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFoldPanel Then
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFoldPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    End If
                Next
                If cuantos >= queBotones.Count Then Exit For
            Next
            If cuantos >= queBotones.Count Then Exit For
        Next
    End Sub
    ''
    Public Sub Botones_PonSplitEnMiRibbonPanel(ByRef RibbonPanelDestino As adWin.RibbonPanel,
                              enRibbon As String,
                              enPanel As String,
                              queSplit As String,
                              Optional contexto As Boolean = True)
        Dim encontrado As Boolean = False
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            If tab.AutomationName <> enRibbon Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                If oPanel.Source.AutomationName <> enPanel Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    '' Si no es un RibbonSplitButton, continuar
                    If Not (TypeOf oRi Is adWin.RibbonSplitButton) Then Continue For
                    ''
                    If oRi.AutomationName = queSplit Then
                        'ritem = oRi.Clone
                        'ritem.LargeImage = oRi.LargeImage
                        'ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        'If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        RibbonPanelDestino.Source.Items.Add(oRi)
                        encontrado = True
                    End If
                Next
                If encontrado = True Then Exit For
            Next
            If encontrado = True Then Exit For
        Next
    End Sub
    '
    Public Sub Cierra_DocumentosTODOS(Optional cerrarRevit As Boolean = False, Optional guardarsinPath As Boolean = True)
        ' Si no hay ningún documento, salir
        If evRevit.evAppUI.ActiveUIDocument Is Nothing Then Exit Sub
        '
        For Each oDoc As Document In evRevit.evAppUI.Application.Documents
            If oDoc.PathName = evRevit.evAppUI.ActiveUIDocument.Document.PathName Then Continue For
            If guardarsinPath = False AndAlso oDoc.PathName = "" Then Continue For
            Try
                Dim resultado As Boolean = oDoc.Close(True)
                'Threading.Thread.Sleep(1000)
            Catch ex As Exception
                ' No se ha podido cerrar el documento
            End Try
        Next
        '
        If cerrarRevit = True Then
            Process.GetCurrentProcess.Kill()
        End If
    End Sub
    '
    Public Sub RibbonPanels_ActivarDesactivar(queRibbonTab As String, colDejar As List(Of String), Optional desactivar As Boolean = True)
        Dim lstPanels As List(Of RibbonPanel) = evRevit.evAppUIC.GetRibbonPanels(queRibbonTab)
        For Each oRp As RibbonPanel In lstPanels
            If colDejar.Contains(oRp.Name) OrElse colDejar.Contains(oRp.Title) Then Continue For
            ' Desactiva el RibbonPanel y todos sus botones
            For Each oRi As RibbonItem In oRp.GetItems
                oRi.Enabled = False
            Next
            oRp.Enabled = desactivar
        Next
    End Sub
    Public Sub Botones_EnMiRibbonPanelPon(ByRef RibbonPanelDestino As adWin.RibbonPanel,
                              enRibbon As String,
                              enPanel As String,
                              queBotones As ArrayList,
                              Optional contexto As Boolean = True)
        Dim cuantos As Integer = 0
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        '' *** Primero cogemos RibbonPanelDestino como adWin.RibbonPanel
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            If tab.AutomationName <> enRibbon Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                If oPanel.Source.AutomationName <> enPanel Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    ''
                    If TypeOf oRi Is adWin.RibbonButton AndAlso queBotones.Contains(oRi.AutomationName) Then
                        ritem = oRi.Clone
                        ritem.LargeImage = oRi.LargeImage
                        ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        RibbonPanelDestino.Source.Items.Add(ritem)
                        cuantos += 1
                    ElseIf TypeOf oRi Is adWin.RibbonSplitButton Then
                        '' RibbonItem de cada RibbonSplitButton
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonSplitButton).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonRowPanel Then
                        '' RibbonItem de cada RibbonRowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonRowPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFlowPanel Then
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFlowPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFoldPanel Then
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFoldPanel).Items
                            If queBotones.Contains(oRi1.AutomationName) Then
                                ritem = oRi1.Clone
                                ritem.LargeImage = oRi1.LargeImage
                                ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                                If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                                RibbonPanelDestino.Source.Items.Add(oRi1)
                                cuantos += 1
                            End If
                        Next
                    End If
                Next
                If cuantos >= queBotones.Count Then Exit For
            Next
            If cuantos >= queBotones.Count Then Exit For
        Next
    End Sub
    ''
    Public Sub SplitEnMiRibbonPanelPon(ByRef RibbonPanelDestino As adWin.RibbonPanel,
                              enRibbon As String,
                              enPanel As String,
                              queSplit As String,
                              Optional contexto As Boolean = True)
        Dim encontrado As Boolean = False
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            If tab.AutomationName <> enRibbon Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                If oPanel.Source.AutomationName <> enPanel Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    '' Si no es un RibbonSplitButton, continuar
                    If Not (TypeOf oRi Is adWin.RibbonSplitButton) Then Continue For
                    ''
                    If oRi.AutomationName = queSplit Then
                        'ritem = oRi.Clone
                        'ritem.LargeImage = oRi.LargeImage
                        'ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        'If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        RibbonPanelDestino.Source.Items.Add(oRi)
                        encontrado = True
                    End If
                Next
                If encontrado = True Then Exit For
            Next
            If encontrado = True Then Exit For
        Next
    End Sub
    ''
    Public Function SketchPlaneVistaActiva(od As Document) As SketchPlane
        '' Setting workplane (Transacción tiene que estar ya iniciada)
        Dim sp As SketchPlane = od.ActiveView.SketchPlane
        od.ActiveView.SketchPlane = sp
        od.ActiveView.ShowActiveWorkPlane()
        Return sp
    End Function
    ''
    Public Function RibbonPanelsListarAFichero() As Autodesk.Revit.UI.Result
        ''
        Dim mensaje As String = ""
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            'If tab.AutomationName = _uiLabName OrElse tab.Name = _uiLabName Then Continue For
            '' RIBBONTABS
            mensaje &= NombreSinSaltos(tab.AutomationName) & " (RibbonTab)" & vbCrLf
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                mensaje &= vbTab & If(oPanel.Source.Name = "", NombreSinSaltos(oPanel.Source.AutomationName), NombreSinSaltos(oPanel.Source.Name)) & " (RibbonPanel)" & vbCrLf
                ''
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    ''
                    If TypeOf oRi Is adWin.RibbonSplitButton Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonSplitButton /" & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonSplitButton
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonSplitButton).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem /" & oRi1.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonRowPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonRowPanel /" & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonRowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonRowPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem /" & oRi1.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFlowPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonFlowPanel /" & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFlowPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem /" & oRi1.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFoldPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonFoldPanel /" & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFoldPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem /" & oRi1.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    Else
                        '' RibbonButton
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonItem /" & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                    End If
                Next
            Next
        Next
        'MsgBox(mensaje)
        Dim ficherofin As String = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.Temp, "RevitRibbons.txt")
        Try
            IO.File.WriteAllText(ficherofin, mensaje)
            Process.Start(ficherofin)
        Catch ex As Exception
            MsgBox("Error in file --> " & ficherofin)
        End Try

        ''
        Return Result.Succeeded
    End Function
    ''
    Public Function NombreSinSaltos(queNombre As String) As String
        If queNombre Is Nothing Then queNombre = ""
        If queNombre <> "" Then
            queNombre = queNombre.Replace(vbCrLf, " ")
            queNombre = queNombre.Replace(vbCr, " ")
            queNombre = queNombre.Replace(vbLf, " ")
        End If
        Return queNombre
    End Function
    ''
    Public Function Documento_EstaAbierto(oApp As Autodesk.Revit.ApplicationServices.Application, fulldoc As String, Optional cerrar As Boolean = False) As Document
        Dim resultado As Document = Nothing
        ''
        For Each queDoc As Autodesk.Revit.DB.Document In oApp.Documents
            If queDoc.PathName = fulldoc Then
                If cerrar = False Then
                    resultado = queDoc
                Else
                    If evRevit.evAppUI.ActiveUIDocument.Document.PathName <> queDoc.PathName Then
                        Try
                            queDoc.Close(False)
                            resultado = Nothing
                        Catch ex As Exception
                            resultado = Nothing
                        End Try
                    End If
                End If
                Exit For
            End If
        Next
        Return resultado
    End Function
    ''
    Public Function Documento_EsDeULMA(queDoc As Document) As Boolean
        Dim resultado As Boolean = False
        If queDoc.IsFamilyDocument = True OrElse queDoc.Title.EndsWith(".rfa") Then
            ' Es un documento de Familia RFA
            Dim manu As String = queDoc.FamilyManager.Parameter(BuiltInParameter.ALL_MODEL_MANUFACTURER).Formula
            '' Si Manufacturer no contiene ULMA O si el tipo de familia no es Modelo Genérico (typeFamily)
            If manu IsNot Nothing AndAlso (manu.ToUpper.Contains("ULMA") AndAlso queDoc.OwnerFamily.FamilyCategoryId.IntegerValue.Equals(CInt(typeFamily)) = True) Then
                resultado = True
            End If
        ElseIf queDoc.IsFamilyDocument = False AndAlso queDoc.PathName.EndsWith(".rvt") Then
            ' Es un documento Revit RVT
            Dim colUlmaFam As IList(Of ElementId) = utilesRevit.FamilySymbol_DameULMA_ID(queDoc, BuiltInCategory.OST_GenericModel)
            If colUlmaFam IsNot Nothing AndAlso colUlmaFam.Count > 0 Then
                resultado = True
            End If
        End If
        '
        Return resultado
    End Function
    'Public Sub Documento_PonActivo(queFi As String)
    '    If evRevit.evAppUI Is Nothing Then
    '        Exit Sub
    '    End If
    '    '
    '    For Each queDoc As Autodesk.Revit.DB.Document In evRevit.evAppUI.Application.Documents
    '        If queDoc.PathName = queFi Then
    '            View3D_Activa(New UIDocument(queDoc))
    '            Exit For
    '        End If
    '    Next
    'End Sub
    ' Enviamos un documento y un List(Of string) para que añada en el los PathName de los links que tenga
    Public Sub LinksDocumento_RellenaList(mainDoc As Document, ByRef lPaths As List(Of String))
        ' Primero procesamos sólo los que no tienen Padre. Los link directos
        Dim collI As FilteredElementCollector = New FilteredElementCollector(mainDoc)
        collI.OfClass(GetType(RevitLinkInstance))
        '
        For Each rIns As RevitLinkInstance In collI
            Dim rType As RevitLinkType = CType(mainDoc.GetElement(rIns.GetTypeId), RevitLinkType)
            ' Solo los que no tienen padre (GetParentID = ElementId.InvalidElementId)
            'If rType.GetParentId.Equals(ElementId.InvalidElementId) Then
            Dim docHijo As Document = rIns.GetLinkDocument
            If docHijo IsNot Nothing AndAlso IO.File.Exists(docHijo.PathName) Then
                Try
                    Dim path As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(docHijo.PathName)
                    Dim Result As LinkLoadResult = rType.LoadFrom(path, New WorksetConfiguration)
                    If lPaths.Contains(docHijo.PathName) = False Then lPaths.Add(docHijo.PathName)
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End If
            'End If
        Next
    End Sub
    Public Function LinksDocumento_DameRevitLinkTypes(ByRef queDoc As Autodesk.Revit.DB.Document) As List(Of RevitLinkType)
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        Dim resultado As New List(Of RevitLinkType)
        Dim filtro As New Autodesk.Revit.DB.FilteredElementCollector(queDoc)
        filtro.OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(GetType(RevitLinkType))
        ''
        If filtro.Count > 0 Then
            For Each e As Element In filtro
                If TypeOf e Is RevitLinkType Then
                    Dim oLink As RevitLinkType = CType(queDoc.GetElement(e.Id), RevitLinkType)
                    resultado.Add(oLink)
                End If
            Next
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function LinksDocumento_DameExternalFileReferenceTypes(ByRef queDoc As Autodesk.Revit.DB.Document, queTipo As ExternalFileReferenceType) As List(Of ExternalFileReferenceType)
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        Dim ids As ICollection(Of ElementId) = ExternalFileUtils.GetAllExternalFileReferences(queDoc)
        ''
        Dim resultado As New List(Of ExternalFileReferenceType)
        ''
        For Each queId As ElementId In ids
            'Dim ele As Element = queDoc.GetElement(queId)
            Dim extF As ExternalFileReference = ExternalFileUtils.GetExternalFileReference(queDoc, queId)
            If extF.ExternalFileReferenceType = queTipo Then
                Dim extTyp As ExternalFileReferenceType = extF.ExternalFileReferenceType
                resultado.Add(extTyp)
            End If
            'Debug.Print(extF.ExternalFileReferenceType.ToString)
        Next

        Dim filtro As New Autodesk.Revit.DB.FilteredElementCollector(evRevit.evAppUI.ActiveUIDocument.Document)
        filtro.OfCategory(BuiltInCategory.OST_RvtLinks).OfClass(GetType(RevitLinkType))
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroFamilyInstanceLee(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queFam As Autodesk.Revit.DB.FamilyInstance,
                                               quePar As String) As String
        Dim resultado As String = ""
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            For Each oPar In queFam.GetParameters(quePar)
                If oPar.HasValue Then
                    valor = oPar.AsString
                    If valor = "" Then valor = oPar.AsValueString
                    Exit For
                End If
            Next
            ''
            If valor = "" Then
                '' Buscamos en FamilySymbol.
                For Each oPar In queFam.Symbol.GetParameters(quePar)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                        Exit For
                    End If
                Next
            End If
            ''
            If valor = "" Then
                '' Buscamos en Family.
                For Each oPar In queFam.Symbol.Family.GetParameters(quePar)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                        Exit For
                    End If
                Next
            End If
            ''
            resultado = valor
        Catch ex As Exception
            resultado = ""
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroFamilyInstanceLeeBuiltInParameter(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queFam As Autodesk.Revit.DB.FamilyInstance,
                                               queId As BuiltInParameter) As String
        Dim resultado As String = ""
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            Try
                oPar = queFam.Parameter(queId)
                If oPar.HasValue Then
                    valor = oPar.AsString
                    If valor = "" Then valor = oPar.AsValueString
                End If
            Catch ex As Exception
                '' No exite el parámetro
            End Try
            ''
            If valor = "" Then
                Try
                    '' Buscamos en FamilySymbol.
                    oPar = queFam.Symbol.Parameter(queId)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                    End If
                Catch ex As Exception
                    '' No exite el parámetro
                End Try
            End If
            ''
            If valor = "" Then
                Try
                    '' Buscamos en Family.
                    oPar = queFam.Symbol.Family.Parameter(queId)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                    End If
                Catch ex As Exception
                    '' No exite el parámetro
                End Try
            End If
            ''
            resultado = valor
        Catch ex As Exception
            resultado = ""
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroFamilySymbolLeeBuiltInParameter(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queFam As Autodesk.Revit.DB.FamilySymbol,
                                               queId As BuiltInParameter) As String
        Dim resultado As String = ""
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            Try
                oPar = queFam.Parameter(queId)
                If oPar.HasValue Then
                    valor = oPar.AsString
                    If valor = "" Then valor = oPar.AsValueString
                End If
            Catch ex As Exception
                '' No exite el parámetro
            End Try
            ''
            If valor = "" Then
                Try
                    '' Buscamos en FamilySymbol.
                    oPar = queFam.Family.Parameter(queId)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                    End If
                Catch ex As Exception
                    '' No exite el parámetro
                End Try
            End If
            ''
            resultado = valor
        Catch ex As Exception
            resultado = ""
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementLeeBuiltInParameter(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queFam As Autodesk.Revit.DB.Element,
                                               queId As BuiltInParameter) As String
        Dim resultado As String = ""
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            Try
                oPar = queFam.Parameter(queId)
                If oPar.HasValue Then
                    valor = oPar.AsString
                    If valor = "" Then valor = oPar.AsValueString
                End If
            Catch ex As Exception
                '' No exite el parámetro
            End Try
            ''
            If valor = "" AndAlso TypeOf queFam Is FamilyInstance Then
                Try
                    '' Buscamos en FamilySymbol.
                    oPar = TryCast(queFam, FamilyInstance).Symbol.Parameter(queId)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                    End If
                Catch ex As Exception
                    '' No exite el parámetro
                End Try
            End If
            ''
            If valor = "" AndAlso TypeOf queFam Is FamilyInstance Then
                Try
                    '' Buscamos en Family.
                    oPar = TryCast(queFam, FamilyInstance).Symbol.Family.Parameter(queId)
                    If oPar.HasValue Then
                        valor = oPar.AsString
                        If valor = "" Then valor = oPar.AsValueString
                    End If
                Catch ex As Exception
                    '' No exite el parámetro
                End Try
            End If
            ''
            resultado = valor
        Catch ex As Exception
            resultado = ""
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementLeeString(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queEle As Autodesk.Revit.DB.Element,
                                               quePar As String) As String
        Dim resultado As String = ""
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            For Each oPar In queEle.GetParameters(quePar)
                'Debug.Print(queEle.GetParameters(quePar).Count.ToString)
                If oPar.HasValue Then
                    valor = oPar.AsString
                    If valor = "" Then valor = oPar.AsValueString
                    Exit For
                End If
            Next
            ''
            resultado = valor
        Catch ex As Exception
            resultado = ""
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementLeeObjeto(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queEle As Autodesk.Revit.DB.Element,
                                               quePar As String) As Object
        Dim resultado As Object = Nothing
        ''
        Try
            Dim valor As Object = Nothing
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            For Each oPar In queEle.GetParameters(quePar)
                If oPar.HasValue Then
                    Select Case oPar.StorageType
                        Case StorageType.Double
                            valor = oPar.AsDouble
                        Case StorageType.ElementId
                            valor = oPar.AsElementId
                        Case StorageType.Integer
                            valor = oPar.AsInteger
                        Case StorageType.String
                            valor = oPar.AsString
                        Case StorageType.None
                            valor = ""
                    End Select
                End If
                If valor Is Nothing Then valor = oPar.AsValueString
                Exit For
            Next
            ''
            resultado = valor
        Catch ex As Exception
            resultado = Nothing
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementLeeBuscaObjeto(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queEle As Autodesk.Revit.DB.Element,
                                               quePar As String) As Object
        Dim resultado As Object = Nothing
        ''
        Try
            Dim valor As Object = Nothing
            Dim eleTy As ElementType = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(queEle.GetTypeId), ElementType)
            Dim oPar As Autodesk.Revit.DB.Parameter = eleTy.LookupParameter(quePar)
            '' Buscamos en FamilyInstance.
            If oPar IsNot Nothing AndAlso oPar.HasValue Then
                Select Case oPar.StorageType
                    Case StorageType.Double
                        valor = oPar.AsDouble
                    Case StorageType.ElementId
                        valor = oPar.AsElementId
                    Case StorageType.Integer
                        valor = oPar.AsInteger
                    Case StorageType.String
                        valor = oPar.AsString
                    Case Else
                        valor = ""
                End Select
                ''
                If valor Is Nothing Then valor = oPar.AsValueString
            End If
            ''
            resultado = valor
        Catch ex As Exception
            resultado = Nothing
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementEscribe(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queEle As Autodesk.Revit.DB.Element,
                                               quePar As String,
                                               queVal As Object) As String
        Dim resultado As String = ""
        ''
        If queVal Is Nothing OrElse (queVal IsNot Nothing AndAlso queVal.ToString = "") OrElse queEle Is Nothing Then
            resultado = "Faltan objetos"
            Return resultado
            Exit Function
        End If
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = Nothing
            '' Buscamos en FamilyInstance.
            For Each oPar In queEle.GetParameters(quePar)
                If oPar.IsReadOnly = False Then
                    Select Case oPar.StorageType
                        Case StorageType.Double
                            '' Si lleva unidades. Las ponemos literales.
                            If ParametroLlevaUnidades(queVal) Then
                                oPar.SetValueString(queVal.ToString)
                            Else
                                oPar.Set(CType(queVal, Double))
                            End If
                        Case StorageType.Integer
                            If ParametroLlevaUnidades(queVal) Then
                                oPar.SetValueString(queVal.ToString)
                            Else
                                oPar.Set(CType(queVal, Integer))
                            End If
                        Case StorageType.String
                            oPar.Set(CType(queVal, String))
                    End Select
                End If
                'Console.WriteLine(oPar.AsValueString)
            Next
            ''
        Catch ex As Exception
            'resultado = "ParametroElementEscribe --> " & ex.Message
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function ParametroElementEscribe(ByRef queDoc As Autodesk.Revit.DB.Document,
                                               queEle As Autodesk.Revit.DB.Element,
                                               quePar As Parameter,
                                               queVal As Object) As String
        Dim resultado As String = ""
        ''
        If queVal Is Nothing OrElse (queVal IsNot Nothing AndAlso queVal.ToString = "") OrElse
            queEle Is Nothing Then
            resultado = "Faltan objetos"
            Return resultado
            Exit Function
        End If
        ''
        Try
            Dim valor As String = ""
            Dim oPar As Autodesk.Revit.DB.Parameter = quePar
            '' Buscamos en FamilyInstance.
            If oPar IsNot Nothing AndAlso oPar.HasValue Then
                If oPar.IsReadOnly = False Then
                    Select Case oPar.StorageType
                        'Case StorageType.Double
                        '    oPar.Set(New ElementId(CInt(queVal)))
                        'Case StorageType.Integer
                        '    oPar.Set(New ElementId(CInt(queVal)))
                        'Case StorageType.String
                        '    oPar.Set(New Autodesk.Revit.DB.ElementId(Integer.Parse(TryCast(queVal, String))))
                        Case StorageType.ElementId
                            If queVal.GetType().Equals(GetType(Autodesk.Revit.DB.ElementId)) Then
                                oPar.Set(TryCast(queVal, Autodesk.Revit.DB.ElementId))
                            ElseIf queVal.GetType().Equals(GetType(String)) Then
                                oPar.Set(New Autodesk.Revit.DB.ElementId(Integer.Parse(TryCast(queVal, String))))
                            ElseIf queVal.GetType().Equals(GetType(Double)) Then
                                oPar.Set(New Autodesk.Revit.DB.ElementId(Convert.ToInt32(queVal)))   'CInt(queVal)))
                            ElseIf queVal.GetType().Equals(GetType(Integer)) Then
                                oPar.Set(New Autodesk.Revit.DB.ElementId(Convert.ToInt32(queVal)))   'CInt(queVal)))
                            Else
                                oPar.Set(New Autodesk.Revit.DB.ElementId(Convert.ToInt32(queVal)))
                            End If
                    End Select
                End If
                'Console.WriteLine(oPar.AsValueString)
            End If
            ''
        Catch ex As Exception
            resultado = "ParametroElementEscribe --> " & ex.Message
        End Try
        ''
        Return resultado
    End Function
    Public Function ParametrosDocumentoLee(Optional queDoc As Autodesk.Revit.DB.Document = Nothing) As String
        Dim resultado As String = ""
        Dim listaSh As String = "***** SHARED PARAMETERS *****" & vbCrLf
        'Dim listaShNo As String = "***** PARAMETERS NORMALES *****" & vbCrLf
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        ''
        Dim bM As BindingMap = queDoc.ParameterBindings       '' Todas las asignaciones de parametros del documento
        ''
        Dim it As DefinitionBindingMapIterator = bM.ForwardIterator
        it.Reset()
        ''
        While it.MoveNext
            Dim defin As InternalDefinition = CType(it.Key, InternalDefinition)
            Dim spar As SharedParameterElement = Nothing
            ''
            Try
                spar = CType(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(defin.Id), SharedParameterElement)
            Catch ex As Exception
                '' Si da error, no es un SharedParameter
            End Try
            ''
            If spar IsNot Nothing Then
                '' Si es un SharedParameter
                listaSh &= vbTab & spar.Name & " --> " & spar.GuidValue.ToString & vbCrLf
            End If
        End While
        ''
        listaSh &= vbCrLf
        resultado = listaSh
        ''
        ''
        Return resultado
    End Function
    ''
    Public Sub ParametroProyectoEscribe(ByRef queDoc As Autodesk.Revit.DB.Document, quePar As String, queVal As Object)
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        Dim resultado As Boolean = False
        Dim t As New Transaction(queDoc, "WRITE PROJECT INFO")
        ''
        '' Leer directamente por el nombre del parámetro.
        Try
            t.Start()
            Dim oPar As Parameter = evRevit.evAppUI.ActiveUIDocument.Document.ProjectInformation.GetParameters(quePar).First
            '' Integer, Double o String
            If TypeOf queVal Is Integer Then
                resultado = oPar.Set(CType(queVal, Integer))
            ElseIf TypeOf queVal Is Double Then
                resultado = oPar.Set(CType(queVal, Double))
            ElseIf TypeOf queVal Is String Then
                resultado = oPar.Set(CType(queVal, String))
            End If
        Catch ex As Exception
            '' No existía el parámetro.
            MsgBox("There is no project parameter --> '" & quePar & "'", MsgBoxStyle.Critical, "ATTENTION")
        End Try
        t.Commit()
    End Sub
    ''
    ''
    Public Sub ParametrosSharedCambiaFichero(ByRef queDoc As Autodesk.Revit.DB.Document, nuevoFi As String)
        If IO.File.Exists(nuevoFi) = False Then Exit Sub
        ''
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        If queDoc Is Nothing Then Exit Sub
        'Dim t As New Transaction(queDoc, "CHANGE SHARED PARAMETERS FILE")
        ''
        Try
            't.Start()
            Dim actual As String = queDoc.Application.SharedParametersFilename
            If actual <> nuevoFi Then
                queDoc.Application.SharedParametersFilename = nuevoFi
                't.Commit()
            Else
                't.RollBack()
            End If
        Catch ex As Exception
            '' No existía el parámetro.
            TaskDialog.Show("ATTENTION", "ParametrosSharedCambiaFichero --> Error changing the shared parameter file:  " & nuevoFi)
            't.RollBack()
        End Try
    End Sub
    ''
    Public Function ParametroLlevaUnidades(queVal As Object) As Boolean
        If queVal.ToString.Contains(" ") Then
            Return True
        Else
            Return False
        End If
    End Function
    ''
    '' ***** PROBAR LOS PROCEDIMIENTOS RawCreateProjectParameter...
    'Dim wall As Category = CachedDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Walls)
    'Dim door As Category = CachedDoc.Settings.Categories.get_Item(BuiltInCategory.OST_Doors)
    'Dim cats1 As CategorySet = CachedApp.Create.NewCategorySet()
    'cats1.Insert(wall)
    'cats1.Insert(door)

    'ProjectParameterCreateFromExistingSharedParameter(CachedApp, Document, "ExistingParameter1", cats1, BuiltInParameterGroup.PG_DATA, False)
    'ProjectParameterCreateFromNewSharedParameter(CachedApp, Document, "NewDefinitionGroup1", "NewParameter1", ParameterType.Text, True, cats1, _
    'BuiltInParameterGroup.PG_DATA, False)
    'ProjectParameterCreate(CachedApp, Document, "TemporarySharedParameter", ParameterType.Text, True, cats1, BuiltInParameterGroup.PG_DATA, _
    '	True)
    Public Sub ProjectParameterCreateFromExistingSharedParameter(ByRef doc As Autodesk.Revit.DB.Document,
                                                                    name As String,
                                                                    cats As CategorySet,
                                                                    group As BuiltInParameterGroup,
                                                                    inst As Boolean)
        Dim defFile As DefinitionFile = doc.Application.OpenSharedParameterFile()
        If defFile Is Nothing Then
            Throw New Exception("No SharedParameter File!")
        End If

        Dim v As IEnumerable(Of Definition) = From dg In defFile.Groups
                                              From d In dg.Definitions
                                              Where d.Name = name
                                              Select d
        If v Is Nothing OrElse v.Count < 1 Then
            Throw New Exception("Invalid Name Input!")
        End If

        Dim def As ExternalDefinition = CType(v(0), ExternalDefinition)
        ''
        Dim binding As Autodesk.Revit.DB.Binding = doc.Application.Create.NewTypeBinding(cats)
        'Dim binding As Autodesk.Revit.DB.InstanceBinding = doc.Application.Create.NewInstanceBinding(cats)
        If inst Then
            binding = doc.Application.Create.NewInstanceBinding(cats)
        End If
        ''
        If doc.ParameterBindings.Contains(def) = False Then doc.ParameterBindings.Insert(def, binding, group)
        ''
        ''Dim roofPitch As Autodesk.Revit.DB.Definition = sharedParameterGroup.Definitions.Create("Roof Pitch", Autodesk.Revit.DB.ParameterType.Text, False)
        ''Dim categories As Autodesk.Revit.DB.CategorySet = doc.Application.Create.NewCategorySet
        ''Dim projInfoCategory As Autodesk.Revit.DB.Category =
        ''doc.Settings.Categories.Item("Project Information")
        ''categories.Insert(projInfoCategory)
        ''Dim instanceBinding As Autodesk.Revit.DB.InstanceBinding instanceBinding = doc.Application.Create.NewInstanceBinding(categories)
        ''doc.ParameterBindings.Insert(roofPitch, InstanceBinding)
    End Sub
    ''
    Public Sub ProjectParameterCreateFromNewSharedParameter(ByRef doc As Autodesk.Revit.DB.Document,
                                                                    defGroup As String,
                                                                      name As String,
                                                                      type As ParameterType,
                                                                      visible As Boolean,
                                                                      cats As CategorySet,
        paramGroup As BuiltInParameterGroup, inst As Boolean)
        Dim defFile As DefinitionFile = doc.Application.OpenSharedParameterFile()
        If defFile Is Nothing Then
            Throw New Exception("No SharedParameter File!")
        End If

        Dim def As ExternalDefinition = TryCast(doc.Application.OpenSharedParameterFile().Groups.Create(defGroup).Definitions.Create(New ExternalDefinitionCreationOptions(name, type)), ExternalDefinition)

        Dim binding As Autodesk.Revit.DB.Binding = doc.Application.Create.NewTypeBinding(cats)
        If inst Then
            binding = doc.Application.Create.NewInstanceBinding(cats)
        End If

        If doc.ParameterBindings.Contains(def) = False Then doc.ParameterBindings.Insert(def, binding, paramGroup)
    End Sub
    ''
    Public Sub ProjectParameterCreate(ByRef doc As Autodesk.Revit.DB.Document,
                                                                    name As String,
                                         type As ParameterType,
                                         visible As Boolean,
                                         cats As CategorySet,
                                         group As BuiltInParameterGroup,
                                            inst As Boolean)
        'InternalDefinition def = new InternalDefinition();
        'Definition def = new Definition();

        Dim oriFile As String = doc.Application.SharedParametersFilename
        Dim tempFile As String = Path.GetTempFileName() + ".txt"
        Using File.Create(tempFile)
        End Using
        doc.Application.SharedParametersFilename = tempFile

        Dim def As ExternalDefinition = TryCast(doc.Application.OpenSharedParameterFile().Groups.Create("TemporaryDefintionGroup").Definitions.Create(New ExternalDefinitionCreationOptions(name, type)), ExternalDefinition)

        doc.Application.SharedParametersFilename = oriFile
        File.Delete(tempFile)

        Dim binding As Autodesk.Revit.DB.Binding = doc.Application.Create.NewTypeBinding(cats)
        If inst Then
            binding = doc.Application.Create.NewInstanceBinding(cats)
        End If
        ''
        If doc.ParameterBindings.Contains(def) = False Then doc.ParameterBindings.Insert(def, binding, group)
        ''
    End Sub
    ''
    Public Sub CeldaTablaPlanificaciónEscribe(ByRef queDoc As Autodesk.Revit.DB.Document,
                                              ByRef queVS As ViewSchedule,
                                              ByRef queTabla As Autodesk.Revit.DB.TableData,
                                              ByRef queHead As Autodesk.Revit.DB.TableSectionData,
                                              ByRef queBody As Autodesk.Revit.DB.TableSectionData,
                                              queFila As Integer,
                                              queColumna As Integer,
                                              queValor As String)
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        ' Dim resultado As Boolean = False
        Dim t As New Transaction(queDoc, "WRITE VIEW SCHEDULE")
        ''
        '' Leer directamente por el nombre del parámetro.
        Try
            t.Start()
            'queVS.GetTableData.GetSectionData(SectionType.Body).SetCellText(queFila, queColumna, queValor)
            Dim nombreParam As String = queVS.GetCellText(SectionType.Body, queBody.FirstRowNumber, queColumna)
            Dim oPar As Parameter = queVS.LookupParameter(nombreParam)
            ''
            oPar.SetValueString(queValor)
            'queTabla.GetSectionData(SectionType.Body).SetCellText(queFila, queColumna, queValor)
            'queBody.SetCellText(queFila, queColumna, queValor)
            t.Commit()

            If evRevit.evAppUI.ActiveUIDocument.Document.IsModified = True Then
                evRevit.evAppUI.ActiveUIDocument.Document.Save()
            End If
        Catch ex As Exception
            '' No existía el parámetro.
            'MsgBox("There is no project parameter --> '" & quePar & "'", MsgBoxStyle.Critical, "ATTENTION")
            t.RollBack()
        End Try
    End Sub
    ''
    '' Site cambia a GenericModel
    Public Function FamilyInstanceName_DameIList(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     nombreFamilia As String,
                                                     Optional categoria As Autodesk.Revit.DB.BuiltInCategory = typeFamily) As IList(Of Element)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        'Dim mensaje As String = ""
        'For Each ele As Element In collector
        '    mensaje &= "Name = " & ele.Name & vbTab & "FamilyName = " & CType(ele, Autodesk.Revit.DB.FamilySymbol).FamilyName & vbCrLf & vbCrLf
        'Next
        'TaskDialog.Show("NOMBRES", mensaje)
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Element)
        query = From element In collector
                Where CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName = nombreFamilia
                Select element
        '' ***** Para coger el ID del FamilySymbol
        Dim famSyms As List(Of Element) = query.ToList
        Dim symbolId As ElementId
        If famSyms.Count = 0 Then
            Return famSyms.ToList()
            Exit Function
        Else
            symbolId = famSyms(0).Id
        End If
        ''
        ''
        '' ***** Ahora creamos un filtro para todos los FamilyInstance con el ID del FamilySymbol
        Dim filter As New FamilyInstanceFilter(queDoc, symbolId)
        ''
        '' ***** Aplicamos el filtro al nuevo FilterElementCollector del documento
        collector = New FilteredElementCollector(queDoc)
        collector.WhereElementIsNotElementType()
        Dim familyInstances As IList(Of Element) = collector.WherePasses(filter).ToElements
        ''
        Return familyInstances
        ''
    End Function
    '

    Public Sub FamilyInstance_CopiaParametrosADirectShape(fOri As FamilyInstance, ByRef queDs As DirectShape)
        ' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        For Each oPar As Autodesk.Revit.DB.Parameter In fOri.Parameters
            If oPar.StorageType = StorageType.ElementId Then
                Call ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, queDs, oPar, oPar.AsInteger)
                Continue For
            End If
            Try
                Dim valor As String = ParametroElementLeeString(fOri.Document, fOri, oPar.Definition.Name)
                ' Por si lleva unidades
                valor = Unidades_DameDouble(valor).ToString  ' valor.Split(" "c)(0)
                If CDbl(valor) = 0 Then valor = ""
                ''
                If valor Is Nothing OrElse (valor IsNot Nothing AndAlso valor = "") Then Continue For
                ' Ponemos milímetros a las longitudes, que son doubles y no llevan unidades.
                If oPar.StorageType = StorageType.Double AndAlso ParametroLlevaUnidades(valor) = False Then
                    Select Case oPar.DisplayUnitType
                        Case DisplayUnitType.DUT_MILLIMETERS
                            valor &= " mm"
                        Case DisplayUnitType.DUT_CENTIMETERS
                            valor &= " cm"
                        Case DisplayUnitType.DUT_DECIMETERS
                            valor &= " dm"
                        Case DisplayUnitType.DUT_METERS
                            valor &= " m"
                        Case DisplayUnitType.DUT_SQUARE_MILLIMETERS
                            valor &= " mm2"
                        Case DisplayUnitType.DUT_SQUARE_CENTIMETERS
                            valor &= " cm2"
                        Case DisplayUnitType.DUT_SQUARE_METERS
                            valor &= " m2"
                        Case DisplayUnitType.DUT_DECIMAL_INCHES
                            valor &= " in"
                        Case DisplayUnitType.DUT_DECIMAL_FEET
                            valor &= " ft"
                        Case DisplayUnitType.DUT_PERCENTAGE
                            valor &= " %"
                    End Select
                End If
                ''
                valor = valor.Trim
                '
                ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, queDs, oPar.Definition.Name, valor)
            Catch ex As Exception
                'Debug.Print(ex.Message)
            End Try
        Next
        '
        ' Eliminar todos los parámetros del DirectShape (Menos ITEM_CODE, ITEM_DESCRIPTION y WEIGHT)
        For Each oPar As Autodesk.Revit.DB.Parameter In queDs.Parameters   ' fOri.Parameters
            If oPar.Definition.Name = "ITEM_CODE" OrElse
                    oPar.Definition.Name = "ITEM_DESCRIPTION" OrElse
                    oPar.Definition.Name = "WEIGHT" OrElse
                    oPar.Definition.Name = nManufacturer Then
                Continue For
            Else
                Try
                    queDs.Parameters.Erase(oPar)
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End If
        Next
    End Sub
    '
    Public Sub FamilyInstance_CopiaParametros(fOri As FamilyInstance, ByRef fFin As FamilyInstance, Optional saltaoffset As Boolean = False, Optional sumaDesfase As Double = 0)
        ' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        '
        For Each oPar As Autodesk.Revit.DB.Parameter In fOri.Parameters
            If oPar.IsReadOnly Then Continue For
            If oPar.StorageType = StorageType.ElementId Then
                Call ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, fFin, oPar, oPar.AsInteger)
                Continue For
            End If
            Try
                Dim valor As String = ParametroElementLeeString(fOri.Document, fOri, oPar.Definition.Name)
                valor = Unidades_DameDouble(valor).ToString  ' valor.Split(" "c)(0)
                If CDbl(valor) = 0 Then valor = ""
                ''
                If valor Is Nothing OrElse valor = "" Then Continue For
                If explode_parameters.Contains(oPar.Definition.Name) Then
                    Continue For
                End If
                ''
                If oPar.Definition.Name = nDesfase Then
                    valor = (CDbl(valor) + sumaDesfase).ToString
                End If
                '' Ponemos milímetros a las longitudes, que son doubles y no llevan unidades.
                If oPar.StorageType = StorageType.Double AndAlso ParametroLlevaUnidades(valor) = False Then
                    Select Case oPar.DisplayUnitType
                        Case DisplayUnitType.DUT_MILLIMETERS
                            valor &= " mm"
                        Case DisplayUnitType.DUT_CENTIMETERS
                            valor &= " cm"
                        Case DisplayUnitType.DUT_DECIMETERS
                            valor &= " dm"
                        Case DisplayUnitType.DUT_METERS
                            valor &= " m"
                        Case DisplayUnitType.DUT_SQUARE_MILLIMETERS
                            valor &= " mm2"
                        Case DisplayUnitType.DUT_SQUARE_CENTIMETERS
                            valor &= " cm2"
                        Case DisplayUnitType.DUT_SQUARE_METERS
                            valor &= " m2"
                        Case DisplayUnitType.DUT_DECIMAL_INCHES
                            valor &= " in"
                        Case DisplayUnitType.DUT_DECIMAL_FEET
                            valor &= " ft"
                        Case DisplayUnitType.DUT_PERCENTAGE
                            valor &= " %"
                    End Select
                End If
                ''
                valor = valor.Trim
                ''
                If saltaoffset = True AndAlso oPar.Definition.Name = nDesfase Then
                    ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, fFin, oPar.Definition.Name, "0 mm")
                Else
                    Dim valorIni As String = ParametroElementLeeString(fOri.Document, fFin, oPar.Definition.Name)
                    If valorIni.ToString <> valor.ToString Then
                        ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, fFin, oPar.Definition.Name, valor)
                        Dim valorFin As String = ParametroElementLeeString(fOri.Document, fFin, oPar.Definition.Name)
                        'Console.WriteLine(valorFin.ToString)
                    End If
                End If
            Catch ex As Exception
                'Debug.Print(ex.Message)
            End Try
        Next
        '' ***** Colección de parámetros que hay que volver a pasar de FamOri a FamFin
        FamilyInstance_CopiaParametrosSOLOMEDIDAS(fOri, fFin)
    End Sub
    ''
    Public Sub FamilyInstance_CopiaParametrosSOLOMEDIDAS(fOri As FamilyInstance, ByRef fFin As FamilyInstance)
        ' ***** Colección de parámetros que hay que volver a pasar de FamOri a FamFin
        ' Si está basado en curva, sacar la longitud.
        Dim oLoc As Location = fFin.Location
        Dim lTemp As Double = 0
        If TypeOf oLoc Is LocationCurve Then
            lTemp = CType(oLoc, LocationCurve).Curve.Length
        End If
        '
        Dim LENGTH As Double = 0
        For Each quePar As String In explode_parameters
            Try
                Dim oParOriList As IList(Of Parameter) = fOri.GetParameters(quePar)
                If oParOriList Is Nothing OrElse (oParOriList IsNot Nothing AndAlso oParOriList.Count = 0) Then Continue For
                '
                Dim oParOri As Parameter = oParOriList.FirstOrDefault
                If oParOri.Definition.Name.Contains("LENGTH") = False Then
                    ' Escribir los que no lleven LENGTH.
                    ParametroElementEscribe(fFin.Document, fFin, quePar, oParOri.AsDouble)
                ElseIf oParOri.Definition.Name.Contains("LENGTH") = True Then
                    ' Sacar el mayor que lleve LENGTH.
                    If oParOri.AsDouble > lTemp Then
                        lTemp = oParOri.AsDouble
                    End If
                ElseIf oParOri.Definition.Name = "LENGTH" Then
                    ' Sacar el que lleve ITEM_LENGTH.
                    LENGTH = oParOri.AsDouble
                End If
            Catch ex As Exception
                'Debug.Print(ex.ToString)
            End Try
        Next
        ' Coger finalmente ITEM_LENGTH (Debería ser el final definitivo)
        If LENGTH > 0 AndAlso lTemp <> LENGTH Then
            lTemp = LENGTH
        End If
        '
        ' Escribir LENGTH con el valor mayor que hayamos detectado. Cualquier parametro que contenga LENGTH.
        For Each quePar As String In explode_parameters
            Try
                Dim oParOriList As IList(Of Parameter) = fOri.GetParameters(quePar)
                If oParOriList Is Nothing OrElse (oParOriList IsNot Nothing AndAlso oParOriList.Count = 0) Then Continue For
                '
                Dim oParOri As Parameter = oParOriList.FirstOrDefault
                If oParOri.Definition.Name.Contains("LENGTH") = True Then
                    ' Escribir los que no lleven LENGTH.
                    ParametroElementEscribe(fFin.Document, fFin, quePar, lTemp)
                End If
            Catch ex As Exception
                'Debug.Print(ex.ToString)
            End Try
        Next
        ParametroElementEscribe(fFin.Document, fFin, "LENGTH", lTemp)
    End Sub
    Public Sub FamilyInstance_CopiaParametrosTODOS(fOri As FamilyInstance, fFin As FamilyInstance, Optional desfaseCero As Boolean = False)
        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        For Each oPar As Autodesk.Revit.DB.Parameter In fOri.Parameters
            'If oPar.IsReadOnly Then Continue For
            Try
                Try
                    '' Solo cambiamos los parámetros compartidos (Shared) que tenga.
                    If oPar.IsReadOnly OrElse oPar.IsShared = False Then Continue For
                    ''
                    Dim valor As Object = ParametroElementLeeObjeto(fOri.Document, fOri, oPar.Definition.Name)
                    valor = Unidades_DameDouble(valor.ToString).ToString  ' valor.Split(" "c)(0)
                    If CDbl(valor) = 0 Then valor = ""
                    ''
                    If valor Is Nothing OrElse valor.ToString = "" Then Continue For
                    ''
                    If oPar.Definition.Name = nDesfase AndAlso desfaseCero = True Then
                        ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, fFin, oPar.Definition.Name, "0 mm")
                        Continue For
                    End If
                    '' Ponemos milímetros a las longitudes, que son doubles y no llevan unidades.
                    'If oPar.StorageType = StorageType.Double AndAlso ParametroLlevaUnidades(valor) = False Then
                    '    Select Case oPar.DisplayUnitType
                    '        Case DisplayUnitType.DUT_MILLIMETERS
                    '            valor &= " mm"
                    '        Case DisplayUnitType.DUT_CENTIMETERS
                    '            valor &= " cm"
                    '        Case DisplayUnitType.DUT_DECIMETERS
                    '            valor &= " dm"
                    '        Case DisplayUnitType.DUT_METERS
                    '            valor &= " m"
                    '        Case DisplayUnitType.DUT_SQUARE_MILLIMETERS
                    '            valor &= " mm2"
                    '        Case DisplayUnitType.DUT_SQUARE_CENTIMETERS
                    '            valor &= " cm2"
                    '        Case DisplayUnitType.DUT_SQUARE_METERS
                    '            valor &= " m2"
                    '        Case DisplayUnitType.DUT_DECIMAL_INCHES
                    '            valor &= " in"
                    '        Case DisplayUnitType.DUT_DECIMAL_FEET
                    '            valor &= " ft"
                    '        Case DisplayUnitType.DUT_PERCENTAGE
                    '            valor &= " %"
                    '    End Select
                    'End If
                    ''
                    ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, fFin, oPar.Definition.Name, valor)
                Catch ex As Exception
                    '' Error cambiando propiedades.
                    Console.WriteLine(ex.Message)
                End Try
            Catch ex As Exception
                Console.WriteLine(ex.Message)
            End Try
        Next
        ''
        'ElementTransformUtils.
    End Sub
    ''
    Public Sub FamilyInstance_CopiaTransformaciones(fOri As FamilyInstance, ByRef fFin As FamilyInstance, Optional rotar As Boolean = True)
        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            '' Si son iguales, no hacemos nada
            If fOri.GetTotalTransform.Equals(fFin.GetTotalTransform) Then
                Exit Sub
            End If
            ''
            Dim ptOri As LocationPoint = DirectCast(fOri.Location, LocationPoint)
            Dim ptFin As LocationPoint = DirectCast(fFin.Location, LocationPoint)
            Dim axis As Line = Nothing
            Dim ptOripro As XYZ = FamilyInstance_DamePuntoInsercionProyecto(fOri)
            Dim rotacion As Double = ptOri.Rotation ' FamilyInstanceDameRotacionProyecto(fOri)
            Dim transf As Transform = FamilyInstance_DameTransformacionProyecto(fOri)
            ''
            '' Comprobar la Simetría
            'If fOri.Mirrored = True Then
            '    Dim ptbX As XYZ = New XYZ(1, 0, 0)
            '    Dim ptbZ As XYZ = New XYZ(0, 0, 1)
            '    Dim plano As Plane = Plane.CreateByOriginAndBasis(ptOripro, ptbX, ptbZ)
            '    ElementTransformUtils.MirrorElement(oDoc, fFin.Id, plano)
            'End If
            If ptOri.Equals(ptFin) = False Then
                Try
                    CType(fFin.Location, LocationPoint).Point = ptOri.Point
                    CType(fFin.Location, LocationPoint).Point = CType(fFin.Location, LocationPoint).Point.Subtract(ptOri.Point)
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
            End If
            ''
            '' Comprobar la Rotación
            If ptFin.Rotation <> ptOri.Rotation AndAlso rotacion <> 0 And rotar = True Then
                Dim zVec As XYZ = fFin.GetTotalTransform.BasisZ
                Dim axix As Line = Line.CreateBound(ptFin.Point, fFin.GetTotalTransform.ScaleBasis(2).BasisZ)
                ElementTransformUtils.RotateElement(evRevit.evAppUI.ActiveUIDocument.Document, fFin.Id, axis, rotacion)  ', ptOri.Rotation)    ' grados
            End If
            axis = Nothing
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub
    ''
    ''
    Public Function FamilyInstance_DameTransformacionProyecto(fOri As FamilyInstance) As Transform
        Dim resultado As Transform = fOri.GetTotalTransform

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    Dim oGins As GeometryInstance = DirectCast(oGobj, GeometryInstance)
                    resultado = oGins.Transform
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DamePuntoInsercionProyecto(fOri As FamilyInstance) As XYZ
        Dim resultado As XYZ = DirectCast(fOri.Location, LocationPoint).Point

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    Dim oGins As GeometryInstance = DirectCast(oGobj, GeometryInstance)
                    resultado = oGins.Transform.Origin
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    ''
    ''
    Public Function FamilyInstance_DameRotacionProyecto(fOri As FamilyInstance) As Double
        Dim resultado As Double = DirectCast(fOri.Location, LocationPoint).Rotation

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    Dim oGins As GeometryInstance = DirectCast(oGobj, GeometryInstance)
                    resultado = oGins.Transform.Origin.AngleTo(oGins.Transform.BasisX)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameGeometryInstanceElement(ByRef fOri As FamilyInstance) As GeometryElement
        Dim resultado As GeometryElement = Nothing

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    Dim oGins As GeometryInstance = DirectCast(oGobj, GeometryInstance)
                    resultado = oGins.GetInstanceGeometry.GetTransformed(Transform.Identity)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    '
    Public Function TessellatedFace_DameDesdeFace(oFace As Face) As List(Of TessellatedFace)
        Dim resultado As New List(Of TessellatedFace)
        '
        Dim oMesh As Mesh = oFace.Triangulate
        Dim args As New List(Of XYZ)(3)
        Dim offset As XYZ = New XYZ
        '
        Dim oEle As Element = evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oFace.Reference.ElementId)
        If TypeOf oEle.Location Is LocationPoint Then
            Dim lPoint As LocationPoint = CType(oEle.Location, LocationPoint)
            offset = lPoint.Point
        End If
        '
        ' Recorrer cada triangulo de la malla. Y crear el TessellatedFace con sus vertices.
        For x As Integer = 0 To oMesh.NumTriangles - 1
            Dim oTri As MeshTriangle = oMesh.Triangle(x)
            Dim p0 As XYZ = oTri.Vertex(0)
            Dim p1 As XYZ = oTri.Vertex(1)
            Dim p2 As XYZ = oTri.Vertex(2)

            p0 = p0.Add(offset)
            p1 = p1.Add(offset)
            p2 = p2.Add(offset)

            args.Clear()
            args.Add(p0)
            args.Add(p1)
            args.Add(p2)
            resultado.Add(New TessellatedFace(args, ElementId.InvalidElementId))
        Next
        '
        Return resultado
    End Function
    Public Function FamilyInstance_CreaDameDirectShapeDesdeMallasCaras(fOri As FamilyInstance) As DirectShape
        Dim resultado As DirectShape = Nothing
        '
        Dim builder As New TessellatedShapeBuilder()
        builder.OpenConnectedFaceSet(True)
        '
        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Dim nombre As String = ""
        nombre = fOri.Name & "_" & fOri.UniqueId    ' oGins.GetInstanceGeometry.GetHashCode.ToString
        Dim oGins As GeometryInstance = Nothing
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    oGins = DirectCast(oGobj, GeometryInstance)
                    Dim objetos As IEnumerator(Of GeometryObject) = oGins.GetInstanceGeometry.GetEnumerator
                    While objetos.MoveNext
                        If (TypeOf objetos.Current Is Solid AndAlso CType(objetos.Current, Solid).Volume > 0) AndAlso objetos.Current.IsElementGeometry Then
                            For Each oFace As Face In CType(objetos.Current, Solid).Faces
                                Dim colTFace As List(Of TessellatedFace) = TessellatedFace_DameDesdeFace(oFace)
                                If colTFace IsNot Nothing AndAlso colTFace.Count > 0 Then
                                    For Each oTFace As TessellatedFace In colTFace
                                        If oTFace.IsValidObject AndAlso builder.DoesFaceHaveEnoughLoopsAndVertices(oTFace) Then
                                            builder.AddFace(oTFace)
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    End While
                End If
            Next
            '
            builder.CloseConnectedFaceSet()
            builder.Target = TessellatedShapeBuilderTarget.AnyGeometry
            builder.Fallback = TessellatedShapeBuilderFallback.Mesh
            builder.Build()

            Dim result As TessellatedShapeBuilderResult = builder.GetBuildResult()

            'ElementId categoryId = New ElementId(BuiltInCategory.OST_GenericModel);
            'DirectShape ds = DirectShape.CreateElement(doc, categoryId,Assembly.GetExecutingAssembly().GetType().GUID.ToString(), Guid.NewGuid().ToString() );
            ''
            If builder.NumberOfCompletedFaceSets > 0 Then
                resultado = DirectShape.CreateElement(evRevit.evAppUI.ActiveUIDocument.Document, New ElementId(typeFamily))
                resultado.ApplicationId = "UCRevit" & RevitVersion
                resultado.ApplicationDataId = "Geometry UCRevit" & RevitVersion
                resultado.Name = nombre
                resultado.SetShape(result.GetGeometricalObjects, DirectShapeTargetViewType.Default)
            End If
        Catch ex As Exception
            '' Fallan algunas familias. Como los pasamanos de las escaleras.
            Console.WriteLine(ex.Message)
            resultado = Nothing
        End Try
        ''
        If resultado.IsValidObject Then
            Return resultado
        Else
            Return Nothing
        End If
    End Function
    Public Function FamilyInstance_CreaDameDirectShape(ByRef fOri As FamilyInstance, Optional conGraphicsStyleId As Boolean = True) As DirectShape
        Dim resultado As DirectShape = Nothing

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Dim colGObj As New List(Of GeometryObject)
        Dim nombre As String = ""
        nombre = fOri.Name & "_" & fOri.UniqueId    ' oGins.GetInstanceGeometry.GetHashCode.ToString
        Dim oGins As GeometryInstance = Nothing
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    oGins = DirectCast(oGobj, GeometryInstance)
                    Dim objetos As IEnumerator(Of GeometryObject) = oGins.GetInstanceGeometry.GetEnumerator
                    While objetos.MoveNext
                        If (TypeOf objetos.Current Is Solid AndAlso CType(objetos.Current, Solid).Volume > 0) AndAlso
                            objetos.Current.IsElementGeometry AndAlso objetos.Current.Visibility = Visibility.Visible Then
                            'Dim grId As ElementId = objetos.Current.GraphicsStyleId
                            'If conGraphicsStyleId = True AndAlso grId.Equals(ElementId.InvalidElementId) = False Then
                            If CType(objetos.Current, Solid).Faces.Size >= 4 AndAlso CType(objetos.Current, Solid).SurfaceArea > 0 Then
                                Try
                                    'If CType(CType(objetos.Current, Solid).Faces.Item(0), PlanarFace).FaceNormal.Z <> 1 Then
                                    colGObj.Add(objetos.Current)
                                    'End If
                                Catch ex As Exception

                                End Try
                            End If
                            'ElseIf conGraphicsStyleId = False AndAlso grId.Equals(ElementId.InvalidElementId) = True Then
                            '    colGObj.Add(objetos.Current)
                            'End If
                        End If
                    End While
                    Exit For
                End If
            Next
            ''
            If colGObj.Count > 0 Then
                resultado = DirectShape.CreateElement(evRevit.evAppUI.ActiveUIDocument.Document, New ElementId(typeFamily))
                resultado.ApplicationId = "UCRevit" & RevitVersion
                resultado.ApplicationDataId = "Geometry UCRevit" & RevitVersion
                resultado.Name = nombre
                resultado.SetShape(colGObj, DirectShapeTargetViewType.Default)
                utilesRevit.ParametroElementEscribe(Nothing, resultado, "ITEM_CODE", utilesRevit.ParametroElementLeeString(fOri.Document, fOri, "ITEM_CODE"))
                utilesRevit.ParametroElementEscribe(Nothing, resultado, "ITEM_DESCRIPTION", utilesRevit.ParametroElementLeeString(fOri.Document, fOri, "ITEM_DESCRIPTION"))
                utilesRevit.ParametroElementEscribe(Nothing, resultado, "WEIGHT", utilesRevit.ParametroElementLeeString(fOri.Document, fOri, "WEIGHT"))

            End If
        Catch ex As Exception
            '' Fallan algunas familias. Como los pasamanos de las escaleras.
            Console.WriteLine(ex.Message)
            resultado = FamilyInstance_CreaDameDirectShapeDesdeMallasCaras(fOri)
        End Try
        ''
        If resultado IsNot Nothing AndAlso resultado.IsValidObject Then
            Return resultado
        Else
            Return Nothing
        End If
    End Function
    '
    Public Function FamilyInstance_CreaDameDirectShapeConGeometryInstance(ByRef fOri As FamilyInstance) As DirectShape
        Dim resultado As DirectShape = Nothing

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Dim colGObj As New List(Of GeometryObject)
        Dim nombre As String = ""
        nombre = fOri.Name & "_" & fOri.UniqueId    ' oGins.GetInstanceGeometry.GetHashCode.ToString
        Dim oGins As GeometryInstance = Nothing
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    oGins = DirectCast(oGobj, GeometryInstance)
                    colGObj.Add(oGins)
                End If
            Next
            ''
            If colGObj.Count > 0 Then
                resultado = DirectShape.CreateElement(evRevit.evAppUI.ActiveUIDocument.Document, New ElementId(typeFamily))
                resultado.ApplicationId = "UCRevit" & RevitVersion
                resultado.ApplicationDataId = "Geometry UCRevit" & RevitVersion
                resultado.Name = nombre
                resultado.SetShape(colGObj, DirectShapeTargetViewType.Default)
            End If
        Catch ex As Exception
            '' Fallan algunas familias. Como los pasamanos de las escaleras.
            Console.WriteLine(ex.Message)
            'Try
            '    resultado = FamilyInstanceDameDirectShapeOriginal(fOri, oGins)
            'Catch ex1 As Exception
            '    Console.WriteLine(ex1.Message)
            'End Try
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameDirectShapeOriginal(ByRef fOri As FamilyInstance, geoIns As GeometryInstance) As DirectShape
        Dim resultado As DirectShape = Nothing

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Dim colGObj As New List(Of GeometryObject)
        Dim nombre As String = ""
        ''
        Try
            Dim oGeleTemp As GeometryElement = fOri.GetOriginalGeometry(utilesRevit.OpcionesDame(fOri.Document, False))
            Dim oGele As GeometryElement = oGeleTemp.GetTransformed(geoIns.Transform)
            ''
            For Each oGobj As GeometryObject In oGele
                nombre = oGobj.GetHashCode.ToString
                If TypeOf oGobj Is Solid AndAlso CType(oGobj, Solid).Volume > 0 Then
                    colGObj.Add(oGobj)
                End If
            Next
            ''
            If colGObj.Count > 0 Then
                resultado = DirectShape.CreateElement(evRevit.evAppUI.ActiveUIDocument.Document, New ElementId(typeFamily))
                resultado.ApplicationId = "UCRevit" & RevitVersion
                resultado.ApplicationDataId = "Geometry UCRevit" & RevitVersion
                resultado.Name = nombre
                resultado.SetShape(colGObj, DirectShapeTargetViewType.Default)
            End If
        Catch ex As Exception
            '' Fallan algunas familias. Como los pasamanos de las escaleras.
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Function FamilyInstance_DameGeometryInstanceSymbol(ByRef fOri As FamilyInstance) As GeometryElement
        Dim resultado As GeometryElement = Nothing

        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            Dim oGele As GeometryElement = fOri.Geometry(utilesRevit.OpcionesDame(fOri.Document))
            ''
            For Each oGobj As GeometryObject In oGele
                If TypeOf oGobj Is GeometryInstance Then
                    Dim oGins As GeometryInstance = DirectCast(oGobj, GeometryInstance)
                    resultado = oGins.GetSymbolGeometry.GetTransformed(Transform.Identity)
                End If
            Next
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Sub FamilyInstance_CopiaTransformacionesGeometry(fOri As FamilyInstance, ByRef fFin As FamilyInstance)
        '' No iniciamos transacción. Cuando llamamos a este procedimiento ya está iniciada.
        Try
            '' Si son iguales, no hacemos nada
            If fOri.GetTotalTransform.Equals(fFin.GetTotalTransform) Then
                Exit Sub
            End If
            '' Geometria
            Dim opciones As New Options
            opciones.ComputeReferences = True
            ''
            Dim oTOri As Transform = Nothing    ' fOri.GetTotalTransform
            '' Cogemos la transformación que tenga el Origen
            Dim geoOri As GeometryElement = fOri.Geometry(opciones)
            For Each queO As GeometryObject In geoOri
                If TypeOf queO Is GeometryInstance Then
                    Dim oGeoIns As GeometryInstance = DirectCast(queO, GeometryInstance)
                    oTOri = oGeoIns.Transform
                End If
            Next
            '' Aplicamos la transformación que tenga el Origen
            Dim geoFin As GeometryElement = fFin.Geometry(opciones)
            For Each queO As GeometryObject In geoFin
                If TypeOf queO Is GeometryInstance Then
                    Call DirectCast(queO, GeometryInstance).GetInstanceGeometry(oTOri)
                    'SolidUtils.CreateTransformed(
                End If
            Next
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
    End Sub
    ''
    Public Function OpcionesDame(oDoc As Document,
                                 Optional computeref As Boolean = True,
                                 Optional includeNonVisible As Boolean = False,
                                 Optional detailL As ViewDetailLevel = ViewDetailLevel.Fine) As Autodesk.Revit.DB.Options
        Dim options As New Options
        options.ComputeReferences = computeref
        options.IncludeNonVisibleObjects = includeNonVisible
        'options.View = oDoc.ActiveView
        options.DetailLevel = detailL
        Return options
    End Function
    ''
    Public Function Curve_DameTransformada(app As Autodesk.Revit.ApplicationServices.Application, element As Autodesk.Revit.DB.Element, geoOptions As Options) As Curve
        Dim resultado As Curve = Nothing
        ' Get geometry element of the selected element
        Dim geoElement As Autodesk.Revit.DB.GeometryElement = element.Geometry(geoOptions)

        ' Get geometry object
        For Each geoObject As GeometryObject In geoElement
            ' Get the geometry instance which contains the geometry information
            Dim instance As Autodesk.Revit.DB.GeometryInstance = TryCast(geoObject, Autodesk.Revit.DB.GeometryInstance)
            If instance IsNot Nothing Then
                For Each o As GeometryObject In instance.SymbolGeometry
                    ' Get curve
                    Dim curve As Curve = TryCast(o, Curve)
                    If curve IsNot Nothing Then
                        ' transfrom the curve to make it in the instance's coordinate space
                        resultado = curve.CreateTransformed(instance.Transform)
                    End If
                Next
            End If
        Next
        Return resultado
    End Function
    ''
    Public Sub FamilyInstance_MueveTodo(colIds As IList(Of ElementId), oFi As FamilyInstance)
        ElementTransformUtils.MoveElements(evRevit.evAppUI.ActiveUIDocument.Document, colIds, DirectCast(oFi.Location, LocationPoint).Point)
    End Sub
    ''
    ''
    '' Site cambia a GenericModel
    Public Function FamilyInstanceName_DameICollection(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     nombreFamilia As String,
                                                     Optional categoria As Autodesk.Revit.DB.BuiltInCategory = typeFamily) As ICollection(Of Element)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance)).OfCategory(categoria)
        '
        ' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        ' Buscaremos en FamilySymbol.FamilyName y en FamilySymbol.Family.Name.
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Element)
        query = From element In collector
                Where CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol.FamilyName = nombreFamilia OrElse
                    CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol.Name = nombreFamilia OrElse
                    CType(element, Autodesk.Revit.DB.FamilyInstance).Name = nombreFamilia
                Select element

        Dim familyInstances As ICollection(Of Element) = query.ToList
        '
        Return familyInstances
        '
    End Function
    '' Site cambia a GenericModel
    Public Function FamilyInstanceName_DameICollectionViejo(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     nombreFamilia As String,
                                                     Optional categoria As Autodesk.Revit.DB.BuiltInCategory = typeFamily) As ICollection(Of Element)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        'Dim mensaje As String = ""
        'For Each ele As Element In collector
        '    mensaje &= "Name = " & ele.Name & vbTab & "FamilyName = " & CType(ele, Autodesk.Revit.DB.FamilySymbol).FamilyName & vbCrLf & vbCrLf
        'Next
        'TaskDialog.Show("NOMBRES", mensaje)
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        '' Buscaremos en FamilySymbol.FamilyName y en FamilySymbol.Family.Name.
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Element)
        query = From element In collector
                Where CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName = nombreFamilia OrElse
                CType(element, Autodesk.Revit.DB.FamilySymbol).Name = nombreFamilia
                Select element

        '' ***** Para coger el ID del FamilySymbol
        Dim famSyms As List(Of Element) = query.ToList
        Dim symbolId As ElementId
        Dim contador As Integer = 0
OTROID:
        If famSyms.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            symbolId = famSyms(contador).Id
        End If
        ''
        ''
        '' ***** Ahora creamos un filtro para todos los FamilyInstance con el ID del FamilySymbol
        Dim filter As New FamilyInstanceFilter(queDoc, symbolId)
        ''
        '' ***** Aplicamos el filtro al nuevo FilterElementCollector del documento
        collector = New FilteredElementCollector(queDoc)
        Dim familyInstances As ICollection(Of Element) = collector.WherePasses(filter).ToElements
        ''
        contador += 1
        If familyInstances.Count = 0 AndAlso contador <= famSyms.Count - 1 Then
            GoTo OTROID
        End If
        ''
        Return familyInstances
        ''
    End Function
    ''
    Public Function FamilyInstanceCategory_DameIList(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     categoria As Autodesk.Revit.DB.BuiltInCategory) As IList(Of Element)
        ''
        '' ***** Filtro para FamilyInstance
        Dim filtrofamilia As New ElementClassFilter(GetType(Autodesk.Revit.DB.FamilyInstance))
        '' ***** Filtro para categoria
        Dim filtrocategoria As New Autodesk.Revit.DB.ElementCategoryFilter(categoria)
        ''
        '' ***** Crear el filtro completo de FamilyInstance + Category
        Dim filtrocompleto As New LogicalAndFilter(filtrofamilia, filtrocategoria)
        ''
        '' ***** Aplicamos el filtro completo al documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector.WhereElementIsNotElementType()
        Dim familyInstances As IList(Of Element) = collector.WherePasses(filtrocompleto).ToElements
        ''
        Return familyInstances
        ''
    End Function
    ''
    Public Function FamilyInstance_DameULMA_FAMILYINSTANCE(queDoc As Autodesk.Revit.DB.Document,
                                               Optional categoria As Autodesk.Revit.DB.BuiltInCategory = typeFamily,
                                               Optional fabricanteContiene As String = fabricante,
                                               Optional Supercomponente As Boolean = False,
                                                Optional vistaActual As Boolean = False,
                                                Optional soloulma As Boolean = True) As List(Of FamilyInstance)
        '' Supercomponente = False. Solo devolverá los que no estén metidos dentro de otra familia.
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As FilteredElementCollector = Nothing
        If vistaActual = False Then
            collector = New FilteredElementCollector(queDoc)
        ElseIf vistaActual = True Then
            collector = New FilteredElementCollector(queDoc, queDoc.ActiveView.Id)
        End If
        ' Solo FamilyInstance
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance))
        '' ***** LINQ para crear IEnumerable de los FamilyInstance y coger el ID
        '' fabricante = ULMA, Supercomponente = False
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.FamilyInstance)
        ''
        If soloulma Then
            ' Si soloulma=True, sólo categoria de Modelos Genéricos
            'collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance)).OfCategory(categoria)
            collector = collector.OfCategory(categoria)
            query = From element In collector
                    Where
                    ParametroFamilySymbolLeeBuiltInParameter(queDoc, CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol, BuiltInParameter.ALL_MODEL_MANUFACTURER).ToUpper.Contains(fabricanteContiene) AndAlso CType(element, Autodesk.Revit.DB.FamilyInstance).SuperComponent Is Nothing = Not Supercomponente
                    Select CType(element, Autodesk.Revit.DB.FamilyInstance)
        Else
            ' Si soloulma=False, no aplicamos filtro de categoría (Todos las categorías de familias)
            query = From element In collector
                    Where CType(element, Autodesk.Revit.DB.FamilyInstance).SuperComponent Is Nothing = Not Supercomponente
                    Select CType(element, Autodesk.Revit.DB.FamilyInstance)
        End If
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    Public Function FamilyInstance_DameULMA_ID(queDoc As Autodesk.Revit.DB.Document,
                                               Optional categoria As Autodesk.Revit.DB.BuiltInCategory = typeFamily,
                                               Optional fabricanteContiene As String = fabricante,
                                               Optional Supercomponente As Boolean = False,
                                                Optional vistaActual As Boolean = False,
                                                Optional soloulma As Boolean = True) As List(Of ElementId)

        '' Supercomponente = False. Solo devolverá los que no estén metidos dentro de otra familia.
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As FilteredElementCollector = Nothing
        If vistaActual = False Then
            collector = New FilteredElementCollector(queDoc)
        Else
            collector = New FilteredElementCollector(queDoc, queDoc.ActiveView.Id)
        End If
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance))
        '' ***** LINQ para crear IEnumerable de los FamilyInstance y coger el ID
        '' fabricante = ULMA, Supercomponente = False
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
        ''
        If soloulma Then
            ' Si soloulma=True, sólo categoria de Modelos Genéricos
            collector = collector.OfCategory(categoria)
            query = From element In collector
                    Where
                    ParametroFamilySymbolLeeBuiltInParameter(queDoc, CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol, BuiltInParameter.ALL_MODEL_MANUFACTURER).ToUpper.Contains(fabricanteContiene) AndAlso CType(element, Autodesk.Revit.DB.FamilyInstance).SuperComponent Is Nothing = Not Supercomponente
                    Select CType(element, Autodesk.Revit.DB.FamilyInstance).Id
        Else
            ' Si soloulma=False, no aplicamos filtro de categoría (Todos las categorías de familias)
            query = From element In collector
                    Where CType(element, Autodesk.Revit.DB.FamilyInstance).SuperComponent Is Nothing = Not Supercomponente
                    Select CType(element, Autodesk.Revit.DB.FamilyInstance).Id
        End If
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    Public Sub FamilyInstance_DameIdsRecursivo(oFam As FamilyInstance, ByRef oIds As List(Of ElementId))
        ''
        If oIds Is Nothing Then oIds = New Collections.Generic.List(Of ElementId)
        'If oIds.Contains(oFam.Id) = Nothing Then oIds.Add(oFam.Id)
        If oFam.GetSubComponentIds.Count = 0 Then
            If oIds.Contains(oFam.Id) = Nothing Then oIds.Add(oFam.Id)
        Else
            For Each oId As ElementId In oFam.GetSubComponentIds
                If TypeOf evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId) Is FamilyInstance Then
                    Dim oFiSub As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId), FamilyInstance)
                    Dim name As String = oFiSub.Name
                    If oFiSub.GetSubComponentIds.Count > 0 Then
                        ' No añadimos los Supercomponentes que tengan Ids
                        FamilyInstance_DameIdsRecursivo(oFiSub, oIds)
                    Else
                        ' Añadimos la FamilyInstance que no tiene Ids dentro.
                        If oIds.Contains(oFiSub.Id) = Nothing Then oIds.Add(oFiSub.Id)
                    End If
                    oFiSub.Dispose()
                    oFiSub = Nothing
                End If
            Next
        End If
    End Sub
    ''
    Public Sub FamilyInstance_DameIdsRecursivoGeometryInstance(oFam As FamilyInstance, ByRef oIds As List(Of ElementId))
        ''
        If oIds Is Nothing Then oIds = New Collections.Generic.List(Of ElementId)
        For Each oId As ElementId In oFam.GetSubComponentIds
            If TypeOf evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId) Is FamilyInstance Then
                Dim oFiSub As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId), FamilyInstance)
                If oIds.Contains(oFiSub.Id) = Nothing Then oIds.Add(oFiSub.Id)
                If oFiSub.GetSubComponentIds.Count > 0 Then
                    FamilyInstance_DameIdsRecursivoGeometryInstance(oFiSub, oIds)
                End If
                oFiSub.Dispose()
                oFiSub = Nothing
            End If
        Next
        ''
    End Sub

    Public Function FamilyInstance_EsDeULMA(queDoc As Autodesk.Revit.DB.Document, oFi As FamilyInstance) As Boolean
        '
        Dim manufacturer As String = ParametroFamilySymbolLeeBuiltInParameter(queDoc, oFi.Symbol, BuiltInParameter.ALL_MODEL_MANUFACTURER)
        If manufacturer.ToUpper.Contains("ULMA") Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function FamilyInstanceID_DameNombresRecursivo(queIdFam As ElementId, indent As Integer, ByRef mensaje As String) As String
        ''
        Dim resultado As String = mensaje
        Dim oFam As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(queIdFam), FamilyInstance)
        For Each oId As ElementId In oFam.GetSubComponentIds
            If TypeOf evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId) Is FamilyInstance Then
                Dim oFiSub As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId), FamilyInstance)
                resultado &= StrDup(indent * 2, " ") & oFiSub.Name & vbCrLf
                If oFiSub.GetSubComponentIds.Count > 0 Then
                    FamilyInstanceID_DameNombresRecursivo(oFiSub.Id, indent + 1, resultado)
                End If
                oFiSub = Nothing
            End If
        Next
        oFam.Dispose()
        oFam = Nothing
        ''
        Return resultado
    End Function
    ''
    Public Sub FamilyInstanceID_DameNombresRecursivos(r As ElementId, ByRef colnom As SortedList(Of String, Integer))
        ''
        Dim oFam As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(r), FamilyInstance)
        For Each oId As ElementId In oFam.GetSubComponentIds
            If TypeOf evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId) Is FamilyInstance Then
                Dim oFiSub As FamilyInstance = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(oId), FamilyInstance)
                ''
                Dim nomFam As String = oFiSub.Symbol.Name
                If colnom.ContainsKey(nomFam) Then
                    colnom(nomFam) += 1
                Else
                    colnom.Add(nomFam, 1)
                End If
                ''
                If oFiSub.GetSubComponentIds IsNot Nothing AndAlso oFiSub.GetSubComponentIds.Count > 0 Then
                    FamilyInstanceID_DameNombresRecursivos(oFiSub.Id, colnom)
                End If
                oFiSub = Nothing
            End If
        Next
        oFam.Dispose()
        oFam = Nothing
        ''
    End Sub
    Public Function FamilySymbolDameNombreEmpiezaPor_ICollection(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     SymbolEmpiezaPor As String,
                                                     categoria As Autodesk.Revit.DB.BuiltInCategory,
                                                     Optional noCopiar As List(Of String) = Nothing) As ICollection(Of ElementId)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        '' Buscaremos en FamilySymbol.Family.Name.
        '' Para coger el ID del FamilySymbol
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
        If noCopiar IsNot Nothing AndAlso noCopiar.Count > 0 Then
            'query = From element In collector _
            '        Where _
            '        (CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName.StartsWith(SymbolEmpiezaPor) OrElse _
            '        CType(element, Autodesk.Revit.DB.FamilySymbol).Name.StartsWith(SymbolEmpiezaPor)) AndAlso _
            '        noCopiar.Contains(CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName) = False AndAlso _
            '        noCopiar.Contains(CType(element, Autodesk.Revit.DB.FamilySymbol).Name) = False
            '        Select element.Id
            query = From element In collector
                    Where
                    CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Name.ToUpper.StartsWith(SymbolEmpiezaPor.ToUpper) AndAlso
                    noCopiar.Contains(CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Name.ToUpper) = False
                    Select element.Id
        Else
            'query = From element In collector _
            '        Where _
            '        (CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName.StartsWith(SymbolEmpiezaPor) OrElse _
            '        CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Name.StartsWith(SymbolEmpiezaPor)) _
            '        Select element.Id
            query = From element In collector
                    Where
                    CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Name.ToUpper.StartsWith(SymbolEmpiezaPor.ToUpper)
                    Select element.Id
        End If

        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    ''
    Public Function FamilySymbol_DameULMA(queDoc As Autodesk.Revit.DB.Document,
                                                     categoria As Autodesk.Revit.DB.BuiltInCategory,
                                                     nombreSymbol As String,
                                                     Optional fabricanteContiene As String = "ULMA") As List(Of FamilySymbol)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        '' Buscaremos en FamilySymbol.Family.Name.
        '' Para coger el ID del FamilySymbol
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.FamilySymbol)
        ''
        query = From element In collector
                Where
                ParametroFamilySymbolLeeBuiltInParameter(queDoc, CType(element, Autodesk.Revit.DB.FamilySymbol), BuiltInParameter.ALL_MODEL_MANUFACTURER).ToUpper.Contains(fabricanteContiene) AndAlso CType(element, Autodesk.Revit.DB.FamilySymbol).Name.StartsWith(nombreSymbol)
                Select CType(element, Autodesk.Revit.DB.FamilySymbol)
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    '
    Public Function Family_DameTODAS(queDoc As Autodesk.Revit.DB.Document) As List(Of ElementId)
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.Family))
        ''
        '' ***** LINQ para crear IEnumerable de los ID de cada family
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
        ''
        query = From element In collector
                Select CType(element, Autodesk.Revit.DB.Family).Id
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    Public Function Family_DameTODASID_DesdeFamilyInstance(queDoc As Autodesk.Revit.DB.Document) As List(Of ElementId)
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance))
        ''
        '' ***** LINQ para crear IEnumerable de los ID de cada family
        Dim query As System.Collections.Generic.IEnumerable(Of ElementId)
        ''
        query = From element In collector
                Select CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol.Family.Id
                Distinct
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    Public Function Family_DameTODASNombres_DesdeFamilyInstance(queDoc As Autodesk.Revit.DB.Document) As List(Of String)
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilyInstance))
        ''
        '' ***** LINQ para crear IEnumerable de los ID de cada family
        Dim query As System.Collections.Generic.IEnumerable(Of String)
        ''
        query = From element In collector
                Select CType(element, Autodesk.Revit.DB.FamilyInstance).Symbol.Family.Name
                Distinct
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    '
    Public Function FamilySymbol_DameULMA_Family(queDoc As Autodesk.Revit.DB.Document,
                                                     categoria As Autodesk.Revit.DB.BuiltInCategory,
                                                     Optional fabricanteContiene As String = "ULMA") As List(Of Family)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        '' Buscaremos en FamilySymbol.Family.Name.
        '' Para coger el ID del FamilySymbol
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Family)
        ''
        query = From element In collector
                Where
                ParametroFamilySymbolLeeBuiltInParameter(queDoc, CType(element, Autodesk.Revit.DB.FamilySymbol), BuiltInParameter.ALL_MODEL_MANUFACTURER).ToUpper.Contains(fabricanteContiene)
                Select CType(element, Autodesk.Revit.DB.FamilySymbol).Family
                Distinct
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    ''
    Public Function FamilySymbol_DameULMA_ID(queDoc As Autodesk.Revit.DB.Document,
                                                     categoria As Autodesk.Revit.DB.BuiltInCategory,
                                                     Optional fabricanteContiene As String = "ULMA") As List(Of ElementId)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.FamilySymbol)).OfCategory(categoria)
        ''
        '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
        '' Buscaremos en FamilySymbol.Family.Name.
        '' Para coger el ID del FamilySymbol
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
        ''
        query = From element In collector
                Where
                ParametroFamilySymbolLeeBuiltInParameter(queDoc, CType(element, Autodesk.Revit.DB.FamilySymbol), BuiltInParameter.ALL_MODEL_MANUFACTURER).ToUpper.Contains(fabricanteContiene)
                Select CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Id
                Distinct
        ''
        If query.Count = 0 Then
            Return Nothing
            Exit Function
        Else
            Return query.ToList
        End If
    End Function
    ''
    'Public Function ParametrosProyectoDame_ICollection(ByRef queDoc As Autodesk.Revit.DB.Document, _
    '                                                 Optional noCopiar As List(Of String) = Nothing) As ICollection(Of ElementId)
    '    ''
    '    '' ***** FilterelementCollector de todos los FamilySymbol del documento
    '    Dim collector As New FilteredElementCollector(queDoc)
    '    collector = collector.OfClass(GetType(Autodesk.Revit.DB.ProjectInfo)).OfCategory(BuiltInCategory.OST_ProjectInformation)
    '    ''
    '    ''
    '    '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
    '    '' Buscaremos en FamilySymbol.FamilyName y en FamilySymbol.Family.Name.
    '    '' Para coger el ID del FamilySymbol
    '    Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
    '    If noCopiar IsNot Nothing Then
    '        query = From element In collector _
    '                Where _
    '                (CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName.StartsWith(SymbolEmpiezaPor) OrElse _
    '                CType(element, Autodesk.Revit.DB.FamilySymbol).Name.StartsWith(SymbolEmpiezaPor)) AndAslo _
    '                noCopiar.Contains(CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName) = False AndAlso _
    '                noCopiar.Contains(CType(element, Autodesk.Revit.DB.FamilySymbol).Name) = False
    '                Select element.Id
    '    Else
    '        query = From element In collector _
    '                Where _
    '                (CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName.StartsWith(SymbolEmpiezaPor) OrElse _
    '                CType(element, Autodesk.Revit.DB.FamilySymbol).Name.StartsWith(SymbolEmpiezaPor)) _
    '                Select element.Id
    '    End If

    '    ''
    '    If query.Count = 0 Then
    '        Return Nothing
    '        Exit Function
    '    Else
    '        Return query.ToList
    '    End If
    'End Function
    ''
    ' Activar vista de un Document (Seleccionando todo y haciendo zoom sobre seleccion)
    Public Function Zomm_Elements_View(queDoc As Autodesk.Revit.DB.Document, Optional conZoom As Boolean = True) As List(Of ElementId)
        ''
        '' ***** FilterelementCollector de todos los FamilySymbol del documento
        Dim uDoc As UIDocument = New UIDocument(queDoc)
        Dim collector As New FilteredElementCollector(queDoc, uDoc.GetOpenUIViews.FirstOrDefault.ViewId)
        ''
        '' ***** LINQ para crear IEnumerable de los ID
        Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.ElementId)
        query = From element In collector
                Select element.Id
        '
        ' Hacer un zoom y activar vista
        uDoc.ShowElements(query.ToList)
        '
        If query.Count = 0 Then
            Return Nothing
        Else
            Return query.ToList
        End If
    End Function
    Public Function ParametroCompartido_DameGUID(ByRef queDoc As Autodesk.Revit.DB.Document,
                                                     nombre As String) As System.Guid
        ''
        '' ***** Filtro para FamilyInstance
        'Dim filtroParametroCompartido As New ElementClassFilter(GetType(Autodesk.Revit.DB.SharedParameterElement))
        ' ''
        ' '' ***** Crear el filtro completo
        'Dim filtrocompleto As New LogicalAndFilter(filtroParametroCompartido)
        ' ''
        ' '' ***** Aplicamos el filtro completo al documento
        'Dim collector As New FilteredElementCollector(queDoc)
        'collector.WhereElementIsNotElementType()
        'Dim familyInstances As IList(Of Element) = collector.WherePasses(filtrocompleto).ToElements
        ' ''
        'Return familyInstances
        ''
    End Function
    ''
    Public Sub CopiaUsuario_RevitIni()
        If IO.File.Exists(IO.Path.Combine(_dirApp, "Revit.ini")) = False Then Exit Sub
        ''
        'Dim RevitUsuario As String = _
        'Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile) & _
        '"\AppData\Roaming\Autodesk\Revit\Autodesk Revit 2018\Revit.ini"
        ''
        Try
            utilesRevit.Fichero_QuitarSoloLectura(RevitIniUser)
            IO.File.Copy(IO.Path.Combine(_dirApp, "Revit.ini"), RevitIniUser, True)
        Catch ex As Exception
            '' No hacemos nada
        End Try
    End Sub
    ''
    Public Sub Cambia_OmniClassTaxonomytxt(onmiclassfile As String)
        ''
        If IO.File.Exists(onmiclassfile) = False Then Exit Sub
        ''
        '' *** Leer "OmniClassTaxonomy_ULMA.txt", linea a linea. Desde carpeta instalación. En Hashtable
        Dim colUlma As New SortedList(Of String, String())
        Try
            Using objReader As IO.StreamReader = New IO.StreamReader(onmiclassfile, System.Text.Encoding.Default, True)
                Dim sLine As String = ""
                Dim contador As Integer = 0
                Do
                    sLine = objReader.ReadLine()
                    If sLine Is Nothing Then Exit Do
                    If sLine = "" Then Exit Do
                    ''
                    contador += 1
                    ''
                    Dim valores() As String = sLine.ToUpper.Split(Chr(9))    'vbtab
                    '' Comprobaciones varias, antes de procesar
                    '' Si el codigo está en valores(0) Si no, no podemos continuar
                    If valores.Count < 3 Then
                        Continue Do
                    ElseIf valores(0) = "" Then
                        Continue Do
                    End If
                    '' Comprobar si están los 4 valores (por si no tiene el tabulador al final)
                    If valores.Count < 4 Then
                        ReDim Preserve valores(3)   ' 4 valores (de 0 a 3)
                    End If
                    ''
                    '' Añadirlo al Hashtable
                    If colUlma.ContainsKey(valores(0)) = False Then colUlma.Add(valores(0), valores)
                Loop Until sLine Is Nothing
                ''
                objReader.Close()
            End Using
        Catch ex As Exception
            '' No hacemos nada. colUlma está creado, pero no tendrá elementos o no los que tengan errores.
        End Try
        ''
        '' *** Leer "OmniClassTaxonomy.txt" del usuario. Linea a linea.
        Dim colUsuario As New SortedList(Of String, String())       '' Aquí vamos almacenando cada linea final
        Try
            Using objReader As IO.StreamReader = New IO.StreamReader(OmniClassTaxonomyUser, System.Text.Encoding.Default, True)
                Dim sLine As String = ""
                Dim contador As Integer = 0
                Do
                    sLine = objReader.ReadLine()
                    If sLine Is Nothing Then Exit Do
                    If sLine = "" Then Exit Do
                    ''
                    contador += 1
                    ''
                    Dim valores() As String = sLine.ToUpper.Split(Chr(9))    'vbtab
                    '' Comprobaciones varias, antes de procesar
                    '' Si el codigo está en valores(0) Si no, no podemos continuar
                    If valores.Count < 3 Then
                        Continue Do
                    ElseIf valores(0) = "" Then
                        Continue Do
                    End If
                    '' Comprobar si están los 4 valores (por si no tiene el tabulador al final)
                    If valores.Count < 4 Then
                        ReDim Preserve valores(3)   ' 4 valores (de 0 a 3)
                    End If
                    ''
                    '' Comparar si existe
                    If colUsuario.ContainsKey(valores(0)) = False Then colUsuario.Add(valores(0), valores)
                Loop Until sLine Is Nothing
                ''
                objReader.Close()
            End Using
        Catch ex As Exception
            '' No hacemos nada. colUlma está creado, pero no tendrá elementos o no los que tengan errores.
        End Try
        ''
        '' Comparar codigo con el de "OmniClassTaxonomy_ULMA.txt"
        '' Para ver si lo incluimos, lo saltamos o lo editamos
        '' ** Modificar los códigos del usuario que estén cambiados.
        Dim cambiado As Boolean = False
        For Each queCo As String In colUlma.Keys
            '' Si es codigo ya está, lo cambiamos (Comentar, si queremos mantener el del usuario)
            If colUsuario.ContainsKey(queCo) Then
                '' Modificar los datos que ya hubiera
                Dim valUsuario, valUlma As String
                valUsuario = Join(CType(colUsuario(queCo), String()), " ")
                valUlma = Join(CType(colUlma(queCo), String()), " ")
                If valUsuario <> valUlma Then
                    '' Quitar comentarios, si queremos que sobrescriba lo que venga de "OmniClassTaxonomy_ULMA.txt"
                    '' En caso de tener los mismos códigos.
                    'colUsuario(queCo) = colUlma(queCo)
                    'cambiado = True
                End If
            Else
                '' Añadir el codigo y datos que no estaban.
                colUsuario.Add(queCo, colUlma(queCo))
                cambiado = True
            End If
        Next
        ''
        '' *** Escribir el nuevo "OmniClassTaxonomy_ULMA.txt" (Solo si se ha modificado)
        If cambiado = False Then Exit Sub
        ''
        Try
            Using objWriter As IO.StreamWriter = New IO.StreamWriter(OmniClassTaxonomyUser, False, System.Text.Encoding.UTF8)
                Dim sLine As String = ""
                Dim contador As Integer = 0
                For Each queCo As String In colUsuario.Keys
                    ''
                    contador += 1
                    ''
                    Dim txtLinea As String = Join(CType(colUsuario(queCo), String()), vbTab)
                    objWriter.WriteLine(txtLinea)
                Next
                ''
                objWriter.Close()
            End Using
        Catch ex As Exception
            '' No hacemos nada. colUlma está creado, pero no tendrá elementos o no los que tengan errores.
        End Try
        '' 
    End Sub
    ''
    Public Sub CopiaUsuario_PlantillaFamiliasUlma(famtempfile As String)
        If IO.File.Exists(famtempfile) = False Then Exit Sub
        ''
        'Directorio de plantillas de familias de Revit 2018:    C:\ProgramData\Autodesk\RVT 2018\Family Templates
        '' Recorreremos este directorio para sacar todos los subdirectorios y copiar en cada uno la plantilla de familias ULMA.
        Dim dirBase As String = "C:\ProgramData\Autodesk\RVT " & RevitVersion & "\Family Templates\ULMA"
        Dim vieja As String = "Item Template Formwork.rft"
        ''
        For Each queDir As String In IO.Directory.GetDirectories(dirBase, "*.*", SearchOption.TopDirectoryOnly)
            '' Fichero de plantilla a copiar en cada directorio de plantillas de familia de Revit.
            Dim fiDestino As String = IO.Path.Combine(queDir, IO.Path.GetFileName(famtempfile))
            Try
                utilesRevit.Fichero_QuitarSoloLectura(fiDestino)
                IO.File.Copy(famtempfile, fiDestino, True)
            Catch ex As Exception
                '' No hacemos nada.
            End Try
            ''
            '' Si existe la vieja, quitarla
            Dim fiVieja As String = IO.Path.Combine(queDir, vieja)
            Try
                If IO.File.Exists(fiVieja) Then IO.File.Delete(fiVieja)
            Catch ex As Exception
                '' No hacemos nada.
            End Try
        Next
    End Sub
    '
    Public Sub NombresParametrosLocalizados()
        nType = LabelUtils.GetLabelFor(BuiltInParameter.ELEM_TYPE_PARAM)
        'nSite = LabelUtils.GetLabelFor(BuiltInParameter.ELEM_CATEGORY_PARAM).ToUpper
        '' No poner en Mayúsculas. No lo coge bien.
        nManufacturer = LabelUtils.GetLabelFor(BuiltInParameter.ALL_MODEL_MANUFACTURER)
        nCategory = LabelUtils.GetLabelFor(BuiltInParameter.ELEM_CATEGORY_PARAM)
        nCount = LabelUtils.GetLabelFor(BuiltInParameter.DIM_REFERENCE_COUNT)
        nFaseCrea = LabelUtils.GetLabelFor(BuiltInParameter.PHASE_CREATED)
        nDesfase = LabelUtils.GetLabelFor(BuiltInParameter.POINT_ELEMENT_OFFSET)
        nFamily = LabelUtils.GetLabelFor(BuiltInParameter.FAMILY_NAME_PSEUDO_PARAM)
    End Sub
    '
    Public Function SeleccionarElemento(ByVal selection As Autodesk.Revit.UI.Selection.Selection, filter As ISelectionFilter, mensaje As String, Optional ulma As Boolean = True) As Reference
        If ulma = True AndAlso filter Is Nothing Then
            filter = New PickFilterULMASetAndFamilyInstanceCategory
        End If
        ''
        Dim picked As Reference = selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element, filter, mensaje)
        Return picked
    End Function
    ''
    Public Function SeleccionarElementos(ByVal selection As Autodesk.Revit.UI.Selection.Selection, filter As ISelectionFilter, mensaje As String, Optional ulma As Boolean = True) As IList(Of Reference)
        If ulma = True AndAlso filter Is Nothing Then
            filter = New PickFilterULMASetAndFamilyInstanceCategory
        End If
        ''
        Dim picked As IList(Of Reference) = Nothing
        Try
            picked = selection.PickObjects(Autodesk.Revit.UI.Selection.ObjectType.Element, filter, mensaje)
        Catch ex As Exception
            '' Hemos cancelado la selección.
        End Try
        ''
        Return picked
    End Function
    ''
    Public Function TransformPoint(point As XYZ, transform As Transform) As XYZ
        Dim x As Double = point.X
        Dim y As Double = point.Y
        Dim z As Double = point.Z

        'transform basis of the old coordinate system in the new coordinate // system
        Dim b0 As XYZ = transform.Basis(0)
        Dim b1 As XYZ = transform.Basis(1)
        Dim b2 As XYZ = transform.Basis(2)
        Dim origin As XYZ = transform.Origin

        'transform the origin of the old coordinate system in the new 
        'coordinate system
        Dim xTemp As Double = x * b0.X + y * b1.X + z * b2.X + origin.X
        Dim yTemp As Double = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y
        Dim zTemp As Double = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z

        Return New XYZ(xTemp, yTemp, zTemp)
    End Function
    ''
    Public Sub SetStatusText(text As String)
        Dim processes As Process() = Process.GetProcessesByName("Revit")
        ''
        If 0 < processes.Length Then
            Dim statusBar As IntPtr = clsAPI.FindWindowEx(processes(0).MainWindowHandle, IntPtr.Zero, "msctls_statusbar32", "")

            If statusBar <> IntPtr.Zero Then
                clsAPI.SetWindowText(statusBar, text)
            End If
        End If
    End Sub
    ''
    Public Function RibbonTabExiste(nombreRibbon As String) As adWin.RibbonTab
        Dim resultado As adWin.RibbonTab = Nothing
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON de ULMA, continuar
            If (tab.AutomationName <> nombreRibbon) Then Continue For
            ''
            resultado = ribbon.Tabs(ribbon.Tabs.IndexOf(tab))
        Next
        ''
        Return resultado
    End Function
    ''
    Public Sub RibbonTabOculta(queRibbonTab As String, Optional visible As Boolean = True)
        Dim ribbon As Autodesk.Windows.RibbonControl = Autodesk.Windows.ComponentManager.Ribbon
        ''
        For Each tab As Autodesk.Windows.RibbonTab In ribbon.Tabs
            If tab.Title = queRibbonTab Then
                ribbon.Tabs.Remove(tab)
                tab.IsVisible = visible
            End If
        Next
    End Sub
    ' Quita el Ribbon ULMA
    Public Sub RibbonTabBorra(queRibbonTab As String)
        Dim ribbon As Autodesk.Windows.RibbonControl = Nothing
        Try
            ribbon = Autodesk.Windows.ComponentManager.Ribbon
            For Each tab As Autodesk.Windows.RibbonTab In ribbon.Tabs
                If tab.Title = queRibbonTab Then
                    ribbon.Tabs.Remove(tab)
                End If
            Next
        Catch ex As Exception
            '
        End Try
    End Sub
    ''
#Region "GRUPOS"
    Public Function Grupo_DamePorNombre(ByRef queDoc As Autodesk.Revit.DB.Document, nGrupo As String) As GroupType
        Dim resultado As GroupType = Nothing
        ''
        '' ***** FilterelementCollector de todos los Grupos del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(GroupType))
        ''
        Dim groupTypes = From element In collector
                         Where element.Name = nGrupo
                         Select element

        If groupTypes IsNot Nothing AndAlso groupTypes.Count > 0 Then
            resultado = CType(groupTypes.First, GroupType)
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function Grupo_BorraNombre(ByRef queDoc As Autodesk.Revit.DB.Document, nGrupo As String) As Boolean
        Dim resultado As Boolean = False
        ''
        '' ***** FilterelementCollector de todos los Grupos del documento
        Dim collector As New FilteredElementCollector(queDoc)
        collector = collector.OfClass(GetType(GroupType))
        ''
        Dim groupTypes = From element In collector
                         Where element.Name = nGrupo
                         Select element

        If groupTypes IsNot Nothing AndAlso groupTypes.Count > 0 Then
            Dim gType As GroupType = CType(groupTypes.First, GroupType)
            Dim deletedIdSet As ICollection(Of Autodesk.Revit.DB.ElementId) = queDoc.Delete(gType.Id)
            '' Si no hemos borrado nada (False), Si hemos borrado algo (True)
            If deletedIdSet.Count = 0 Then
                resultado = False
            Else
                resultado = True
            End If
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function Grupo_ElementoEstaEnGrupo(document As Autodesk.Revit.DB.Document, element As Autodesk.Revit.DB.Element) As Boolean
        Dim resultado As Boolean = False
        If element.GroupId.Equals(ElementId.InvalidElementId) Then
            resultado = False
        Else
            resultado = True
        End If
        ''
        Return resultado
    End Function
    ''
    Public Sub Grupo_CambiaNombreGrupo(document As Autodesk.Revit.DB.Document, queGrupo As Autodesk.Revit.DB.Group, nNuevo As String)
        ' Change the default group name to a new name �MyGroup�
        queGrupo.GroupType.Name = nNuevo
    End Sub
    ''
    Public Sub Grupo_CambiaNombreBuscaPrimero(document As Autodesk.Revit.DB.Document, queGrupo As Autodesk.Revit.DB.Group, nViejo As String, nNuevo As String)

        ''
        '' ***** FilterelementCollector de todos los Grupos del documento
        Dim collector As New FilteredElementCollector(document)
        collector = collector.OfClass(GetType(Group))
        ''
        Dim groupTypes = From element In collector
                         Where CType(element, Group).GroupType.Name = nViejo
                         Select element

        If groupTypes IsNot Nothing AndAlso groupTypes.Count > 0 Then
            Dim gType As GroupType = CType(groupTypes.First, GroupType)
            gType.Name = nNuevo
        End If
    End Sub
    ''
    Public Sub Grupo_PonNombreNumerado(document As Autodesk.Revit.DB.Document, ByRef queGrupo As Group, preNombre As String)
        '' ***** FilterelementCollector de todos los Grupos del documento
        Dim collector As New FilteredElementCollector(document)
        collector = collector.OfClass(GetType(Group))
        ''
        Dim groupTypes = From element In collector
                         Where CType(element, Group).GroupType.Name.StartsWith(preNombre & "_")
                         Select CType(element, Group)

        If groupTypes IsNot Nothing AndAlso groupTypes.Count > 0 Then
            Dim mayor As Integer = 0
            For Each oGr As Group In groupTypes
                Dim numero As String = oGr.GroupType.Name.Replace(preNombre & "_", "")
                If IsNumeric(numero) AndAlso CDbl(numero) > mayor Then mayor = CInt(numero)
            Next
            queGrupo.GroupType.Name = preNombre & "_" & mayor.ToString.PadLeft(3, "0"c)
        Else
            queGrupo.GroupType.Name = preNombre & "_" & "1".PadLeft(3, "0"c)
        End If
    End Sub
    Public Sub Grupos_DocumentoDesagrupaGuarda(ByRef document As Autodesk.Revit.DB.Document)
        ''
        '' ***** FilterelementCollector de todos los Grupos del documento
        Dim collector As New FilteredElementCollector(document)
        collector = collector.OfClass(GetType(Group))
        ''
        Dim groupTypes = From element In collector
                         Select element
        ''
        If groupTypes IsNot Nothing AndAlso groupTypes.Count > 0 Then
            colGrupos = New Dictionary(Of String, ICollection(Of ElementId))
            ''
            '' Creamos e iniciamos la transacción (Agrupa el ungroup de todos los grupos)
            Dim tx As New Transaction(document)
            tx.Start("Ungroup document Groups")
            ''
            '' Recorrer todos los grupos y desagruparlos.
            '' Guardamos el nombre del grupo y la colección de elementId que tenía
            For Each oEle As Element In groupTypes
                Dim oGru As Group = CType(document.GetElement(oEle.Id), Group)
                Dim queId As ElementId = oGru.Id
                Dim clave As String = oGru.Name & "·" & queId.ToString
                ''
                If colGrupos.ContainsKey(clave) = False Then
                    colGrupos.Add(clave, oGru.UngroupMembers)
                Else
                    Console.WriteLine("El grupo " & clave & " ya existía")
                End If
            Next
            ''
            tx.Commit()
        End If
    End Sub
#End Region
    Public Function FacePickSetWorkPlaneAndPickPoint(uidoc As UIDocument, ByRef point_in_3d As XYZ) As Boolean
        point_in_3d = Nothing

        Dim doc As Document = uidoc.Document

        Dim r As Reference = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, "Please select a face to define work plane")

        Dim e As Element = doc.GetElement(r.ElementId)

        If e IsNot Nothing Then
            Dim face As Face = TryCast(e.GetGeometryObjectFromReference(r), Face)
            'Dim facePlanar As PlanarFace = TryCast(e.GetGeometryObjectFromReference(r), PlanarFace)
            'Dim faceCylindrical As CylindricalFace = TryCast(e.GetGeometryObjectFromReference(r), CylindricalFace)

            If face IsNot Nothing Then
                Dim plano As Plane = Nothing
                ''
                If TypeOf face Is PlanarFace Then
                    Dim facePlanar As PlanarFace = TryCast(e.GetGeometryObjectFromReference(r), PlanarFace)
                    plano = Plane.CreateByNormalAndOrigin(facePlanar.FaceNormal, facePlanar.Origin)
                ElseIf TypeOf face Is CylindricalFace Then
                    Dim faceCylindrical As CylindricalFace = TryCast(e.GetGeometryObjectFromReference(r), CylindricalFace)
                    Dim eje As XYZ = faceCylindrical.Axis
                    plano = Plane.CreateByNormalAndOrigin(New XYZ(0, 0, 1), faceCylindrical.Origin)
                End If
                ''
                Dim t As New Transaction(doc)

                t.Start("Temporarily set work plane" + " to pick point in 3D")

                Dim sp As SketchPlane = SketchPlane.Create(doc, plano)

                uidoc.ActiveView.SketchPlane = sp
                uidoc.ActiveView.ShowActiveWorkPlane()

                Try
                    point_in_3d = uidoc.Selection.PickPoint("Please pick a point on the plane" + " defined by the selected face")
                Catch generatedExceptionName As OperationCanceledException
                End Try

                t.RollBack()
            End If
        End If
        Return point_in_3d IsNot Nothing
    End Function
    ''
    Public Function FacePickSetWorkPlaneAndPickPointView3D(uidoc As UIDocument, ByRef point_in_3d As XYZ) As Boolean
        point_in_3d = Nothing

        Dim doc As Document = uidoc.Document

        Dim r As Reference = uidoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Face, "Please select a face to define work plane")

        Dim e As Element = doc.GetElement(r.ElementId)

        If e IsNot Nothing Then
            Dim face As Face = TryCast(e.GetGeometryObjectFromReference(r), Face)
            'Dim facePlanar As PlanarFace = TryCast(e.GetGeometryObjectFromReference(r), PlanarFace)
            'Dim faceCylindrical As CylindricalFace = TryCast(e.GetGeometryObjectFromReference(r), CylindricalFace)

            If face IsNot Nothing Then
                Dim plano As Plane = Nothing
                ''
                If TypeOf face Is PlanarFace Then
                    Dim facePlanar As PlanarFace = TryCast(e.GetGeometryObjectFromReference(r), PlanarFace)
                    plano = Plane.CreateByNormalAndOrigin(facePlanar.FaceNormal, facePlanar.Origin)
                ElseIf TypeOf face Is CylindricalFace Then
                    Dim faceCylindrical As CylindricalFace = TryCast(e.GetGeometryObjectFromReference(r), CylindricalFace)
                    Dim eje As XYZ = faceCylindrical.Axis
                    plano = Plane.CreateByNormalAndOrigin(New XYZ(0, 0, 1), faceCylindrical.Origin)
                End If
                ''


                Dim t As New Transaction(doc)

                t.Start("Temporarily set work plane" + " to pick point in 3D")
                ''
                ''
                Dim quePlano As Plane = Plane.CreateByNormalAndOrigin(doc.ActiveView.ViewDirection, doc.ActiveView.Origin)
                Dim sp As SketchPlane = SketchPlane.Create(doc, quePlano)
                ''
                doc.ActiveView.SketchPlane = sp
                doc.ActiveView.ShowActiveWorkPlane()
                ''
                Try
                    Dim snapTypes As Autodesk.Revit.UI.Selection.ObjectSnapTypes =
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Endpoints Or
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Intersections Or
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Centers Or
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Nearest Or
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Points Or
                        Autodesk.Revit.UI.Selection.ObjectSnapTypes.Points
                    point_in_3d = uidoc.Selection.PickPoint(snapTypes, "Please pick a point on the plane" + " defined by the selected face")
                    t.Commit()
                Catch generatedExceptionName As OperationCanceledException
                    t.RollBack()
                End Try
                If t.GetStatus = TransactionStatus.Started Then t.RollBack()
            End If
        End If
        Return point_in_3d IsNot Nothing
    End Function
#Region "AYUDAS PLANOS"
    ''
    ''' <summary>
    ''' Return signed distance from plane to a given point.
    ''' </summary>
    <System.Runtime.CompilerServices.Extension>
    Public Function SignedDistanceTo(plane As Plane, p As XYZ) As Double
        Debug.Assert(plane.Normal.GetLength = 1, "expected normalised plane normal")
        Dim v As XYZ = p - plane.Origin
        ''
        Return plane.Normal.DotProduct(v)
    End Function
    ''
    ''' <summary>
    ''' Project given 3D XYZ point onto plane.
    ''' </summary>
    <System.Runtime.CompilerServices.Extension>
    Public Function ProjectOnto(plane As Plane, p As XYZ) As XYZ
        Dim d As Double = plane.SignedDistanceTo(p)
        Dim q As XYZ = p + d * plane.Normal
        Debug.Assert(plane.SignedDistanceTo(q) = 0, "expected point on plane to have zero distance to plane")
        ''
        Return q
    End Function
    ''
    ''' <summary>
    ''' Project given 3D XYZ point into plane, 
    ''' returning the UV coordinates of the result 
    ''' in the local 2D plane coordinate system.
    ''' </summary>
    <System.Runtime.CompilerServices.Extension>
    Public Function ProjectInto(plane As Plane, p As XYZ) As UV
        Dim q As XYZ = plane.ProjectOnto(p)
        Dim o As XYZ = plane.Origin
        Dim d As XYZ = q - o
        Dim u As Double = d.DotProduct(plane.XVec)
        Dim v As Double = d.DotProduct(plane.YVec)
        ''
        Return New UV(u, v)
    End Function
    ''
#End Region
    Public Sub SketchPlaneActiva(ByRef uidoc As UIDocument, ByRef oSp As SketchPlane, Optional visible As Boolean = False, Optional conTr As Boolean = True)
        Dim t As Transaction = Nothing
        If conTr = True Then
            t = New Transaction(uidoc.Document)
            t.Start("Temporarily set work plane")
        End If
        ''
        Try
            uidoc.Document.ActiveView.SketchPlane = oSp
            If visible = True Then
                uidoc.Document.ActiveView.ShowActiveWorkPlane()
            Else
                uidoc.Document.ActiveView.HideActiveWorkPlane()
            End If
            ''
            If conTr = True AndAlso t IsNot Nothing Then
                uidoc.Document.Regenerate()
                uidoc.RefreshActiveView()
                t.Commit()
            End If
        Catch ex As Exception
            If conTr = True AndAlso t IsNot Nothing Then t.RollBack()
        End Try
        If (conTr = True AndAlso t IsNot Nothing) AndAlso t.GetStatus = TransactionStatus.Started Then
            t.RollBack()
        End If
    End Sub
    ''
    Private Function GetTransformToZ(v As XYZ) As Transform
        Dim t As Transform

        Dim a As Double = XYZ.BasisZ.AngleTo(v)

        If a = 0 Then
            t = Transform.Identity
        Else
            Dim axis As XYZ
            If a = Math.PI Then
                axis = XYZ.BasisX
            Else
                axis = v.CrossProduct(XYZ.BasisZ)
            End If
            ''
            t = Transform.CreateRotationAtPoint(axis, a, XYZ.Zero)
        End If
        ''
        Return t
    End Function
    ''
    Private Function NewSketchPlanePassLine(oD As Autodesk.Revit.DB.Document, line As Line) As SketchPlane
        Dim p As XYZ = line.GetEndPoint(0)
        Dim q As XYZ = line.GetEndPoint(1)
        Dim norm As XYZ
        If p.X = q.X Then
            norm = XYZ.BasisX
        ElseIf p.Y = q.Y Then
            norm = XYZ.BasisY
        Else
            norm = XYZ.BasisZ
        End If
        Dim plane As Plane = Plane.CreateByNormalAndOrigin(norm, p)

        Return SketchPlane.Create(oD, plane)
    End Function
    ''
    '' Devuelve un array(1).
    '' 0 = Autodesk.Revit.DB.XYZ (PickPoint)
    '' 1 = Angulo desde VectorX del plano al punto en Planta.
    Public Function PickPointInSketchPlanePlanta(uidoc As UIDocument, oFip As FamilyInstance) As Object()
        ''
        Dim crearModelLines As Boolean = False
        '' Poner vista Superior y seleccionar un punto.
        Dim ptins As LocationPoint = CType(oFip.Location, LocationPoint)
        Dim eyePt As XYZ = Nothing
        Dim upPt As XYZ = Nothing
        Dim forPt As XYZ = Nothing
        Dim factor As Double = 0.2
        eyePt = ptins.Point
        upPt = New XYZ(0, 1, 0)
        forPt = New XYZ(0, 0, -1)
        utilesRevit.View_Orientation3DPon(CType(evRevit.evAppUI.ActiveUIDocument.Document.ActiveView, View3D), eyePt, upPt, forPt, oFip, factor)
        uidoc.RefreshActiveView()
        ''
        ''
        'End If
        Dim datos As Object() = New Object() {Nothing, Nothing} 'datos(1)
        Dim pointInPlane As Autodesk.Revit.DB.XYZ = Nothing
        ''
        Dim point_in_3d As XYZ = Nothing
        '' Activate SketchPlane oSpSup
        SketchPlaneActiva(uidoc, oSpSup, False, True)
        'oDoc.Regenerate()
        'uidoc.RefreshActiveView()
        ''
        Dim snapTypes As ObjectSnapTypes = ObjectSnapTypes.WorkPlaneGrid Or ObjectSnapTypes.Endpoints Or ObjectSnapTypes.Intersections Or ObjectSnapTypes.Centers Or ObjectSnapTypes.Points
        Try
            point_in_3d = uidoc.Selection.PickPoint(snapTypes, "Select Point to Rotate.")
        Catch ex As Exception
            ''
        End Try
        If point_in_3d Is Nothing Then
            GoTo FINAL
        Else
            datos(0) = point_in_3d
        End If
        ''
        Dim t As New Transaction(uidoc.Document)
        t.Start("Select Point in plane SUP.")
        ''
        ''
        'Dim vector1 As Line = Nothing
        'Dim vector2 As Line = Nothing
        Try
            ''
            'pointInPlane = uidoc.Selection.PickPoint("Please pick a point on the plane SUP.")
            ''
            Dim anguloR As Double = 0       '' En Radianes
            Dim angulo As Double = 0        '' En Grados
            Dim angulofin As Double = 0     '' En Grados
            Dim distancia As Double = 0
            Dim elev As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, oFip, BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM) ' Elevación.
            ''
            '' ROTATION
            Dim oPl As Plane = oSpSup.GetPlane
            'Dim pt As XYZ = ptins.Point 'ptins.Point.Add(New XYZ(0, 0, CDbl(elev)))
            Dim pt As XYZ = oPl.ProjectOnto(ptins.Point)
            distancia = pt.DistanceTo(point_in_3d)
            ''
            Dim origin As XYZ = New XYZ(pt.X, pt.Y, 0)
            Dim ptDirX As XYZ = New XYZ(pt.X + distancia, pt.Y, origin.Z)
            Dim ptDirY As XYZ = New XYZ(pt.X, pt.Y + distancia, origin.Z)
            Dim pt2 As XYZ = New XYZ(point_in_3d.X, point_in_3d.Y, origin.Z)
            Dim normal As XYZ = New XYZ(origin.X, origin.Y, distancia)
            ''
            Dim oLineX As Line = Line.CreateBound(origin, ptDirX)
            Dim oLineY As Line = Line.CreateBound(origin, ptDirY)
            Dim oLineRotate As Line = Line.CreateBound(origin, pt2)
            ''
            ''
            If crearModelLines = True Then
                'Dim geomPlane As Plane = Plane.CreateByNormalAndOrigin(normal, origin)
                Dim geomPlane As Plane = Plane.CreateByOriginAndBasis(origin, oLineX.Direction.Normalize, oLineY.Direction.Normalize)
                Dim sketch As SketchPlane = SketchPlane.Create(evRevit.evAppUI.ActiveUIDocument.Document, geomPlane)
                ' Create a ModelLine element using the created geometry line and sketch plane
                Dim line__1 As ModelLine = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.Create.NewModelCurve(oLineX, sketch), ModelLine)
                Dim line__2 As ModelLine = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.Create.NewModelCurve(oLineRotate, sketch), ModelLine)
            End If

            'Dim oMc1 As ModelCurve = oDoc.Create.NewModelCurve(vector1, oSpSup)
            'Dim oMc2 As ModelCurve = oDoc.Create.NewModelCurve(vector2, oSpSup)
            ''
            'anguloR = vector1.Direction.AngleOnPlaneTo(vector2.Direction, oSp.GetPlane.Normal)
            'anguloR = vector1.Direction.AngleOnPlaneTo(vector2.Direction, oSpBase.GetPlane.Normal)
            'anguloR = vector1.Direction.AngleOnPlaneTo(vector2.Direction, oSpSup.GetPlane.Normal)
            'anguloR = oSpSup.GetPlane.XVec.AngleOnPlaneTo(vector2.Direction, oSpSup.GetPlane.Normal)
            'anguloR = oSpSup.GetPlane.Normal.AngleTo(vector2.Direction)
            anguloR = oLineX.Direction.AngleTo(oLineRotate.Direction)
            If pt2.Y < origin.Y Then
                anguloR *= -1
            End If
            angulo = U_DameGrados_DesdeRadianes(anguloR, False)
            ''
            angulofin = angulo
            If angulo > 0 Then
                If angulo > 270 Then
                    'angulofin = 360 - angulo
                ElseIf angulo > 180 Then
                    'angulofin = 270 - angulo
                ElseIf angulo > 90 Then
                    'angulofin = 180 - angulo
                ElseIf angulo > 0 Then
                    'angulofin = 90 - angulo
                End If
            ElseIf angulo < 0 Then
                If angulo < -270 Then
                    'angulofin = angulo + 270
                ElseIf angulo < -180 Then
                    'angulofin = angulo + 180
                ElseIf angulo < -90 Then
                    'angulofin = angulo + 90
                End If
            End If
            datos(1) = angulofin    ' Math.Abs(angulofin)
            t.Commit()
        Catch generatedExceptionName As OperationCanceledException
            t.RollBack()
        End Try
        ''
        If t IsNot Nothing AndAlso t.GetStatus = TransactionStatus.Started Then t.RollBack()
        ''
FINAL:
        Return datos
    End Function
    ''
    '' Devuelve un array(1).
    '' 0 = Autodesk.Revit.DB.XYZ (PickPoint)
    '' 1 = Angulo desde VectorX del plano al punto en Alzado
    Public Function PickPointInSketchPlaneAlzado(uidoc As UIDocument, oFip As FamilyInstance) As Object()
        ''
        Dim crearModelLines As Boolean = False
        '' Poner vista Alzado y seleccionar un punto.
        Dim ptins As LocationPoint = CType(oFip.Location, LocationPoint)
        Dim eyePt As XYZ = Nothing
        Dim upPt As XYZ = Nothing
        Dim forPt As XYZ = Nothing
        Dim factor As Double = 0.2
        eyePt = ptins.Point
        upPt = New XYZ(0, 0, 1)
        forPt = New XYZ(0, 1, 0)
        utilesRevit.View_Orientation3DPon(CType(evRevit.evAppUI.ActiveUIDocument.Document.ActiveView, View3D), eyePt, upPt, forPt, oFip, factor)
        uidoc.RefreshActiveView()
        ''
        ''
        Dim datos As Object() = New Object() {Nothing, Nothing} 'datos(1)
        Dim pointInPlane As Autodesk.Revit.DB.XYZ = Nothing
        ''
        Dim point_in_3d As XYZ = Nothing
        '' Activate SketchPlane oSpSup
        SketchPlaneActiva(uidoc, oSpFro, False, True)
        'oDoc.Regenerate()
        'uidoc.RefreshActiveView()
        ''
        Dim snapTypes As ObjectSnapTypes = ObjectSnapTypes.WorkPlaneGrid Or ObjectSnapTypes.Endpoints Or ObjectSnapTypes.Intersections Or ObjectSnapTypes.Centers Or ObjectSnapTypes.Points
        Try
            point_in_3d = uidoc.Selection.PickPoint(snapTypes, "Select Point to Rotate.")
        Catch ex As Exception
            ''
        End Try
        ''
        If point_in_3d Is Nothing Then
            GoTo FINAL
        Else
            datos(0) = point_in_3d
        End If
        ''
        Dim t As New Transaction(uidoc.Document)
        t.Start("Select Point in plane FRONT")
        ''
        Try
            ''
            'pointInPlane = uidoc.Selection.PickPoint("Please pick a point on the plane SUP.")
            ''
            Dim anguloR As Double = 0       '' En Radianes
            Dim angulo As Double = 0        '' En Grados
            Dim angulofin As Double = 0     '' En Grados
            Dim distancia As Double = 0
            Dim elev As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, oFip, BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM) ' Elevación.
            ''
            '' ROTATION
            Dim oPl As Plane = oSpSup.GetPlane
            'Dim pt As XYZ = New XYZ(ptins.Point.X, ptins.Point.Y, ptins.Point.Z + CDbl(elev))
            Dim pt As XYZ = oPl.ProjectOnto(ptins.Point)
            distancia = pt.DistanceTo(point_in_3d)
            ''
            Dim origin As XYZ = New XYZ(pt.X, pt.Y, pt.Z)
            Dim ptDirX As XYZ = New XYZ(origin.X + distancia, origin.Y, origin.Z)
            Dim ptDirY As XYZ = New XYZ(origin.X, origin.Y, origin.Z + distancia)
            Dim pt2 As XYZ = New XYZ(point_in_3d.X, origin.Y, point_in_3d.Z)
            Dim normal As XYZ = New XYZ(origin.X, -distancia, origin.Z)
            ''
            Dim oLineX As Line = Line.CreateBound(origin, ptDirX)
            Dim oLineY As Line = Line.CreateBound(origin, ptDirY)
            Dim oLineRotate As Line = Line.CreateBound(origin, pt2)
            ''
            If crearModelLines = True Then
                'Dim geomPlane As Plane = Plane.CreateByNormalAndOrigin(normal, origin)
                Dim geomPlane As Plane = Plane.CreateByOriginAndBasis(origin, oLineX.Direction.Normalize, oLineY.Direction.Normalize)
                Dim sketch As SketchPlane = SketchPlane.Create(evRevit.evAppUI.ActiveUIDocument.Document, geomPlane)
                ' Create a ModelLine element using the created geometry line and sketch plane
                Dim line__1 As ModelLine = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.Create.NewModelCurve(oLineX, sketch), ModelLine)
                Dim line__2 As ModelLine = TryCast(evRevit.evAppUI.ActiveUIDocument.Document.Create.NewModelCurve(oLineRotate, sketch), ModelLine)
            End If
            ''
            anguloR = oLineX.Direction.AngleTo(oLineRotate.Direction)
            If pt2.Z < origin.Z Then
                anguloR *= -1
            End If
            angulo = U_DameGrados_DesdeRadianes(anguloR, False)
            ''
            angulofin = angulo
            If angulo > 0 Then
                If angulo > 270 Then
                    'angulofin = 360 - angulo
                ElseIf angulo > 180 Then
                    'angulofin = 270 - angulo
                ElseIf angulo > 90 Then
                    'angulofin = 180 - angulo
                ElseIf angulo > 0 Then
                    'angulofin = 90 - angulo
                End If
            ElseIf angulo < 0 Then
                If angulo < -270 Then
                    'angulofin = angulo + 270
                ElseIf angulo < -180 Then
                    'angulofin = angulo + 180
                ElseIf angulo < -90 Then
                    'angulofin = angulo + 90
                End If
            End If
            datos(1) = angulofin    ' Math.Abs(angulofin)
            t.Commit()
        Catch generatedExceptionName As OperationCanceledException
            t.RollBack()
        End Try
        ''
        If t IsNot Nothing AndAlso t.GetStatus = TransactionStatus.Started Then t.RollBack()
        ''
FINAL:
        Return datos
    End Function
    ''
    '' 0 = Solo Rotación
    '' 1 = Solo Inclinacion
    '' 2 = Rotación + Inclinación
    Public Function TaskDialogRotacion1() As Integer
        Dim resultado As Integer = -1

        '' TaskDialog para ver que tipo de giro haremos (sólo rotación, sólo inclinación o ambos)
        Dim td As New TaskDialog("Rotation and/or Inclination...")
        td.MainIcon = TaskDialogIcon.TaskDialogIconWarning
        'td.Title = ""
        td.MainInstruction = "Select Rotation or Inclination or Rotation+Inclinarion"
        '' Add commmandLink options to task dialog
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Only Rotation")
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Only Inclination")
        td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Rotation and Inclination")
        '' Set common buttons and default button. If no CommonButton or CommandLink is added,
        '' task dialog will show a Close button by default
        td.CommonButtons = TaskDialogCommonButtons.Close
        td.DefaultButton = TaskDialogResult.Close
        ''
        Dim tResult As TaskDialogResult = td.Show
        ''
        If (TaskDialogResult.CommandLink1 = tResult) Then
            resultado = 0
        ElseIf (TaskDialogResult.CommandLink2 = tResult) Then
            resultado = 1
        ElseIf TaskDialogResult.CommandLink3 = tResult Then
            resultado = 2
        Else
            resultado = -1
        End If
        ''
        Return resultado
    End Function
    '
    Public Function DameTrans(queTrans As Transform, tipot As tipotrans) As Double

        Dim resultado As Double = 0
        'XYZ vectorTran = transform.OfVector (transform.BasisX);
        'double d1 = transform.BasisX.AngleOnPlaneTo (vectorTran, transform.BasisZ); // radianes
        'd1 = d1 * (180 / Math .PI); // grados 
        ''
        Dim origen As XYZ = queTrans.Origin         '' Origen de las transformaciones
        'Dim oriBase As XYZ = New XYZ(0, 0, 0)       '' Origen absoluto. Pará cálculos
        ''
        'Dim finXBase As XYZ = XYZ.BasisX
        Dim finXBaseT As XYZ = queTrans.BasisX   'queTrans.OfVector(finXBase)  ' queTrans.OfPoint(finXBase)
        Dim finX As XYZ = New XYZ(finXBaseT.X, finXBaseT.Y, 0)
        ''
        'Dim finYBase As XYZ = XYZ.BasisY
        Dim finYBaseT As XYZ = queTrans.BasisY 'queTrans.OfPoint(finYBase)   ' queTrans.OfVector(finYBase)
        Dim finY As XYZ = New XYZ(0, finYBaseT.Y, finYBaseT.Z)
        ''
        'Dim finZBase As XYZ = XYZ.BasisZ
        Dim finZBaseT As XYZ = queTrans.BasisZ  'queTrans.OfVector(finZBase)  ' queTrans.OfPoint(finZBase)
        Dim finZ As XYZ = New XYZ(finZBaseT.X, 0, finZBaseT.Z)
        ''
        Dim finRot As Double = XYZ.BasisX.AngleOnPlaneTo(queTrans.BasisX, XYZ.BasisZ)
        Dim finInc As Double = queTrans.Inverse.BasisZ.AngleOnPlaneTo(queTrans.BasisZ, queTrans.Inverse.BasisY.Negate)
        '
        Dim finRotD As Double = U_DameGrados_DesdeRadianes(finRot, False)
        Dim finIncD As Double = U_DameGrados_DesdeRadianes(finInc, False)
        ''
        'MsgBox("Rotacion = " & rotation & vbCrLf & "Inclination = " & inclination)
        Dim fullturn As Double = U_DameRadianes_DesdeGrados(360, False)
        If tipot = tipotrans.rotation Then
            If finRot >= fullturn Then finRot -= fullturn
            resultado = finRot  'rotation
        ElseIf tipot = tipotrans.inclination Then
            If finInc >= fullturn Then finInc -= fullturn
            resultado = finInc  ' inclination
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function DameTrans1(queTrans As Transform, tipot As tipotrans) As Double
        Dim resultado As Double = 0
        'XYZ vectorTran = transform.OfVector (transform.BasisX);
        'double d1 = transform.BasisX.AngleOnPlaneTo (vectorTran, transform.BasisZ); // radianes
        'd1 = d1 * (180 / Math .PI); // grados 
        ''
        Dim origen As XYZ = queTrans.Origin         '' Origen de las transformaciones
        '
        Dim vectorTran As XYZ = queTrans.OfVector(queTrans.BasisX)
        Dim finRot As Double = queTrans.BasisX.AngleOnPlaneTo(vectorTran, queTrans.BasisZ)
        ''
        vectorTran = queTrans.OfVector(queTrans.BasisZ)
        Dim finInc As Double = XYZ.BasisZ.AngleOnPlaneTo(queTrans.BasisZ, queTrans.BasisY.Negate)
        '
        Dim finRotD As Double = U_DameGrados_DesdeRadianes(finRot, True)
        Dim finIncD As Double = U_DameGrados_DesdeRadianes(finInc, True)
        ''
        'utilesrevit.UnidadesRadiansToDegrees(Autodesk.Revit.DB.XYZ.BasisZ.AngleOnPlaneTo(queTrans.BasisZ, queTrans.BasisY.Negate))
        'MsgBox("Rotacion = " & rotation & vbCrLf & "Inclination = " & inclination)
        Dim fullturn As Double = U_DameRadianes_DesdeGrados(360, False)
        If tipot = tipotrans.rotation Then
            If finRot >= fullturn Then finRot -= fullturn
            resultado = finRot  'rotation
        ElseIf tipot = tipotrans.inclination Then
            If finInc >= fullturn Then finInc -= fullturn
            resultado = finInc  ' inclination
        End If
        ''
        Return resultado
    End Function
    ''
    Public Enum tipotrans
        rotation
        inclination
    End Enum
    '
    '' Rellena las variables "modVar.angFin", "modVar.incliFin" y "modVar.elevFin" en unidades idioma.
    Public Sub TransformPon(queFam As FamilyInstance)
        '' Transformación actual del FamilyInstance.
        'Dim oFiTrans As Transform = queFam.GetTotalTransform
        ''' Cogemos las Rotación e Inclinación actual
        'modVar.angFinRad = DameTrans(oFiTrans, tipotrans.rotation)       '' Rotacion en radians
        'modVar.incliFinRad = DameTrans(oFiTrans, tipotrans.inclination)    '' Inclinacion en radians
        '' Guardar (En grados) y como cadena en variables globales.
        'modVar.angFin = U_DameGrados_DesdeRadianes(modVar.angFinRad, False).ToString
        'If (modVar.angFin = "" OrElse (IsNumeric(modVar.angFin) AndAlso Math.Round(CDbl(modVar.angFin)) >= 360)) Then modVar.angFin = "0"
        ''
        'modVar.incliFin = U_DameGrados_DesdeRadianes(modVar.incliFinRad, False).ToString
        'If (modVar.incliFin = "" OrElse (IsNumeric(modVar.incliFin) AndAlso Math.Round(CDbl(modVar.incliFin)) >= 360)) Then modVar.incliFin = "0"
        ''
        'modVar.elevFin = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, queFam, BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
        'If modVar.elevFin = "" Then modVar.elevFin = "0"
    End Sub
    '
    ' Crear planos de trabajo
    Public Sub ProxyCreaPlanosTrabajoPoint(ByRef oDoc As Document, ByRef oFiProx As FamilyInstance)

        ''
        '' Guardar el SketchPlane actual
        ' Dim sp As SketchPlane = Nothing
        'If SketchPlane3D Is Nothing Then SketchPlane3D = oDoc.ActiveView.SketchPlane
        ''
        ''
        Dim oTx As Transaction = Nothing
        oTx = New Transaction(oDoc, "Create Work planes in FamilyInstance")
        oTx.Start()
        Try
            Dim ptIns As XYZ = CType(oFiProx.Location, LocationPoint).Point
            'Dim elev As Double = CDbl(ParametroElementLeeBuiltInParameter(oDoc, oFiProx, BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM))
            ''
            Dim ptOri As XYZ = New XYZ(ptIns.X, ptIns.Y, ptIns.Z)   ' + elev)
            Dim lNormalPl As Line = Line.CreateBound(ptOri, ptOri.Add(New XYZ(0, 0, 1)))
            ''
            Dim oPlane As Plane = Nothing
            '' Face Sup Proxy.
            oPlane = Plane.CreateByNormalAndOrigin(lNormalPl.Direction, ptOri)
            oSpSup = SketchPlane.Create(oDoc, oPlane)
            Try
                oSpSup.Name = "FaceSupProxy"
            Catch ex As Exception
                ''
            End Try
            ''
            '' Face Front Proxy.
            Dim lNormalFr As Line = Line.CreateBound(ptOri, oPlane.YVec.Negate)
            oPlane = Plane.CreateByNormalAndOrigin(lNormalFr.Direction, ptOri)
            oSpFro = SketchPlane.Create(oDoc, oPlane)
            Try
                oSpFro.Name = "FaceFroProxy"
            Catch ex As Exception
                ''
            End Try
            ''
            oTx.Commit()
        Catch ex As Exception
            oTx.RollBack()
            MsgBox(ex.ToString)
        End Try
        ''
        If oTx IsNot Nothing AndAlso oTx.GetStatus = TransactionStatus.Started Then oTx.RollBack()
    End Sub

    ''
    Public Function FamilyInstanceDameRotación(fi As FamilyInstance) As Double
        Dim resultado As Double = 0
        ''
        Dim trf As Transform = fi.GetTransform
        Dim viewDirection As XYZ = evRevit.evAppUI.ActiveUIDocument.ActiveView.ViewDirection
        Dim rightDirection As XYZ = evRevit.evAppUI.ActiveUIDocument.ActiveView.RightDirection
        Dim upDirection As XYZ = evRevit.evAppUI.ActiveUIDocument.ActiveView.UpDirection
        ''
        TaskDialog.Show("Trf", String.Format("X{0} Y{1} Z{2}\nX{3} Y{4} Z{5}",
                trf.BasisX, trf.BasisY, trf.BasisZ,
                rightDirection, upDirection, viewDirection))


        TaskDialog.Show("Rotation Angle", (trf.BasisX.AngleOnPlaneTo(rightDirection, viewDirection) * 180 / Math.PI).ToString)

        resultado = trf.BasisX.AngleOnPlaneTo(rightDirection, viewDirection) * 180 / Math.PI

        Return resultado
    End Function
#Region "Family"
    '' Inserta familias genéricas "GenericModel" en un punto
    Public Function FamilyInstance_InsertaEnPunto(doc As Document, nombreSymbol As String, quePt As XYZ) As FamilyInstance
        Dim famIns As FamilyInstance = Nothing
        Dim collector As New FilteredElementCollector(doc)
        collector.OfClass(GetType(FamilySymbol))
        Dim fs As FamilySymbol = CType(
            collector.FirstOrDefault(Function(q) _
                                     CType(q, FamilySymbol).Family.FamilyCategory.Id.IntegerValue.Equals(BuiltInCategory.OST_GenericModel) And
                                     CType(q, FamilySymbol).FamilyName = nombreSymbol),
                                 FamilySymbol)

        If fs Is Nothing Then
            Return Nothing
            Exit Function
        End If
        '' Activar el FamilySymbol
        If fs.IsActive = False Then fs.Activate()
        '' Insertarlo
        famIns = doc.Create.NewFamilyInstance(quePt, fs, [Structure].StructuralType.NonStructural)
        '' Insertar en punto (Basado en Punto) FamilyPlacementType.WorkPlaneBased()
        'oFiProx = oDoc.Create.NewFamilyInstance(ptIns, oFipProxSym, [Structure].StructuralType.NonStructural)
        ''
        Return famIns
    End Function
    ''
    '' Inserta familias genéricas "GenericModel" en una cara
    Public Function FamilyInstance_InsertaEnCara(doc As Document, nombreSymbol As String, quePt As XYZ, queFace As Face) As FamilyInstance
        Dim famIns As FamilyInstance = Nothing
        Dim collector As New FilteredElementCollector(doc)
        collector.OfClass(GetType(FamilySymbol))
        Dim fs As FamilySymbol = CType(
            collector.FirstOrDefault(Function(q) _
                                     CType(q, FamilySymbol).Family.FamilyCategory.Id.IntegerValue.Equals(BuiltInCategory.OST_GenericModel) And
                                     CType(q, FamilySymbol).FamilyName = nombreSymbol),
                                 FamilySymbol)

        If fs Is Nothing Then
            Return Nothing
            Exit Function
        End If
        '' Activar el FamilySymbol
        If fs.IsActive = False Then fs.Activate()
        '' Insertarlo
        famIns = doc.Create.NewFamilyInstance(queFace.Reference, quePt, queFace.ComputeNormal(New UV(0, 0)), fs)
        ''
        Return famIns
        '' Insertar en Face (Basado en Punto) FamilyPlacementType.WorkPlaneBased()
        'oFiFin = oDoc.Create.NewFamilyInstance(oPlanarFace.Reference, ptIns, dirX, symFin)
        ''
        '' Insertar en Face (Basado en Curva) FamilyPlacementType.CurveBased
        'oFiFin = oDoc.Create.NewFamilyInstance(oPlanarFace.Reference, oLine, symFin)    '' No pone origen
    End Function

#End Region
    ''
    '' Devolvemos el resultado en milímetros por defecto si no especificamos FALSE
    Public Function PuntoLocalToWorld(quePunto As XYZ, Optional enmilimetros As Boolean = True) As XYZ
        Dim resultado As XYZ = Nothing
        Dim document As Document = evRevit.evAppUI.ActiveUIDocument.Document
        ''
        Dim oProjectLocation As ProjectLocation = document.ActiveProjectLocation
        Dim oTransform As Transform = oProjectLocation.GetTotalTransform.Inverse
        Dim oSurveyPoint As XYZ = oTransform.OfPoint(quePunto)
        ''
        If enmilimetros Then
            resultado = New XYZ(
                CDbl(Unidades_DameMilimetros(oSurveyPoint.X.ToString, 10)),
                CDbl(Unidades_DameMilimetros(oSurveyPoint.Y.ToString, 10)),
                CDbl(Unidades_DameMilimetros(oSurveyPoint.Z.ToString, 10)))
            'resultado = New XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z) * 304.8
        Else

            resultado = New XYZ(oSurveyPoint.X, oSurveyPoint.Y, oSurveyPoint.Z)
        End If
        ''
        Return resultado
    End Function
    ''
    '' Devolvemos el resultado en milímetros por defecto si no especificamos FALSE
    Public Function PuntoWorldToLocal(quePunto As XYZ, Optional enmilimetros As Boolean = True) As XYZ
        Dim resultado As XYZ = Nothing
        Dim document As Document = evRevit.evAppUI.ActiveUIDocument.Document
        ''
        Dim filter As ElementCategoryFilter = New ElementCategoryFilter(BuiltInCategory.OST_ProjectBasePoint)
        Dim collector As FilteredElementCollector = New FilteredElementCollector(document)
        Dim oProjectBasePoints As IList(Of Element) = collector.WherePasses(filter).ToElements
        ''
        Dim oProyectBasePoint As Element = Nothing
        For Each bp As Element In oProjectBasePoints
            oProyectBasePoint = bp
            Exit For
        Next
        ''
        Dim x As Double = oProyectBasePoint.Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).AsDouble
        Dim y As Double = oProyectBasePoint.Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).AsDouble
        Dim z As Double = oProyectBasePoint.Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM).AsDouble
        Dim r As Double = oProyectBasePoint.Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble
        ''
        Dim ptFin As XYZ = New XYZ(
                quePunto.X * Math.Cos(r) - quePunto.Y * Math.Sin(r),
                quePunto.X * Math.Sin(r) + quePunto.Y * Math.Cos(r),
                0)
        ''
        If enmilimetros Then
            resultado = New XYZ(
                CDbl(Unidades_DameMilimetros((quePunto.X * Math.Cos(r) - quePunto.Y * Math.Sin(r)).ToString, 10)),
                CDbl(Unidades_DameMilimetros((quePunto.X * Math.Sin(r) + quePunto.Y * Math.Cos(r)).ToString, 10)),
                0)
            'resultado = New XYZ( _
            '    quePunto.X * Math.Cos(r) - quePunto.Y * Math.Sin(r),
            '    quePunto.X * Math.Sin(r) + quePunto.Y * Math.Cos(r),
            '    0) * 304.8      '' De feet a milimetros es multiplicar por 304.8
        Else
            resultado = New XYZ(
                quePunto.X * Math.Cos(r) - quePunto.Y * Math.Sin(r),
                quePunto.X * Math.Sin(r) + quePunto.Y * Math.Cos(r),
                0)
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function ImagenElementoDame(queDoc As Document, queEl As Element) As System.Drawing.Bitmap
        Dim resultado As System.Drawing.Bitmap = Nothing
        ''
        Try
            If TypeOf queEl Is Family Then
                'queEl = CType(oDoc.GetElement(CType(queEl, Family).GetFamilySymbolIds.FirstOrDefault), FamilySymbol)
                For Each queid As ElementId In CType(queEl, Family).GetFamilySymbolIds
                    resultado = CType(evRevit.evAppUI.ActiveUIDocument.Document.GetElement(queid), FamilySymbol).GetPreviewImage(New System.Drawing.Size(200, 200))
                    If resultado IsNot Nothing Then Exit For
                Next
            ElseIf TypeOf queEl Is FamilySymbol Then
                resultado = CType(queEl, FamilySymbol).GetPreviewImage(New System.Drawing.Size(200, 200))
            ElseIf TypeOf queEl Is FamilyInstance Then
                resultado = ImagenElementoDame(queDoc, CType(queEl, FamilyInstance).Symbol)
            Else
                resultado = CType(queEl, ElementType).GetPreviewImage(New System.Drawing.Size(200, 200))
            End If
        Catch ex As Exception
            resultado = Nothing
        End Try
        ''
        Return resultado
    End Function
    ''
    Public Sub ImagenElementoGuarda(queDoc As Document, queEl As Element, queFi As String)
        Dim bitm As System.Drawing.Bitmap = ImagenElementoDame(queDoc, queEl)
        ''
        If bitm Is Nothing Then Exit Sub
        ''
        If IO.Path.GetExtension(queFi).ToLower <> ".jpg" Then
            queFi = IO.Path.ChangeExtension(queFi, ".jpg")
        End If
        If bitm IsNot Nothing Then
            bitm.Save(queFi, System.Drawing.Imaging.ImageFormat.Jpeg)
        End If
    End Sub
    ''
    Public Sub ImagenDocumentoGuarda(queDoc As Document, Optional dirimgfinal As String = "")
        Dim oV3d As Autodesk.Revit.DB.View3D = utilesRevit.View3D_Dame(queDoc)
        If oV3d IsNot Nothing Then
            Dim fiPrev As String = IO.Path.ChangeExtension(queDoc.PathName, ".jpg")
            If dirimgfinal <> "" AndAlso IO.Directory.Exists(dirimgfinal) Then
                fiPrev = IO.Path.Combine(dirimgfinal, IO.Path.GetFileName(fiPrev))
            End If
            Dim fiGenSuf As String = ImageExportOptions.GetFileName(queDoc, oV3d.Id)    '' Esto es el texto que se ha añadido al nombre del fichero
            Dim fiGen As String = IO.Path.Combine(
                IO.Path.GetDirectoryName(queDoc.PathName),
                IO.Path.GetFileNameWithoutExtension(queDoc.PathName) & fiGenSuf & ".jpg")
            '' Borrar antes los ficheros, si existen.
            If IO.File.Exists(fiPrev) Then
                Try
                    IO.File.Delete(fiPrev)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
            End If
            ''
            If IO.File.Exists(fiGen) Then
                Try
                    IO.File.Delete(fiGen)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
            End If
            '' Exportar vista 3D con opciones.
            Using optIE As New Autodesk.Revit.DB.ImageExportOptions
                optIE.FilePath = fiPrev
                optIE.ExportRange = ExportRange.SetOfViews
                Dim oLi As New List(Of ElementId) : oLi.Add(oV3d.Id)
                optIE.SetViewsAndSheets(oLi)
                optIE.ZoomType = ZoomFitType.FitToPage
                'optIE.FitDirection = Autodesk.Revit.DB.FitDirectionType.Vertical
                'optIE.ViewName = ""
                Try
                    queDoc.ExportImage(optIE)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
                '' Renombrar el fichero generado para quitar el nombre de la vista.
                If IO.File.Exists(fiGen) Then
                    IO.File.Copy(fiGen, fiPrev)
                    Try
                        IO.File.Delete(fiGen)
                    Catch ex As Exception
                        Console.WriteLine(ex.ToString)
                    End Try
                End If
            End Using
        End If
        oV3d = Nothing
    End Sub
    '' Devuelve un bitmap de la imagen previa del documento.
    Public Function ImagenDocumentoDame(queDoc As Document) As System.Drawing.Bitmap
        Dim resultado As System.Drawing.Bitmap = Nothing
        ''
        Dim oV3d As Autodesk.Revit.DB.View3D = utilesRevit.View3D_Dame(queDoc)
        If oV3d IsNot Nothing Then
            'Dim fiPrev As String = IO.Path.ChangeExtension(queDoc.PathName, ".jpg")
            Dim fiGenSuf As String = ImageExportOptions.GetFileName(queDoc, oV3d.Id)    '' Esto es el texto que se ha añadido al nombre del fichero
            Dim fiGen As String = IO.Path.Combine(
                IO.Path.GetTempPath,
                IO.Path.GetFileNameWithoutExtension(queDoc.PathName) & fiGenSuf & ".jpg")
            '' Borrar antes el fichero que se generará, si existe.
            If IO.File.Exists(fiGen) Then
                Try
                    IO.File.Delete(fiGen)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
            End If
            '' Exportar vista 3D con opciones.
            Using optIE As New Autodesk.Revit.DB.ImageExportOptions
                optIE.FilePath = fiGen
                optIE.ExportRange = ExportRange.SetOfViews
                Dim oLi As New List(Of ElementId) : oLi.Add(oV3d.Id)
                optIE.SetViewsAndSheets(oLi)
                optIE.ZoomType = ZoomFitType.FitToPage
                'optIE.FitDirection = Autodesk.Revit.DB.FitDirectionType.Vertical
                'optIE.ViewName = ""
                Try
                    queDoc.ExportImage(optIE)
                Catch ex As Exception
                    Console.WriteLine(ex.ToString)
                End Try
                '' Renombrar el fichero generado para quitar el nombre de la vista.
                If IO.File.Exists(fiGen) Then
                    Try
                        'resultado = New System.Drawing.Bitmap(System.Drawing.Bitmap.FromFile(fiGen))
                        resultado = New System.Drawing.Bitmap(fiGen)
                        IO.File.Delete(fiGen)
                    Catch ex As Exception
                        Console.WriteLine(ex.ToString)
                    End Try
                End If
            End Using
        End If
        oV3d = Nothing
        ''
        Return resultado
    End Function
    ''
    Public Function ImagenPreviaFicheroDame(queFi As String) As System.Drawing.Image
        '' Generar imagen previa y devuelve imagen previa
        Using clsS As New RevitPreview.StructuredS.Storage(queFi)
            Dim queImagen As System.Drawing.Image = Nothing
            queImagen = clsS.ThumbnailImage.GetPreviewAsImage
            ''
            If queImagen IsNot Nothing Then
                'Return New System.Drawing.Bitmap(queImagen)
                Return CType(queImagen.Clone, System.Drawing.Image)
            Else
                Return Nothing
            End If
            queImagen = Nothing
        End Using
    End Function
    '
    Public Function XYZ_DameEnTexto(queXYZ As XYZ, Optional conparentesis As Boolean = True) As String
        Dim parIni As String = "("
        Dim parFin As String = ")"
        If conparentesis Then
            Return parIni & queXYZ.X & ", " & queXYZ.Y & ", " & queXYZ.Z & parFin
        Else
            Return queXYZ.X & ", " & queXYZ.Y & ", " & queXYZ.Z
        End If
    End Function
    '
    Public Function FamilyInstance_DameLineVisible(fi As FamilyInstance) As List(Of Line)
        Dim resultado As New List(Of Line)
        Dim transforms = fi.GetTransform()
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = False
        options.ComputeReferences = True
        options.View = evRevit.evAppUI.ActiveUIDocument.Document.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        Dim oGeoEl As GeometryElement = fi.Geometry(options)
        '
        If oGeoEl IsNot Nothing Then
            For Each oGeo As GeometryObject In oGeoEl
                If TypeOf oGeo Is GeometryInstance Then
                    Dim gi As GeometryInstance = TryCast(oGeo, GeometryInstance)
                    For Each goSymbol As GeometryObject In gi.GetInstanceGeometry
                        If TypeOf goSymbol Is Line Then
                            Dim line As Line = TryCast(goSymbol, Line)
                            If line.Visibility = Visibility.Visible Then
                                resultado.Add(line)
                            End If
                        End If
                    Next
                End If
            Next
        End If
        '
        Return resultado
    End Function
    Public Function FamilyInstance_DamePlanosReferencia(doc As Document, fi As FamilyInstance) As List(Of Line)
        Dim resultado As New List(Of Line)
        'Dim fi = TryCast(oDoc.GetElement(ElementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        Dim r = fi.GetFamilyPointPlacementReferences()
        '
        For Each oGeo As GeometryObject In fi.Geometry(options)
            data += " - " & oGeo.[GetType]().ToString() + Environment.NewLine

            If TypeOf oGeo Is GeometryInstance Then
                Dim gi As GeometryInstance = TryCast(oGeo, GeometryInstance)

                For Each goSymbol As GeometryObject In gi.GetSymbolGeometry()
                    data += " -- " & goSymbol.[GetType]().ToString() + Environment.NewLine

                    If TypeOf goSymbol Is Line Then
                        Dim line As Line = TryCast(goSymbol, Line)
                        data += " --- " & line.[GetType]().ToString() & " (" + line.Origin.ToString() & ")" & Environment.NewLine
                        'Dim linefin As Line = Line.CreateBound(
                        '    transforms.OfPoint(line.GetEndPoint(0)),
                        '    transforms.OfPoint(line.GetEndPoint(1))
                        '    )
                        'resultado.Add(linefin)
                        resultado.Add(line)
                    End If
                Next
            End If
        Next

        'TaskDialog.Show("data", data)
        '
        Return resultado
    End Function
    '
    Public Function FamilyInstance_DameFacesSymbolGeometry(ByVal doc As Document, fi As FamilyInstance) As List(Of Face)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        '
        Dim resultado As New List(Of Face)

        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetSymbolGeometry
                    If TypeOf go1 Is Solid Then
                        Dim oSolid As Solid = CType(go1, Solid)
                        If oSolid.Faces.Size > 0 Then
                            For Each queF As Face In oSolid.Faces
                                resultado.Add(queF)
                            Next
                        End If
                    End If
                Next
            End If
        Next
        Return resultado
    End Function

    Public Function FamilyInstance_DameFacesInstanceGeometry(ByVal doc As Document, fi As FamilyInstance) As List(Of Face)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        options.DetailLevel = ViewDetailLevel.Fine
        '
        Dim resultado As New List(Of Face)

        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetInstanceGeometry
                    If TypeOf go1 Is Solid Then
                        Dim oSolid As Solid = CType(go1, Solid)
                        If oSolid.Faces.Size > 0 Then
                            For Each queF As Face In oSolid.Faces
                                resultado.Add(queF)
                            Next
                        End If
                    End If
                Next
            End If
        Next
        Return resultado
    End Function
    Public Function FamilyInstance_DameGeometryInstance(ByVal doc As Document, fi As FamilyInstance) As GeometryInstance
        Dim resultado As GeometryInstance = Nothing
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                resultado = CType(go, GeometryInstance)
                Exit For
            End If
        Next
        '
        Return resultado
    End Function
    Public Function FamilyInstance_DameSymbolGeometryObjects(ByVal doc As Document, fi As FamilyInstance) As List(Of GeometryObject)
        Dim resultado As New List(Of GeometryObject)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetSymbolGeometry
                    resultado.Add(go)
                Next
            End If
        Next
        '
        Return resultado
    End Function
    ' StableRef
    Public Function FamilyInstance_DameSymbolGeometryObjectsStableRef(ByVal doc As Document, fi As FamilyInstance) As List(Of String)
        Dim resultado As New List(Of String)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetSymbolGeometry
                    If TypeOf go1 Is Curve Then
                        resultado.Add(CType(go1, Curve).Reference.ConvertToStableRepresentation(doc))
                    ElseIf TypeOf go1 Is Solid AndAlso CType(go1, Solid).Faces.Size > 0 Then
                        For Each oFace As Face In CType(go1, Solid).Faces
                            resultado.Add(oFace.Reference.ConvertToStableRepresentation(evRevit.evAppUI.ActiveUIDocument.Document))
                        Next
                    End If
                Next
            End If
        Next
        '
        Return resultado
    End Function
    '
    Public Function FamilyInstance_DameInstanceGeometryObjects(ByVal doc As Document, fi As FamilyInstance) As List(Of GeometryObject)
        Dim resultado As New List(Of GeometryObject)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        ' options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetInstanceGeometry
                    resultado.Add(go)
                Next
            End If
        Next
        '
        Return resultado
    End Function
    ' StableRef
    Public Function FamilyInstance_DameInstanceGeometryObjectsStableRef(ByVal doc As Document, fi As FamilyInstance) As List(Of String)
        Dim resultado As New List(Of String)
        'Dim fi = TryCast(doc.GetElement(elementId), FamilyInstance)
        Dim transforms = fi.GetTransform()
        Dim data As String = String.Empty
        Dim options As Options = New Options()
        options.IncludeNonVisibleObjects = True
        options.ComputeReferences = True
        'options.View = doc.ActiveView
        'options.DetailLevel = ViewDetailLevel.Fine
        For Each go As GeometryObject In fi.Geometry(options)
            If TypeOf go Is GeometryInstance Then
                For Each go1 As GeometryObject In CType(go, GeometryInstance).GetInstanceGeometry
                    If TypeOf go1 Is Curve Then
                        resultado.Add(CType(go1, Curve).Reference.ConvertToStableRepresentation(doc))
                    ElseIf TypeOf go1 Is Solid AndAlso CType(go1, Solid).Faces.Size > 0 Then
                        For Each oFace As Face In CType(go1, Solid).Faces
                            resultado.Add(oFace.Reference.ConvertToStableRepresentation(evRevit.evAppUI.ActiveUIDocument.Document))
                        Next
                    End If
                Next
            End If
        Next
        '
        Return resultado
    End Function
    '
    Public Function Face_DameLine(oFace As Face) As Line
        Dim bbox As BoundingBoxUV = oFace.GetBoundingBox()
        Dim lowerLeft As UV = bbox.Min
        Dim upperRight As UV = bbox.Max
        Dim deltaU As Double = upperRight.U - lowerLeft.U
        Dim deltaV As Double = upperRight.V - lowerLeft.V
        Dim vOffset As Double = deltaV * 0.6
        ' 60% up the face
        Dim firstPoint As UV = lowerLeft + New UV(deltaU * 0.2, vOffset)
        Dim lastPoint As UV = lowerLeft + New UV(deltaU * 0.8, vOffset)
        Dim line__1 As Line = Line.CreateBound(oFace.Evaluate(firstPoint), oFace.Evaluate(lastPoint))
        Return line__1
    End Function
    Public Function GetFamilyReferenceByIndex(ByVal inst As FamilyInstance, ByVal idx As Integer) As Reference
        Dim indexRef As Reference = Nothing
        If inst IsNot Nothing Then
            Dim dbDoc As Document = inst.Document
            Dim sampleStableRef As String = Nothing
            Dim customStableRef As String = Nothing
            Dim oFace As Face = Nothing
            '
            Dim geomOptions As Options = dbDoc.Application.Create.NewGeometryOptions()

            If geomOptions IsNot Nothing Then
                geomOptions.ComputeReferences = True
                geomOptions.DetailLevel = ViewDetailLevel.Undefined
                geomOptions.IncludeNonVisibleObjects = True
            End If
            '
            Dim gElement As GeometryElement = inst.Geometry(geomOptions)
            Dim gInst As GeometryInstance = Nothing
            For Each gObj As GeometryObject In gElement
                If TypeOf gObj Is GeometryInstance Then
                    gInst = TryCast(gObj, GeometryInstance)
                    Exit For
                End If
            Next
            '
            If gInst IsNot Nothing Then
                Try
                    Dim gSymbol As GeometryElement = gInst.GetSymbolGeometry()
                    '
                    If gSymbol IsNot Nothing AndAlso gSymbol.Count() > 0 Then
                        For Each gObj As GeometryObject In gSymbol
                            If TypeOf gObj Is Solid AndAlso CType(gObj, Solid).Faces.Size > 0 Then
                                Dim solid As Solid = TryCast(gObj, Solid)
                                For Each oFa As Face In solid.Faces
                                    oFace = oFa
                                    If oFace IsNot Nothing Then Exit For
                                Next
                                If oFace IsNot Nothing Then Exit For
                            End If
                        Next
                        'Dim solid As Solid = TryCast(gSymbol.First(), Solid)
                        'For Each oFa As Face In solid.Faces
                        '    oFace = oFa
                        '    If oFace IsNot Nothing Then Exit For
                        'Next
                    End If
                    '
                    If oFace Is Nothing Then
                        Return Nothing
                        Exit Function
                    End If
                    sampleStableRef = oFace.Reference.ConvertToStableRepresentation(dbDoc)
                    '
                    If sampleStableRef IsNot Nothing Then
                        Dim tokenList As String() = sampleStableRef.Split(New Char() {":"c})
                        customStableRef = tokenList(0) & ":" + tokenList(1) & ":" + tokenList(2) & ":" + tokenList(3) & ":" & idx.ToString()
                        indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef)
                        Dim geoObj As GeometryObject = inst.GetGeometryObjectFromReference(indexRef)
                        '
                        If geoObj IsNot Nothing Then
                            Dim finalToken As String = ""
                            If TypeOf geoObj Is Edge Then
                                finalToken = ":LINEAR"
                            End If
                            If TypeOf geoObj Is Face Then
                                finalToken = ":SURFACE"
                            End If
                            customStableRef += finalToken
                            indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef)
                        Else
                            indexRef = Nothing
                        End If
                    End If
                Catch ex As Exception
                    indexRef = Nothing
                End Try
            Else
                Throw New Exception("No Symbol Geometry found...")
            End If
        End If
        '
        Return indexRef
    End Function
    Public Enum SpecialReferenceType
        Left = 0
        CenterLR = 1
        Right = 2
        Front = 3
        CenterFB = 4
        Back = 5
        Bottom = 6
        CenterElevation = 7
        Top = 8
    End Enum

    Public Function GetSpecialFamilyReference(ByVal inst As FamilyInstance, ByVal refType As SpecialReferenceType) As Reference
        Dim indexRef As Reference = Nothing
        Dim idx As Integer = CInt(refType)

        If inst IsNot Nothing Then
            Dim dbDoc As Document = inst.Document
            Dim geomOptions As Options = dbDoc.Application.Create.NewGeometryOptions()

            If geomOptions IsNot Nothing Then
                geomOptions.ComputeReferences = True
                geomOptions.DetailLevel = ViewDetailLevel.Undefined
                geomOptions.IncludeNonVisibleObjects = True
            End If

            Dim gElement As GeometryElement = inst.Geometry(geomOptions)
            Dim gInst As GeometryInstance = TryCast(gElement.First(), GeometryInstance)
            Dim sampleStableRef As String = Nothing

            If gInst IsNot Nothing Then
                Dim gSymbol As GeometryElement = gInst.GetSymbolGeometry()

                If gSymbol IsNot Nothing Then

                    For Each geomObj As GeometryObject In gSymbol

                        If TypeOf geomObj Is Solid Then
                            Dim solid As Solid = TryCast(geomObj, Solid)
                            If solid.Faces.Size > 0 Then
                                Dim face As Face = solid.Faces.Item(0)
                                sampleStableRef = face.Reference.ConvertToStableRepresentation(dbDoc)
                                Exit For
                            End If
                        ElseIf TypeOf geomObj Is Curve Then
                            Dim curve As Curve = TryCast(geomObj, Curve)
                            sampleStableRef = curve.Reference.ConvertToStableRepresentation(dbDoc)
                            Exit For
                        ElseIf TypeOf geomObj Is Point Then
                            Dim point As Point = TryCast(geomObj, Point)
                            sampleStableRef = point.Reference.ConvertToStableRepresentation(dbDoc)
                            Exit For
                        End If
                    Next
                End If
                '
                If sampleStableRef IsNot Nothing Then
                    Dim refTokens As String() = sampleStableRef.Split(New Char() {":"c})
                    Dim customStableRef As String = refTokens(0) & ":" + refTokens(1) & ":" + refTokens(2) & ":" + refTokens(3) & ":" & idx.ToString()
                    indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef)
                    Dim geoObj As GeometryObject = inst.GetGeometryObjectFromReference(indexRef)

                    If geoObj IsNot Nothing Then
                        Dim finalToken As String = ""
                        If TypeOf geoObj Is Edge Then
                            finalToken = ":LINEAR"
                        End If
                        If TypeOf geoObj Is Face Then
                            finalToken = ":SURFACE"
                        End If
                        customStableRef += finalToken
                        indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef)
                    Else
                        indexRef = Nothing
                    End If
                End If
            Else
                Throw New Exception("No Symbol Geometry found...")
            End If
        End If
        Return indexRef
    End Function
    ''
    Public Function FASE_DameId(ByRef queDoc As Autodesk.Revit.DB.Document, nFase As String) As ElementId
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        Dim resultado As ElementId = ElementId.InvalidElementId
        Dim filtro As New Autodesk.Revit.DB.FilteredElementCollector(queDoc)
        filtro.OfCategory(BuiltInCategory.OST_Phases)
        ''
        If filtro.Count > 0 Then
            For Each e As Phase In filtro
                If e.Name.ToUpper = nFase.ToUpper Then
                    resultado = e.Id
                    Exit For
                End If
            Next
        End If
        ''
        Return resultado
    End Function
    '
    Public Function FicheroRevit_EsVersion(fullFi As String) As Boolean
        Dim resultado As Boolean = False
        Dim partes() As String = fullFi.Split("."c)
        Dim fullFiBase As String = ""
        '
        For x As Integer = LBound(partes) To UBound(partes)
            ' Si es el bloque que tiene los números, saltar
            If x = UBound(partes) - 1 Then Continue For
            ' Vamos concatenando las partes
            fullFiBase &= partes(x)
        Next
        '
        ' Si existe el fichero sin los número y si la parte partes(UBound(partes) - 1) es numérico. Es una versión
        If IO.File.Exists(fullFiBase) = True AndAlso IsNumeric(partes(UBound(partes) - 1)) Then
            resultado = True
        Else
            resultado = False
        End If
        '
        Return resultado
    End Function
    '
    'Public Sub Fichero_AbreYActivaViejo(queFile As String, Optional activar As Boolean = True)
    '    utMuestraDialogos = False
    '    'AddHandler crUIAppCont.DialogBoxShowing, AddressOf ApplicationUIEvent_DialogBoxShowing_Handler
    '    If evRevit.evAppUI.ActiveUIDocument Is Nothing Then ' utApp1.Documents.Size = 0 Then
    '        Dim oMp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(queFile)
    '        Dim oOop As New OpenOptions
    '        oOop.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
    '        evRevit.evAppUI.ActiveUIDocument = evRevit.evAppUI.OpenAndActivateDocument(oMp, oOop, False)
    '    Else
    '        Dim oMp As ModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(queFile)
    '        Dim oOop As New OpenOptions
    '        oOop.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
    '        If utApp1 Is Nothing Then if utApp1 is Nothing then utApp1 = evRevit.evAppUI.Application
    '        evRevit.evAppUI.ActiveUIDocument.Document = Documento_EstaAbierto(utApp1, queFile, False)
    '        If evRevit.evAppUI.ActiveUIDocument.Document Is Nothing Then
    '            'evRevit.evAppUI.ActiveUIDocument.Document = utApp1.OpenDocumentFile(oMp, oOop)
    '            evRevit.evAppUI.ActiveUIDocument.Document = utApp1.OpenDocumentFile(queFile)
    '        End If
    '        If evRevit.evAppUI.ActiveUIDocument.Document IsNot Nothing Then
    '            Dim v3D As View3D = View3D_Dame(evRevit.evAppUI.ActiveUIDocument.Document)
    '            Dim uiDoc As UIDocument = New UIDocument(evRevit.evAppUI.ActiveUIDocument.Document)
    '            Dim collector As New FilteredElementCollector(evRevit.evAppUI.ActiveUIDocument.Document, v3D.Id) ', uiDoc.ActiveGraphicalView.Id)
    '            'collector.OfClass(GetType(FamilyInstance))
    '            'collector.OfClass(GetType(Element))
    '            Dim lstElement As New List(Of ElementId)
    '            For Each oEle As Element In collector
    '                lstElement.Add(oEle.Id)
    '            Next
    '            '
    '            uiDoc.ShowElements(lstElement)
    '            uiDoc.RefreshActiveView()
    '            'uiDoc.ActiveView = v3D
    '            'uiDoc.RefreshActiveView()
    '            Try
    '                If activar Then View3D_ActivaCierraOtras(evRevit.evAppUI.ActiveUIDocument.Document)
    '            Catch ex As Exception
    '                Debug.Print(ex.ToString)
    '            End Try
    '        End If
    '    End If
    '    'RemoveHandler crUIAppCont.DialogBoxShowing, AddressOf ApplicationUIEvent_DialogBoxShowing_Handler
    '    utMuestraDialogos = True
    'End Sub
    '

    Public Sub DocumentoYaAbierto_PonActivo(queFile As String)
        utMuestraDialogos = False
        ' Si solo hay un documento activo. Salir.
        If evRevit.evAppUI.Application.Documents.Size = 1 Then Exit Sub
        ' Si ya es el activo. Salir
        If evRevit.evAppUI.ActiveUIDocument.Document.PathName = queFile Then Exit Sub
        '
        For Each oD As Document In evRevit.evAppUI.Application.Documents
            Dim uiDoc As UIDocument = New UIDocument(oD)
            Dim v3D As View3D = View3D_Dame(oD)
            Dim collector As New FilteredElementCollector(oD, v3D.Id) ', uiDoc.ActiveGraphicalView.Id)
            'collector.OfClass(GetType(FamilyInstance))
            'collector.OfClass(GetType(Element))
            Dim lstElement As New List(Of ElementId)
            For Each oEle As Element In collector
                lstElement.Add(oEle.Id)
            Next
            '
            uiDoc.ShowElements(lstElement)
            uiDoc.RefreshActiveView()
            '
            uiDoc = Nothing
            v3D = Nothing
            collector = Nothing
            lstElement = Nothing
        Next
        utMuestraDialogos = True
    End Sub
    Public Sub DocumentoYaAbierto_PonActivo(oD As Document)
        If oD Is Nothing Then Exit Sub
        utMuestraDialogos = False
        ' Si solo hay un documento activo. Salir.
        If evRevit.evAppUI.Application.Documents.Size = 1 Then Exit Sub
        ' Si ya es el activo. Salir
        If evRevit.evAppUI.ActiveUIDocument.Document.PathName = oD.PathName Then Exit Sub
        '
        Dim uiDoc As UIDocument = New UIDocument(oD)
        Dim v3D As View3D = View3D_Dame(oD)
        Dim collector As ICollection(Of ElementId) = New FilteredElementCollector(oD, v3D.Id).OfClass(GetType(FamilyInstance)).ToElementIds
        'collector.OfClass(GetType(FamilyInstance))
        'collector.OfClass(GetType(Element))
        'Dim lstElement As New List(Of ElementId)
        'For Each oEle As Element In collector
        '    lstElement.Add(oEle.Id)
        'Next
        '
        'uiDoc.ActiveView = v3D
        uiDoc.ShowElements(collector)
        'ZoomObjecto(uiDoc, lstElement)
        uiDoc.RefreshActiveView()
        '
        uiDoc = Nothing
        v3D = Nothing
        collector = Nothing
        'lstElement = Nothing
        '
        utMuestraDialogos = True
    End Sub
    Public Function CategoriasDocumento_DameNombres(queDoc As Autodesk.Revit.DB.Document) As List(Of String)
        Dim resultado As New List(Of String)
        '
        Dim documentSettings As Settings = queDoc.Settings
        ' Todas las categorias del documento actual
        Dim groups As Categories = documentSettings.Categories
        ' Coger el item "Floor category" desde BuiltInCategory.OST_Floors
        'Dim floorCategory As Category = groups.Item(BuiltInCategory.OST_Floors)
        'floorCategory.Name
        For Each queCat As Category In groups
            resultado.Add(queCat.Name)
        Next
        '
        resultado.Sort()
        Return resultado
    End Function
    ' Ejectuar este procedimiento con System.Threading.ThreadPool.QueueUserWorkItem(New System.Threading.WaitCallback(AddressOf Documento_CierraActual))
    Public Sub Documento_CierraActual(stateinfo As Object)
        Try
            System.Windows.Forms.SendKeys.SendWait("^{F4}")
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
    End Sub

    '
    Public Function PhaseFilter_DameTodos(quedoc As Autodesk.Revit.DB.Document) As List(Of PhaseFilter)
        Dim collector As New FilteredElementCollector(quedoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.PhaseFilter))
        Dim resultado As New List(Of PhaseFilter)
        For Each oEle As Element In collector
            resultado.Add(CType(oEle, PhaseFilter))
        Next
        Return resultado
    End Function
    '
    Public Function PhaseFilter_DameUno(quedoc As Autodesk.Revit.DB.Document, quePf As String) As PhaseFilter
        Dim collector As New FilteredElementCollector(quedoc)
        collector = collector.OfClass(GetType(Autodesk.Revit.DB.PhaseFilter))
        Dim resultado As PhaseFilter = Nothing
        For Each oEle As Element In collector
            If oEle.Name.ToUpper.Contains(quePf) Then
                Debug.Print(oEle.Name)
                resultado = CType(oEle, PhaseFilter)
                Exit For
            End If
        Next
        Return resultado
    End Function
    '
    Private Sub Fichero_QuitarSoloLectura(queFi As String)
        Try
            Dim FileInfo As New FileInfo(queFi)
            FileInfo.Attributes = IO.FileAttributes.Normal
            FileInfo.Refresh()
        Catch ex As Exception

        End Try
    End Sub
    Public Function Fichero_EsSoloLectura(queFi As String, Optional quitarSoloLectura As Boolean = False) As Boolean
        Dim resultado As Boolean = False
        Dim FileInfo As New FileInfo(queFi)
        If (FileInfo.Attributes And FileAttributes.ReadOnly) = FileAttribute.ReadOnly Then
            resultado = True
        End If
        '
        If resultado = True AndAlso quitarSoloLectura = True Then
            Fichero_QuitarSoloLectura(queFi)
        End If
        Return resultado
    End Function
    '
    Public Sub ExecuteCommandId(queComando As Object)
        'ID_REVIT_FILE_OPEN (OpenRevitFile)
        'ID_WINDOW_CLOSE_HIDDEN
        Dim cId As RevitCommandId = Nothing
        If TypeOf queComando Is PostableCommand Then
            cId = RevitCommandId.LookupPostableCommandId(CType(queComando, PostableCommand))
        ElseIf TypeOf queComando Is String Then
            cId = RevitCommandId.LookupCommandId(CType(queComando, String))
        End If
        If cId IsNot Nothing Then
            Try
                evRevit.evAppUI.PostCommand(cId)
            Catch ex As Exception
                MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
            End Try
        End If
    End Sub

    Public Sub ExecuteCommandIdRevit(queComando As PostableCommand)
        'ID_REVIT_FILE_OPEN (OpenRevitFile)
        Dim cId As RevitCommandId = RevitCommandId.LookupPostableCommandId(queComando)
        If cId IsNot Nothing Then
            Try
                evRevit.evAppUI.PostCommand(cId)
            Catch ex As Exception
                MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
            End Try
        End If
    End Sub

    Public Sub ExecuteCommandIdAddIn(queComando As String)
        'ID_REVIT_FILE_OPEN (OpenRevitFile)
        'ID_WINDOW_CLOSE_HIDDEN
        Dim cId As RevitCommandId = RevitCommandId.LookupCommandId(queComando)
        If cId IsNot Nothing Then
            Try
                evRevit.evAppUI.PostCommand(cId)
            Catch ex As Exception
                'MsgBox("Error --> Command " & queComando.ToString & "Is Not exist")
            End Try
        End If
    End Sub
End Module

#Region "hWnd Wrapper Class"
' This class is used to wrap a Win32 hWnd as a .Net IWind32Window =class.
' This is primarily used for parenting a dialog to the Inventor =window.
'
' For example:
' myForm.Show(New =WindowWrapper(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle))
'
' Private Sub m_featureCountButtonDef_OnExecute( =... )
'' Display the dialog.
'Dim myForm As New InsertBoltForm
'myForm.Show(New =WindowWrapper(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle))
'Sub

Public Class WindowWrapper
    Implements System.Windows.Forms.IWin32Window

    Private _hwnd As IntPtr

    Public Sub New(ByVal handle As IntPtr)
        _hwnd = handle
    End Sub

    Public ReadOnly Property Handle() As IntPtr _
      Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

End Class

#End Region

#Region "Filtros"

'' Filtro para varias clases de objetos Revit
Public Class PickFilterIds
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        If arrIds.Contains(elem.Category.Id.IntegerValue) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class

'' Filtro para varias clases de objetos Revit
Public Class PickFilterGeneral
    Implements ISelectionFilter
    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        If elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Rooms)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Doors)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Windows)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Walls)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_StructuralColumns)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Roofs)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_GenericModel)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Casework)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_SpecialityEquipment)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Stairs)) OrElse
            elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Furniture)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class

''' Filter to constrain picking to rooms
Public Class PickFilterRoom
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        Return (elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Rooms)))
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class

Public Class PickFilterDoor
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        Return (elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Doors)))
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class

Public Class PickFilteWindowr
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        Return (elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Windows)))
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class

Public Class PickFilterWall
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        Return (elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_Walls)))
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class
#End Region

#Region "PICKFILTERULMA"
''
Public Class PickFilterULMAFamilyInstance
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        ''
        If TypeOf elem Is FamilyInstance = False Then
            Return False
            Exit Function
        End If
        ''
        Dim manu As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, elem, BuiltInParameter.ALL_MODEL_MANUFACTURER)
        ''
        If manu.ToUpper.Contains("ULMA") Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class
''
Public Class PickFilterULMASetFamilyInstanceCategory
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        If TypeOf elem Is FamilyInstance = False Then
            Return False
            Exit Function
        End If
        'Return (elem.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_GenericModel)))
        Dim manu As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, elem, BuiltInParameter.ALL_MODEL_MANUFACTURER)
        If manu.ToUpper.Contains("ULMA") AndAlso elem.Category.Id.IntegerValue.Equals(CInt(typeFamily)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class
''
Public Class PickFilterULMAFamilyInstanceCategory
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        If TypeOf elem Is FamilyInstance = False Then
            Return False
            Exit Function
        End If
        ''
        Dim manu As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, elem, BuiltInParameter.ALL_MODEL_MANUFACTURER)
        ''
        If manu.ToUpper.Contains("ULMA") AndAlso elem.Category.Id.IntegerValue.Equals(CInt(typeFamily)) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
End Class
''
Public Class PickFilterULMASetAndFamilyInstanceCategory
    Implements ISelectionFilter

    Public Function AllowElement(elem As Element) As Boolean Implements ISelectionFilter.AllowElement
        If TypeOf elem Is FamilyInstance = False Then
            Return False
            Exit Function
        End If
        ''
        Dim manu As String = ParametroElementLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, elem, BuiltInParameter.ALL_MODEL_MANUFACTURER)
        ''
        If manu.ToUpper.Contains("ULMA") And
            (elem.Category.Id.IntegerValue.Equals(CInt(typeFamily)) Or
             elem.Category.Id.IntegerValue.Equals(CInt(typeSet))) Then
            Return True
        Else
            Return False
        End If
        ''
    End Function

    Public Function AllowReference(reference As Reference, position As XYZ) As Boolean Implements ISelectionFilter.AllowReference
        Return False
    End Function
    '
    Public Shared Function XYZ_TransformPoint(point As XYZ, transform As Transform) As XYZ
        Dim x As Double = point.X
        Dim y As Double = point.Y
        Dim z As Double = point.Z

        'transform basis of the old coordinate system in the new coordinate // system
        Dim b0 As XYZ = transform.Basis(0)
        Dim b1 As XYZ = transform.Basis(1)
        Dim b2 As XYZ = transform.Basis(2)
        Dim origin As XYZ = transform.Origin

        'transform the origin of the old coordinate system in the new 
        'coordinate system
        Dim xTemp As Double = x * b0.X + y * b1.X + z * b2.X + origin.X
        Dim yTemp As Double = x * b0.Y + y * b1.Y + z * b2.Y + origin.Y
        Dim zTemp As Double = x * b0.Z + y * b1.Z + z * b2.Z + origin.Z
        Return New XYZ(xTemp, yTemp, zTemp)
    End Function
    '
End Class
#End Region
''
#Region "AVISOS REVIT"
''
'Public Class FamilyOptions
'    Implements IFamilyLoadOptions

'    Public Function OnFamilyFound(familyInUse As Boolean, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnFamilyFound
'        overwriteParameterValues = True
'        Return True
'    End Function

'    Public Function OnSharedFamilyFound(sharedFamily As Family, familyInUse As Boolean, ByRef source As FamilySource, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnSharedFamilyFound
'        Return True
'    End Function
'    ''
'    Public Sub InsertaFamilia(ultFicheroInsertar As String)
'        ' ''
'        Dim familia As Family = Nothing
'        Using tx As Transaction = New Transaction(oDoc)
'            tx.Start("Cargar Familia")
'            '    Dim ultFicheroInsertar As String = CType(e.Item, System.Windows.Forms.ListViewItem).Tag.ToString
'            If oDoc.LoadFamily(ultFicheroInsertar, familia) = False Then
'                '        MsgBox("Error cargando la familia " & ultFicheroInsertar)
'                tx.RollBack()
'            Else
'                tx.Commit()
'            End If
'        End Using
'        ' ''
'        If familia Is Nothing Then Exit Sub
'        Dim ultFamilia As Family = CType(oDoc.GetElement(familia.Id), Family)
'        ' ''
'        ' '' Determine the family symbol
'        Dim symbol As FamilySymbol = Nothing
'        For Each oId As ElementId In familia.GetFamilySymbolIds
'            symbol = CType(oDoc.GetElement(oId), FamilySymbol)
'            Exit For
'        Next
'        ' ''
'        ' '' Si no tiene ningún elemento, salimos
'        If symbol Is Nothing Then Exit Sub
'        Dim ultSimbolo As FamilySymbol = CType(oDoc.GetElement(symbol.Id), FamilySymbol)
'        ' ''

'        ' ''
'        If ultSimbolo Is Nothing Then
'            Exit Sub
'        Else
'            '    'uidoc.PromptForFamilyInstancePlacement(ultSimbolo)
'            Dim filtro As Autodesk.Revit.DB.FilteredElementCollector = New Autodesk.Revit.DB.FilteredElementCollector(oDoc)
'            filtro.OfCategoryId(ultSimbolo.Category.Id)
'            filtro.OfClass(GetType(ElementType))
'            filtro.FirstElement()
'            Dim oEt As ElementType = CType(filtro.FirstElement(), ElementType)
'            uidoc.PostRequestForElementTypePlacement(oEt)
'        End If
'        ''
'    End Sub
'End Class
''
'This class supresses duplication type warning dialog by forcing copy paste to use destination project types
Public Class DuplicateTypeNamesHandler
    Implements IDuplicateTypeNamesHandler

    Public Function OnDuplicateTypeNamesFound(args As DuplicateTypeNamesHandlerArgs) As DuplicateTypeAction Implements IDuplicateTypeNamesHandler.OnDuplicateTypeNamesFound
        'Always use duplicate destination types when asked
        Return DuplicateTypeAction.UseDestinationTypes
    End Function
    '' ***** Para utilizar esta clase.
    ''// create an instance
    'Dim copyOptions as new CopyPasteOptions;
    'copyOptions.SetDuplicateTypeNamesHandler(new DuplicateTypeNamesHandler)
    '// now the copy
    'ElementTransformUtils.CopyElements(doc1, ids, doc2, Transform.Identity, copyOptions)
End Class

'This class silently handles failures caused by finding duplicates and allows revit to continue	
Public Class FailuresPreprocessor
    Implements IFailuresPreprocessor

    '' Despues de ElementTransformUtils.CopyElements(doc1, ids, doc2, Transform.Identity, copyOptions) Poner esto.
    'setup a failure handler to supress warnings when the transaction is commited
    'Dim failureOptions As FailureHandlingOptions = Transaction.GetFailureHandlingOptions
    'failureOptions.SetFailuresPreprocessor(New FailuresPreprocessor)

    Public Function PreprocessFailures(ByVal failuresAccessor As FailuresAccessor) As FailureProcessingResult Implements IFailuresPreprocessor.PreprocessFailures
        'look through all the failure messages
        failuresAccessor.DeleteAllWarnings()
        For Each failureMessage As FailureMessageAccessor In failuresAccessor.GetFailureMessages
            'Delete any "Can't paste duplicate types.  Only non duplicate types will be pasted." messages
            'If failureMessage.GetFailureDefinitionId = BuiltInFailures.CopyPasteFailures.CannotCopyDuplicates Then
            failuresAccessor.DeleteWarning(failureMessage)
            'End If
        Next
        'all othe messages will have to be dealt with interactively by the user
        Return FailureProcessingResult.Continue
    End Function
End Class
''
Public Class LoadedFamilyDropHandler
    Implements Autodesk.Revit.UI.IDropHandler
    ''
    Public Sub Execute(document As UIDocument, data As Object) Implements IDropHandler.Execute
        ' ''
        Dim familia As Family = Nothing
        ' '' Determine the family symbol
        Dim symbol As FamilySymbol = Nothing
        Dim fichero As String = data.ToString
        Dim nombre As String = IO.Path.GetFileNameWithoutExtension(fichero)
        Using tx As Transaction = New Transaction(document.Document)
            tx.Start("Load Family")
            '    Dim ultFicheroInsertar As String = CType(e.Item, System.Windows.Forms.ListViewItem).Tag.ToString
            If document.Document.LoadFamilySymbol(fichero,
                                                  nombre,
                                                  New FamilyOptions,
                                                  symbol) = False Then
                tx.RollBack()
            Else
                tx.Commit()
            End If
        End Using
        '' Si no tiene ningún elemento, salimos
        If symbol Is Nothing Then
            Exit Sub
        End If
        ''
        'Using tx As Transaction = New Transaction(document.Document)
        'tx.Start("Insert FamilySymbol")
        Dim ultSimbolo As FamilySymbol = CType(document.Document.GetElement(symbol.Id), FamilySymbol)
        ''
        If ultSimbolo Is Nothing Then
            'tx.RollBack()
            Exit Sub
        Else
            Dim filtro As Autodesk.Revit.DB.FilteredElementCollector = New Autodesk.Revit.DB.FilteredElementCollector(document.Document)
            filtro.OfCategoryId(ultSimbolo.Category.Id)
            filtro.OfClass(GetType(ElementType))
            '' Buscamos el nombre
            '' ***** LINQ para crear IEnumerable de los FamilySymbol con el nombre = nombreFamilia
            Dim query As System.Collections.Generic.IEnumerable(Of Autodesk.Revit.DB.Element)
            query = From element In filtro
                    Where (TypeOf element Is FamilySymbol) AndAlso
                    (CType(element, Autodesk.Revit.DB.FamilySymbol).FamilyName = nombre OrElse
                    CType(element, Autodesk.Revit.DB.FamilySymbol).Family.Name = nombre)
                    Select element
            '' ***** Para coger el ID del FamilySymbol
            Dim famSyms As List(Of Element) = query.ToList
            Dim symbolId As ElementType = Nothing
            If famSyms.Count <> 0 Then
                symbolId = CType(famSyms(0), ElementType)
            End If
            ''
            If symbolId IsNot Nothing Then
                Dim oEt As ElementType = symbolId   ' CType(filtro.LastOrDefault, ElementType)
                compruebacambios = True
                document.PostRequestForElementTypePlacement(oEt)   '' Esta no espera y continua
                compruebacambios = False
                'document.PromptForFamilyInstancePlacement(CType(oEt, FamilySymbol)) '' Espera a que insertemos.
                ''
                'tx.Commit()
                'ULMALGFree.clsLogsCSV._ultimaApp =  = ULMALGFree.queApp.UCREVIT
                If ultSimbolo.Document.PathName <> "" AndAlso IO.File.Exists(ultSimbolo.Document.PathName) Then
                    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("DRAG&DROP", FILENAME:=ultSimbolo.Document.PathName, FAMILY:=ultSimbolo.FamilyName)
                Else
                    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("DRAG&DROP", FAMILY:=ultSimbolo.FamilyName)
                End If
            End If
        End If
        'End Using
    End Sub
End Class
''
Public Class AvisosAlPreProcesar
    Implements IFailuresPreprocessor

    Public Function PreprocessFailures(
      ByVal failuresAccessor As FailuresAccessor) _
      As FailureProcessingResult _
      Implements IFailuresPreprocessor.PreprocessFailures

        'Dim deleteWarning As Boolean = True

        Try
            ''
            '' Quitamos todos los mensajes de aviso
            failuresAccessor.DeleteAllWarnings()
            ''
            '' Recorrer todos los tipos de avisos y desactivarlos en función
            '' de nuestra elección.
            'Dim flist As IList(Of FailureMessageAccessor) = failuresAccessor.GetFailureMessages

            'For Each f As FailureMessageAccessor In flist

            'Dim fDefId As FailureDefinitionId = f.GetFailureDefinitionId

            'Select Case fDefId
            '    Case BuiltInFailures.GroupFailures.AtomViolationWhenOnePlaceInstance
            '        deleteWarning = True
            'End Select

            'If deleteWarning = True Then
            '    failuresAccessor.DeleteWarning(f)
            'End If
            ''
            '' Esta línea para desactivarlos 1 a 1
            'failuresAccessor.DeleteWarning(f)
            'Next

        Catch ex As Exception

        End Try
        ''
        Return FailureProcessingResult.Continue
    End Function
End Class
''
'This class silently handles failures caused by finding duplicates and allows revit to continue	
Public Class AvisosAlPreProcesarGrupos
    Implements IFailuresPreprocessor


    Public Function PreprocessFailures(
      ByVal failuresAccessor As FailuresAccessor) _
      As FailureProcessingResult _
      Implements IFailuresPreprocessor.PreprocessFailures

        'Dim deleteWarning As Boolean = True
        ''
        Dim resultado As FailureProcessingResult = FailureProcessingResult.Continue
        'Dim estaresuelto As Boolean = False
        Dim nResueltos As Integer = 0
        Dim flist As IList(Of FailureMessageAccessor) = failuresAccessor.GetFailureMessages
        ''
        Try
            ''
            '' Quitamos todos los mensajes de aviso
            'failuresAccessor.DeleteAllWarnings()
            ''
            '' Recorrer todos los tipos de avisos y desactivarlos en función
            '' de nuestra elección.

            ''
            '' Si no hay mensajes, salimos con continuar.
            If flist.Count = 0 Then
                Return FailureProcessingResult.Continue
                Exit Function
            End If
            ''

            For Each f As FailureMessageAccessor In flist
                Dim fDefId As FailureDefinitionId = f.GetFailureDefinitionId
                ''
                If f.HasResolutions Then
                    failuresAccessor.ResolveFailure(f)
                    'estaresuelto = True
                    nResueltos += 1
                    'Else
                    'estaresuelto = False
                End If
            Next
        Catch ex As Exception

        End Try
        ''
        'If estaresuelto = True Then
        '    resultado = FailureProcessingResult.ProceedWithCommit
        'ElseIf estaresuelto = False Then
        '    resultado = FailureProcessingResult.Continue
        'End If
        If nResueltos = flist.Count Then
            resultado = FailureProcessingResult.ProceedWithCommit
        Else
            resultado = FailureProcessingResult.Continue
        End If
        ''
        Return resultado
        ''My.Computer.Keyboard.SendKeys("{ENTER}", True)
    End Function
End Class
Public Class FamilyOptions
    Implements IFamilyLoadOptions
    ''
    Public Function OnFamilyFound(familyInUse As Boolean, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnFamilyFound
        overwriteParameterValues = True
        Return True
    End Function

    Public Function OnSharedFamilyFound(sharedFamily As Family, familyInUse As Boolean, ByRef source As FamilySource, ByRef overwriteParameterValues As Boolean) As Boolean Implements IFamilyLoadOptions.OnSharedFamilyFound
        source = FamilySource.Family
        overwriteParameterValues = True
        Return True
    End Function
End Class
''
Public Class LoadedFamilyIDDropHandler
    Implements Autodesk.Revit.UI.IDropHandler


    Public Sub Execute(document As UIDocument, data As Object) Implements IDropHandler.Execute
        Dim familySymbolId As ElementId = CType(data, ElementId)
        ''
        Dim symbol As FamilySymbol = CType(document.Document.GetElement(familySymbolId), FamilySymbol)
        ''
        If symbol IsNot Nothing Then
            document.PromptForFamilyInstancePlacement(symbol)
        End If
    End Sub
End Class

'''Using t as new Transaction(doc, "Join All Walls/Column")
'	t.Start()
'	Dim options as FailureHandlingOptions = t.GetFailureHandlingOptions()
'	Dim preprocessor as new WarningOffLinks
'	options.SetFailuresPreprocessor(preproccessor)
'	options.SetClearAfterRollback(true)
'	t.SetFailureHandlingOptions(options)
'	
'	try
'		Dim check as boolean = JoinGeometryUtils.AreElementsJoined(doc, w, c)
'		if (check = true) Then
'			JoinGeometryUtils.SwitchJoinOrder(doc, w, c)
'		else
'			'' No hacemos nada
'		end if
'	catch
'		'' No hacemos nada
'	End Try
'	''
'	t.Commit()
'End Using
''
'' Borrar avisos de borrado de link
Public Class AvisosOffLinks
    Implements IFailuresPreprocessor

    Public Function PreprocessFailures(failuresAccessor As FailuresAccessor) As FailureProcessingResult _
        Implements IFailuresPreprocessor.PreprocessFailures
        ''
        Dim failures As IList(Of FailureMessageAccessor) = failuresAccessor.GetFailureMessages
        For Each f As FailureMessageAccessor In failures
            Dim id As FailureDefinitionId = f.GetFailureDefinitionId()
            If id = BuiltInFailures.LinkFailures.DeleteLinkSymbolPrompt Then
                failuresAccessor.DeleteWarning(f)
            End If
        Next

        Return FailureProcessingResult.Continue
    End Function
End Class
''
'' Borrar todos los avisos.
Public Class AvisosOffAll
    Implements IFailuresPreprocessor


    Public Function PreprocessFailures(failuresAccessor As FailuresAccessor) As FailureProcessingResult _
        Implements IFailuresPreprocessor.PreprocessFailures
        ''
        Dim failures As IList(Of FailureMessageAccessor) = failuresAccessor.GetFailureMessages
        For Each f As FailureMessageAccessor In failures
            'Dim id As FailureDefinitionId = f.GetFailureDefinitionId()
            'If id = BuiltInFailures.LinkFailures.DeleteLinkSymbolPrompt Then
            failuresAccessor.DeleteWarning(f)
            'End If
        Next

        Return FailureProcessingResult.Continue
    End Function
End Class
#End Region
