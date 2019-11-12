Imports System.Text
Imports System.Xml
Imports System.Linq
Imports System.Reflection
Imports System.ComponentModel
Imports System.Collections
Imports System.Collections.Generic
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports System.Windows.Forms

Imports System.IO

Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Events
Imports Autodesk.Revit.DB.Architecture
Imports Autodesk.Revit.DB.Structure

Imports Autodesk.Revit.DB.Mechanical
Imports Autodesk.Revit.DB.Electrical

Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.UI

Imports Autodesk.Revit.UI.Selection

Imports Autodesk.Revit.UI.Events
'Imports Autodesk.Revit.Collections
Imports Autodesk.Revit.Exceptions
'Imports Autodesk.Revit.Utility
Imports Autodesk.Revit.ApplicationServices.Application
Imports Autodesk.Revit.DB.Document
Imports adWin = Autodesk.Windows
'
Partial Public Class evRevit
    ' OBJETOS QUE GENERAN EVENTOS
    Public Shared WithEvents evAppUIC As Autodesk.Revit.UI.UIControlledApplication  ' Igual que evAppUI
    Public Shared evAppUI As Autodesk.Revit.UI.UIApplication                        ' Sin eventos. Ya estan en evAppUIC
    Public Shared WithEvents evAppC As Autodesk.Revit.ApplicationServices.ControlledApplication     ' Igual que evApp
    Public Shared evApp As Autodesk.Revit.ApplicationServices.Application                           ' Sin eventos. Ya estan en evAppC
    Public Shared WithEvents evCompM As Autodesk.Windows.ComponentManager
    'Public Shared WithEvents evDB As Autodesk.Revit.DB.
    '
    'Public Shared _registered As Boolean
    Public Shared comprobarAD As Boolean
    ' CONTROL EVENTOS
    'Public Shared _SubscritoUIControllerApplication As Boolean = True
    'Public Shared _SubscritoControllerApplication As Boolean = True
    Public Shared _SubscritoComponentManager As Boolean = False
    'Public Shared _SubscritoUIApplication As Boolean = False
    'Public Shared _SubscritoApplication As Boolean = False
    Public Shared evOpen As Boolean = True
    Public Shared evClose As Boolean = True
    Public Shared evSave As Boolean = True

    Public Shared Function evDocUI() As Autodesk.Revit.UI.UIDocument
        If evAppUI Is Nothing Then
            TaskDialog.Show("NOTICE TO USER", "evDocUI-->evAppUI is Nothing")
            Return Nothing
        Else
            Return evRevit.evAppUI.ActiveUIDocument
        End If
    End Function
    Public Shared Function evDoc() As Autodesk.Revit.DB.Document
        If evAppUI Is Nothing OrElse evAppUI.ActiveUIDocument Is Nothing Then
            TaskDialog.Show("NOTICE TO USER", "evDoc-->evAppUI or ActiveUIDocument is Nothing")
            Return Nothing
        Else
            Return evRevit.evAppUI.ActiveUIDocument.Document
        End If
    End Function
    Public Shared Function evView() As Autodesk.Revit.DB.View
        If evAppUI Is Nothing OrElse evAppUI.ActiveUIDocument Is Nothing Then
            TaskDialog.Show("NOTICE TO USER", "evView-->evAppUI or ActiveUIDocument is Nothing")
            Return Nothing
        Else
            Return evRevit.evAppUI.ActiveUIDocument.ActiveView
        End If
    End Function
    Public Shared Function evSel() As Autodesk.Revit.UI.Selection.Selection
        If evAppUI Is Nothing OrElse evAppUI.ActiveUIDocument Is Nothing Then
            TaskDialog.Show("NOTICE TO USER", "evSel-->evAppUI or ActiveUIDocument is Nothing")
            Return Nothing
        Else
            Return evRevit.evAppUI.ActiveUIDocument.Selection
        End If
    End Function
#Region "SUBSCRIBE"
    Public Shared Sub SubscribeAll()
        'SubscribeForControllerApplication()
        SubscribeForComponentManager()
        'SubscribeForUIApplication()
        'SubscribeForApplication()
    End Sub
    Public Shared Sub SubscribeForControllerApplication()
        'If evAppC Is Nothing Then Exit Sub
        'If evRevit._SubscritoControllerApplication = True Then Exit Sub
        ''
        'Try
        '    AddHandler evAppC.ApplicationInitialized, AddressOf evAppC_ApplicationInitialized
        '    AddHandler evAppC.DocumentOpened, AddressOf evAppC_DocumentOpened
        '    AddHandler evAppC.DocumentOpening, AddressOf evAppC_DocumentOpening
        '    evRevit._SubscritoControllerApplication = True
        'Catch ex As Exception
        '    evRevit._SubscritoControllerApplication = False
        'End Try
    End Sub
    Public Shared Sub SubscribeForComponentManager()
        If evRevit._SubscritoComponentManager = True Then Exit Sub
        Try
            AddHandler adWin.ComponentManager.UIElementActivated, AddressOf evCompM_UIElementActivated
            AddHandler adWin.ComponentManager.ItemAddedToQuickAccessToolBar, AddressOf evCompM_ItemAddedToQuickAccessToolBar
            AddHandler adWin.ComponentManager.ItemExecuted, AddressOf evCompM_ItemExecuted
            AddHandler adWin.ComponentManager.ItemInitialized, AddressOf evCompM_ItemInitialized
            AddHandler adWin.ComponentManager.PreviewExecute, AddressOf evCompM_PreviewExecute
            AddHandler adWin.ComponentManager.PropertyChanged, AddressOf evCompM_PropertyChanged
            AddHandler adWin.ComponentManager.ToolTipClosed, AddressOf evCompM_ToolTipClosed
            AddHandler adWin.ComponentManager.ToolTipOpened, AddressOf evCompM_ToolTipOpened
            evRevit._SubscritoComponentManager = True
        Catch ex As Exception
            evRevit._SubscritoComponentManager = False
        End Try
    End Sub
    'Public Shared Sub SubscribeForUIApplication()
    '    If evAppUI Is Nothing Then Exit Sub
    '    If evRevit._SubscritoUIApplication = True Then Exit Sub
    '    Try
    '        AddHandler evAppUI.ViewActivated, AddressOf evAppUI_ViewActivated
    '        AddHandler evAppUI.ApplicationClosing, AddressOf evAppUI_ApplicationClosing
    '        AddHandler evAppUI.DialogBoxShowing, AddressOf evAppUI_DialogBoxShowing
    '        AddHandler evAppUI.DisplayingOptionsDialog, AddressOf evAppUI_DisplayingOptionsDialog
    '        AddHandler evAppUI.DockableFrameFocusChanged, AddressOf evAppUI_DockableFrameFocusChanged
    '        AddHandler evAppUI.DockableFrameVisibilityChanged, AddressOf evAppUI_DockableFrameVisibilityChanged
    '        AddHandler evAppUI.FabricationPartBrowserChanged, AddressOf evAppUI_FabricationPartBrowserChanged
    '        ' No subscribir eventos Idling. Da error al llamarlo desde Idling.
    '        'AddHandler evAppUI.Idling, AddressOf evAppUI_Idling
    '        'If newdevelopments >= 1 Then
    '        '    AddHandler evAppUI.Idling, AddressOf evAppUI_Idling_Handler_NOAutorizado
    '        'End If
    '        AddHandler evAppUI.TransferredProjectStandards, AddressOf evAppUI_TransferredProjectStandards
    '        AddHandler evAppUI.TransferringProjectStandards, AddressOf evAppUI_TransferringProjectStandards
    '        AddHandler evAppUI.ViewActivating, AddressOf evAppUI_ViewActivating
    '        evRevit._SubscritoUIApplication = True
    '    Catch ex As Exception
    '        evRevit._SubscritoUIApplication = False
    '    End Try
    'End Sub
    'Public Shared Sub SubscribeForApplication()
    '    If evApp Is Nothing Then Exit Sub
    '    If evRevit._SubscritoApplication = True Then Exit Sub
    '    Try
    '        'AddHandler evApp.ApplicationInitialized, AddressOf evApp_ApplicationInitialized
    '        AddHandler evApp.DocumentChanged, AddressOf evApp_DocumentChanged
    '        AddHandler evApp.DocumentClosing, AddressOf evApp_DocumentClosing
    '        AddHandler evApp.DocumentClosed, AddressOf evApp_DocumentClosed
    '        AddHandler evApp.DocumentCreated, AddressOf evApp_DocumentCreated
    '        AddHandler evApp.DocumentCreating, AddressOf evApp_DocumentCreating
    '        '
    '        AddHandler evApp.DocumentOpened, AddressOf evAppCont_DocumentOpened
    '        AddHandler evApp.DocumentOpening, AddressOf evAppCont_DocumentOpening
    '        '
    '        AddHandler evApp.DocumentSaving, AddressOf evApp_DocumentSaving
    '        AddHandler evApp.DocumentSaved, AddressOf evApp_DocumentSaved
    '        AddHandler evApp.DocumentSavingAs, AddressOf evApp_DocumentSavingAs
    '        AddHandler evApp.DocumentSavedAs, AddressOf evApp_DocumentSavedAs
    '        AddHandler evApp.FailuresProcessing, AddressOf evApp_FailuresProcessing
    '        ' Eventos sin funciones.
    '        AddHandler evApp.DocumentPrinted, AddressOf evApp_DocumentPrinted
    '        AddHandler evApp.DocumentPrinting, AddressOf evApp_DocumentPrinting
    '        AddHandler evApp.DocumentSynchronizedWithCentral, AddressOf evApp_DocumentSynchronizedWithCentral
    '        AddHandler evApp.DocumentSynchronizingWithCentral, AddressOf evApp_DocumentSynchronizingWithCentral
    '        AddHandler evApp.DocumentWorksharingEnabled, AddressOf evApp_DocumentWorksharingEnabled
    '        AddHandler evApp.ElementTypeDuplicated, AddressOf evApp_ElementTypeDuplicated
    '        AddHandler evApp.ElementTypeDuplicating, AddressOf evApp_ElementTypeDuplicating
    '        AddHandler evApp.FamilyLoadedIntoDocument, AddressOf evApp_FamilyLoadedIntoDocument
    '        AddHandler evApp.FamilyLoadingIntoDocument, AddressOf evApp_FamilyLoadingIntoDocument
    '        AddHandler evApp.FileExported, AddressOf evApp_FileExported
    '        AddHandler evApp.FileExporting, AddressOf evApp_FileExporting
    '        AddHandler evApp.FileImported, AddressOf evApp_FileImported
    '        AddHandler evApp.FileImporting, AddressOf evApp_FileImporting
    '        AddHandler evApp.LinkedResourceOpened, AddressOf evApp_LinkedResourceOpened
    '        AddHandler evApp.LinkedResourceOpening, AddressOf evApp_LinkedResourceOpening
    '        AddHandler evApp.ProgressChanged, AddressOf evApp_ProgressChanged
    '        AddHandler evApp.ViewExported, AddressOf evApp_ViewExported
    '        AddHandler evApp.ViewExporting, AddressOf evApp_ViewExporting
    '        AddHandler evApp.ViewPrinted, AddressOf evApp_ViewPrinted
    '        AddHandler evApp.ViewPrinting, AddressOf evApp_ViewPrinting
    '        AddHandler evApp.WorksharedOperationProgressChanged, AddressOf evApp_WorksharedOperationProgressChanged
    '        evRevit._SubscritoApplication = True
    '    Catch ex As Exception
    '        evRevit._SubscritoApplication = False
    '    End Try
    'End Sub
    Public Shared Sub SubscribeToDoc(Optional doc As Autodesk.Revit.DB.Document = Nothing)
        'If doc Is Nothing Then doc = evDoc()  ' evAppUI.ActiveUIDocument.Document
        If doc Is Nothing Then Exit Sub ' Salimos si no hay documentos abiertos.
        '
        'AddHandler doc.DocumentClosing, AddressOf evDoc_DocumentClosing
        'AddHandler doc.DocumentPrinted, AddressOf evDoc_DocumentPrinted
        'AddHandler doc.DocumentPrinting, AddressOf evDoc_DocumentPrinting
        'AddHandler doc.DocumentSaved, AddressOf evDoc_DocumentSaved
        'AddHandler doc.DocumentSavedAs, AddressOf evDoc_DocumentSavedAs
        'AddHandler doc.DocumentSaving, AddressOf evDoc_DocumentSaving
        'AddHandler doc.DocumentSavingAs, AddressOf evDoc_DocumentSavingAs
        'AddHandler doc.ViewPrinted, AddressOf evDoc_ViewPrinted
        'AddHandler doc.ViewPrinting, AddressOf evDoc_ViewPrinting
    End Sub
#End Region

#Region "UNSUBSCRIBE"
    Public Shared Sub UnsubscribeAll()
        'UnsubscribeForControllerApplication()
        UnsubscribeForComponentManager()
        'UnsubscribeForUIApplication()
        'UnsubscribeForApplication()
        'UnsubscribeFromDoc()
    End Sub
    Public Shared Sub UnsubscribeForControllerApplication()
        'If evApp Is Nothing Then
        '    TaskDialog.Show("NOTICE TO USER", "evApp is Nothing")
        '    Exit Sub
        'End If
        'Try
        '    RemoveHandler evAppCont.ApplicationInitialized, AddressOf evAppCont_ApplicationInitialized
        '    RemoveHandler evAppCont.DocumentOpened, AddressOf evAppCont_DocumentOpened
        '    RemoveHandler evAppCont.DocumentOpening, AddressOf evAppCont_DocumentOpening
        '    evRevit._SubscritoControllerApplication = False
        'Catch ex As Exception
        '    '
        'End Try
    End Sub
    Public Shared Sub UnsubscribeForComponentManager()
        Try
            RemoveHandler adWin.ComponentManager.UIElementActivated, AddressOf evCompM_UIElementActivated
            RemoveHandler adWin.ComponentManager.ItemAddedToQuickAccessToolBar, AddressOf evCompM_ItemAddedToQuickAccessToolBar
            RemoveHandler adWin.ComponentManager.ItemExecuted, AddressOf evCompM_ItemExecuted
            RemoveHandler adWin.ComponentManager.ItemInitialized, AddressOf evCompM_ItemInitialized
            RemoveHandler adWin.ComponentManager.PreviewExecute, AddressOf evCompM_PreviewExecute
            RemoveHandler adWin.ComponentManager.PropertyChanged, AddressOf evCompM_PropertyChanged
            RemoveHandler adWin.ComponentManager.ToolTipClosed, AddressOf evCompM_ToolTipClosed
            RemoveHandler adWin.ComponentManager.ToolTipOpened, AddressOf evCompM_ToolTipOpened
            evRevit._SubscritoComponentManager = False
        Catch ex As Exception
            '
        End Try
    End Sub
    'Public Shared Sub UnsubscribeForUIApplication()
    '    If evAppUI Is Nothing Then
    '        TaskDialog.Show("NOTICE TO USER", "UnsubscribeForUIApplication-->evAppUI is Nothing")
    '        Exit Sub
    '    End If
    '    Try
    '        RemoveHandler evAppUI.ViewActivated, AddressOf evAppUI_ViewActivated
    '        RemoveHandler evAppUI.ApplicationClosing, AddressOf evAppUI_ApplicationClosing
    '        RemoveHandler evAppUI.DialogBoxShowing, AddressOf evAppUI_DialogBoxShowing
    '        RemoveHandler evAppUI.DisplayingOptionsDialog, AddressOf evAppUI_DisplayingOptionsDialog
    '        RemoveHandler evAppUI.DockableFrameFocusChanged, AddressOf evAppUI_DockableFrameFocusChanged
    '        RemoveHandler evAppUI.DockableFrameVisibilityChanged, AddressOf evAppUI_DockableFrameVisibilityChanged
    '        RemoveHandler evAppUI.FabricationPartBrowserChanged, AddressOf evAppUI_FabricationPartBrowserChanged
    '        'RemoveHandler evAppUI.Idling, AddressOf evAppUI_Idling
    '        RemoveHandler evAppUI.Idling, AddressOf evAppUI_Idling_Handler_NOAutorizado
    '        RemoveHandler evAppUI.TransferredProjectStandards, AddressOf evAppUI_TransferredProjectStandards
    '        RemoveHandler evAppUI.TransferringProjectStandards, AddressOf evAppUI_TransferringProjectStandards
    '        RemoveHandler evAppUI.ViewActivating, AddressOf evAppUI_ViewActivating
    '        evRevit._SubscritoUIApplication = False
    '    Catch ex As Exception
    '        '
    '    End Try
    'End Sub
    'Public Shared Sub UnsubscribeForApplication()
    '    If evApp Is Nothing Then
    '        TaskDialog.Show("NOTICE TO USER", "UnsubscribeForApplication-->evApp is Nothing")
    '        Exit Sub
    '    End If
    '    Try
    '        'RemoveHandler evApp.ApplicationInitialized, AddressOf evApp_ApplicationInitialized
    '        RemoveHandler evApp.DocumentChanged, AddressOf evApp_DocumentChanged
    '        RemoveHandler evApp.DocumentClosing, AddressOf evApp_DocumentClosing
    '        RemoveHandler evApp.DocumentClosed, AddressOf evApp_DocumentClosed
    '        RemoveHandler evApp.DocumentCreated, AddressOf evApp_DocumentCreated
    '        RemoveHandler evApp.DocumentCreating, AddressOf evApp_DocumentCreating
    '        '
    '        'RemoveHandler evApp.DocumentOpened, AddressOf evApp_DocumentOpened     'Usar el de ControllerApplication
    '        'RemoveHandler evApp.DocumentOpening, AddressOf evAppt_DocumentOpening  'Usar el de ControllerApplication
    '        RemoveHandler evAppC.DocumentOpened, AddressOf evAppCont_DocumentOpened
    '        RemoveHandler evAppC.DocumentOpening, AddressOf evAppCont_DocumentOpening
    '        '
    '        RemoveHandler evApp.DocumentSaving, AddressOf evApp_DocumentSaving
    '        RemoveHandler evApp.DocumentSaved, AddressOf evApp_DocumentSaved
    '        RemoveHandler evApp.DocumentSavingAs, AddressOf evApp_DocumentSavingAs
    '        RemoveHandler evApp.DocumentSavedAs, AddressOf evApp_DocumentSavedAs
    '        RemoveHandler evApp.FailuresProcessing, AddressOf evApp_FailuresProcessing
    '        ' Eventos sin funciones.
    '        RemoveHandler evApp.DocumentPrinted, AddressOf evApp_DocumentPrinted
    '        RemoveHandler evApp.DocumentPrinting, AddressOf evApp_DocumentPrinting
    '        RemoveHandler evApp.DocumentSynchronizedWithCentral, AddressOf evApp_DocumentSynchronizedWithCentral
    '        RemoveHandler evApp.DocumentSynchronizingWithCentral, AddressOf evApp_DocumentSynchronizingWithCentral
    '        RemoveHandler evApp.DocumentWorksharingEnabled, AddressOf evApp_DocumentWorksharingEnabled
    '        RemoveHandler evApp.ElementTypeDuplicated, AddressOf evApp_ElementTypeDuplicated
    '        RemoveHandler evApp.ElementTypeDuplicating, AddressOf evApp_ElementTypeDuplicating
    '        RemoveHandler evApp.FamilyLoadedIntoDocument, AddressOf evApp_FamilyLoadedIntoDocument
    '        RemoveHandler evApp.FamilyLoadingIntoDocument, AddressOf evApp_FamilyLoadingIntoDocument
    '        RemoveHandler evApp.FileExported, AddressOf evApp_FileExported
    '        RemoveHandler evApp.FileExporting, AddressOf evApp_FileExporting
    '        RemoveHandler evApp.FileImported, AddressOf evApp_FileImported
    '        RemoveHandler evApp.FileImporting, AddressOf evApp_FileImporting
    '        RemoveHandler evApp.LinkedResourceOpened, AddressOf evApp_LinkedResourceOpened
    '        RemoveHandler evApp.LinkedResourceOpening, AddressOf evApp_LinkedResourceOpening
    '        RemoveHandler evApp.ProgressChanged, AddressOf evApp_ProgressChanged
    '        RemoveHandler evApp.ViewExported, AddressOf evApp_ViewExported
    '        RemoveHandler evApp.ViewExporting, AddressOf evApp_ViewExporting
    '        RemoveHandler evApp.ViewPrinted, AddressOf evApp_ViewPrinted
    '        RemoveHandler evApp.ViewPrinting, AddressOf evApp_ViewPrinting
    '        RemoveHandler evApp.WorksharedOperationProgressChanged, AddressOf evApp_WorksharedOperationProgressChanged
    '        evRevit._SubscritoApplication = False
    '    Catch ex As Exception
    '        '
    '    End Try
    'End Sub

    Public Shared Sub UnsubscribeFromDoc(Optional queDoc As Autodesk.Revit.DB.Document = Nothing)
        If evAppUI.ActiveUIDocument Is Nothing Then Exit Sub
        If evAppUI.Application.Documents.Size = 0 Then Exit Sub
        '
        If queDoc Is Nothing Then
            For Each doc As Document In evAppUI.Application.Documents
                Try
                    RemoveHandler doc.DocumentClosing, AddressOf evDoc_DocumentClosing
                    RemoveHandler doc.DocumentPrinted, AddressOf evDoc_DocumentPrinted
                    RemoveHandler doc.DocumentPrinting, AddressOf evDoc_DocumentPrinting
                    RemoveHandler doc.DocumentSaved, AddressOf evDoc_DocumentSaved
                    RemoveHandler doc.DocumentSavedAs, AddressOf evDoc_DocumentSavedAs
                    RemoveHandler doc.DocumentSaving, AddressOf evDoc_DocumentSaving
                    RemoveHandler doc.DocumentSavingAs, AddressOf evDoc_DocumentSavingAs
                    RemoveHandler doc.ViewPrinted, AddressOf evDoc_ViewPrinted
                    RemoveHandler doc.ViewPrinting, AddressOf evDoc_ViewPrinting
                Catch ex As Exception
                    '
                End Try
            Next
        Else
            RemoveHandler queDoc.DocumentClosing, AddressOf evDoc_DocumentClosing
            RemoveHandler queDoc.DocumentPrinted, AddressOf evDoc_DocumentPrinted
            RemoveHandler queDoc.DocumentPrinting, AddressOf evDoc_DocumentPrinting
            RemoveHandler queDoc.DocumentSaved, AddressOf evDoc_DocumentSaved
            RemoveHandler queDoc.DocumentSavedAs, AddressOf evDoc_DocumentSavedAs
            RemoveHandler queDoc.DocumentSaving, AddressOf evDoc_DocumentSaving
            RemoveHandler queDoc.DocumentSavingAs, AddressOf evDoc_DocumentSavingAs
            RemoveHandler queDoc.ViewPrinted, AddressOf evDoc_ViewPrinted
            RemoveHandler queDoc.ViewPrinting, AddressOf evDoc_ViewPrinting
        End If
    End Sub
#End Region
    Public Shared Sub BorraTempAlSalir()
        Try
            '/BorraTempRevit|G:\Temp\TEMPS
            Dim p As Process = New Process()
            'p.WaitForExit(1000 * 40)
            p.PriorityBoostEnabled = True
            Dim psi As ProcessStartInfo = New ProcessStartInfo(IO.Path.Combine(modVar._dirApp, "task2aCAD.exe"))
            psi.UseShellExecute = True
            psi.Arguments = "/BorraTempRevit|G:\Temp\TEMPS"
            psi.WindowStyle = ProcessWindowStyle.Hidden
            'psi.CreateNoWindow = True
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception

        End Try
    End Sub
    '
    Public Enum queSubscribo
        ControllerApplication
        ComponentManager
        UIApplication
        Application
        Todo
    End Enum
End Class