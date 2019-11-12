Imports System.Collections
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Text
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports System.Xml
Imports Autodesk.Revit.ApplicationServices
'Imports Autodesk.Revit.Utility
Imports Autodesk.Revit.ApplicationServices.Application
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Architecture
Imports Autodesk.Revit.DB.Document
Imports Autodesk.Revit.DB.Electrical
Imports Autodesk.Revit.DB.Events
Imports Autodesk.Revit.DB.Mechanical
Imports Autodesk.Revit.DB.Plumbing
Imports Autodesk.Revit.DB.Structure
'Imports Autodesk.Revit.Collections
Imports Autodesk.Revit.Exceptions
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Events
Imports Autodesk.Revit.UI.Selection
Imports ULMAStudio.ULMAStudioApplication

Partial Public Class evRevit
    Public Shared Sub evAppCont_ApplicationInitialized(sender As Object, e As ApplicationInitializedEventArgs) Handles evAppC.ApplicationInitialized
        ' Aquí pondremos lo que queremos que carge, después de tener Revit inicializado
        ' ApplicationInitialized de Application no se activa con doble click en fichero desde explorer, este sí.
        If evApp Is Nothing Then
            Try
                evApp = TryCast(sender, Autodesk.Revit.ApplicationServices.Application)
                evAppUI = New UIApplication(evApp)
                evRevit.SubscribeAll()
                ' Si hay actualizaciones, ejecutar el botón About
                'If cLcsv._ItHasUpdate = True Then
                '    BotonULMAFEjecuta(nombrePanelAbout, "btnAbout")
                'End If
            Catch ex As Exception
                '
            End Try
        End If
    End Sub
    Private Shared Sub evAppCont_DocumentOpened(sender As Object, e As DocumentOpenedEventArgs) Handles evAppC.DocumentOpened
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentOpened")
        If evApp Is Nothing Then
            Try
                evApp = TryCast(sender, Autodesk.Revit.ApplicationServices.Application)
                evAppUI = New UIApplication(evApp)
                evRevit.SubscribeAll()
            Catch ex As Exception
                '
            End Try
        End If
        '
        'If evRevit.evOpen = False Then Exit Sub
        ' Si tiene links a proyectos (Solo primer nivel) crear y revincular en Temp
        'AdskApplication.EventoReLinkTodos.Raise()
        If e.IsCancelled Then Exit Sub
        '
        If e.Document Is Nothing Then Exit Sub
        If evRevit.evAppUI.ActiveUIDocument.Document IsNot Nothing AndAlso evRevit.evAppUI.ActiveUIDocument.Document.PathName = e.Document.PathName Then Exit Sub
        '
        'PonLog("ABRIENDO (" & e.Document.GetType.ToString & ") " & e.Document.PathName)
        '
        ' Para poner todo lo que necesitamos de ULMA
        ' Unidades, fichero de parametros compartidos, etc.
        'utilesRevit.InicioAplicacionObligatorio(shared_parameter_file)
        ' Crear los parametros del proyecto desde los parametros compartidos.
        'modULMA.CrearParametrosProyecto(e.Document)
        ' Copiamos las Tablas de planificación #BOMxxx
        'SchedulesCopiaDesdePlantilla(e.Document)
        ' Cargamos todas las familias que empiezas por GEN... y U_...
        'modULMA.FamiliasCargaDesdePlantilla(e.Document)
    End Sub
    Private Shared Sub evAppCont_DocumentOpening(sender As Object, e As DocumentOpeningEventArgs) Handles evAppC.DocumentOpening
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentOpening")
        'If evRevit.evOpen = False Then Exit Sub
        If evApp Is Nothing Then
            Try
                evApp = TryCast(sender, Autodesk.Revit.ApplicationServices.Application)
                evAppUI = New UIApplication(evApp)
                evRevit.SubscribeAll()
            Catch ex As Exception
                '
            End Try
        End If
    End Sub
End Class
