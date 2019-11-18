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
Partial Public Class evRevit

#Region "EVENTOS CONTROLLEDAPLICACION"
    Public Shared Sub evAppC_ApplicationInitialized(sender As Object, e As ApplicationInitializedEventArgs) Handles evAppC.ApplicationInitialized
        'System.Windows.MessageBox.Show("ControlledApplication.ApplicationInitialized")
        ' Aquí pondremos lo que queremos que carge, después de tener Revit inicializado
        ' ApplicationInitialized de Application no se activa con doble click en fichero desde explorer, este sí.
        If evAppUI Is Nothing Then
            Try
                evAppUI = New UIApplication(TryCast(sender, Autodesk.Revit.ApplicationServices.Application))
                'AddHandler evAppUI.Idling, AddressOf evAppUI_Idling
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        End If
        If evApp Is Nothing Then
            Try
                evApp = TryCast(sender, Autodesk.Revit.ApplicationServices.Application)
            Catch ex As Exception
                System.Windows.MessageBox.Show(ex.ToString)
            End Try
        End If
        evRevit.SubscribeAll()
    End Sub

    Private Shared Sub evAppC_DocumentChanged(sender As Object, e As DocumentChangedEventArgs) Handles evAppC.DocumentChanged
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentChanged")
        If e.GetDeletedElementIds.Count = 0 AndAlso e.GetAddedElementIds.Count = 0 Then ' And e.GetModifiedElementIds.Count = 0 Then
            Exit Sub
        End If
        If e.GetDocument.IsFamilyDocument OrElse actua = False Then
            If oIdsFamI IsNot Nothing Then oIdsFamI = Nothing
            If oIdsAnoI IsNot Nothing Then oIdsAnoI = Nothing
            Exit Sub
        End If
        '
        '
        If oIdsFamI Is Nothing Then oIdsFamI = New List(Of ElementId)
        If oIdsAnoI Is Nothing Then oIdsAnoI = New List(Of ElementId)
        ' Quitar los Ids que estuvieran en oIdsFamI, si se han borrado.
        For Each eid As ElementId In e.GetDeletedElementIds
            If oIdsFamI.Contains(eid) Then oIdsFamI.Remove(eid)
        Next
        Dim oFInsertadas As New List(Of ElementId)
        ' Anadir los Ids que se hayan añadido.
        For Each eId As ElementId In e.GetAddedElementIds()
            Dim oFi As Element = e.GetDocument.GetElement(eId)
            If TypeOf oFi Is FamilyInstance Then
                Dim oFins As FamilyInstance = CType(oFi, FamilyInstance)
                ' Añadir el objeto Family a la colección.
                Try
                    If oFInsertadas.Contains(oFins.Symbol.Family.Id) = False AndAlso oFins.SuperComponent Is Nothing Then oFInsertadas.Add(oFins.Symbol.Family.Id)
                Catch ex As Exception
                    Debug.Print(ex.ToString)
                End Try
                '
                ' Solo procesaremos los que: Category = typeFamily o typeSet (Modelos Genéricos ambos)  y Manufacturer contiene ULMA
                Dim manu As String = ParametroFamilyInstanceLeeBuiltInParameter(evRevit.evAppUI.ActiveUIDocument.Document, CType(oFi, FamilyInstance), BuiltInParameter.ALL_MODEL_MANUFACTURER)
                If manu.ToUpper.Contains("ULMA") And
                    (oFi.Category.Id.IntegerValue.Equals(CInt(typeFamily)) OrElse
                     oFi.Category.Id.IntegerValue.Equals(CInt(typeSet))) Then
                    oIdsFamI.Add(eId)
                ElseIf (oFi.Category.Id.IntegerValue.Equals(CInt(BuiltInCategory.OST_GenericAnnotation))) Then
                    oIdsAnoI.Add(eId)
                End If
            End If
        Next
        ' Añadir al Log las familias insertadas. (Lo quitamos, ya lo registra UCBrowser)
        'For Each oIdins As ElementId In oFInsertadas
        '    Dim oFins As Family = CType(e.GetDocument.GetElement(oIdins), Family)
        '    If ULMALGFree.clsBase._registraLoadInsert = True Then   ' AndAlso oFins.Symbol.FamilyName.Contains("#") Then
        '        If oFins.Document IsNot Nothing AndAlso
        '            oFins.Document.IsFamilyDocument AndAlso
        '            oFins.Document.PathName <> "" _
        '            AndAlso IO.File.Exists(oFins.Document.PathName) Then
        '            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("INSERT_FAMILY", oFins.Document.PathName, oFins.Name)
        '        Else
        '            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("INSERT_FAMILY", "", oFins.Name)
        '        End If
        '    End If
        'Next
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
        ULMALGFree.clsBase._registraLoadInsert = True
    End Sub

    Private Shared Sub evAppC_DocumentClosed(sender As Object, e As DocumentClosedEventArgs) Handles evAppC.DocumentClosed
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentClosed")
        'If evRevit.evClose = False Then Exit Sub
        If evApp IsNot Nothing Then
            evApp.PurgeReleasedAPIObjects()
        End If
    End Sub

    Private Shared Sub evAppC_DocumentClosing(sender As Object, e As DocumentClosingEventArgs) Handles evAppC.DocumentClosing
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentClosing")
        'If evRevit.evClose = False Then Exit Sub
        ' Log en DocumentClosing, porque en DocumentClosed ya no se tiene acceso a el Document
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO

        If e.Document.IsFamilyDocument = True Then
            If CerrandoRevit = False Then
                If e.Document.PathName.ToUpper.EndsWith("RFT") Then
                    If cLcsv IsNot Nothing AndAlso registraLog = True Then cLcsv.PonLog_ULMA("CLOSE_FAMILY_TEMPLATE", FILENAME:=e.Document.PathName, NOTES:=IO.Path.GetFileName(e.Document.PathName))
                Else
                    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("CLOSE_FAMILY", e.Document.PathName, IO.Path.GetFileName(e.Document.PathName))
                End If
            End If
        Else
            If e.Document.PathName.ToUpper.EndsWith("RTE") Then
                If cLcsv IsNot Nothing AndAlso registraLog = True Then cLcsv.PonLog_ULMA("CLOSE_PROJECT_TEMPLATE", FILENAME:=e.Document.PathName, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            Else
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("CLOSE_PROJECT", FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
                'Else
                '    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("CLOSE_FILE", e.Document.PathName)
            End If
        End If
    End Sub

    Private Shared Sub evAppC_DocumentCreated(sender As Object, e As DocumentCreatedEventArgs) Handles evAppC.DocumentCreated
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentCreated")
        If e.Document IsNot Nothing Then SubscribeToDoc(e.Document)
    End Sub

    Private Shared Sub evAppC_DocumentCreating(sender As Object, e As DocumentCreatingEventArgs) Handles evAppC.DocumentCreating
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentCreating")
        If e.Template <> "" AndAlso IO.Path.GetFileName(e.Template) <> "" Then
            If e.DocumentType = DocumentType.Template Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.NEW_FAMILY_TEMPLATE, FILENAME:=e.Template, MARKET:=arrM, LANGUAGE:=arrL, NOTES:="Template = " & IO.Path.GetFileName(e.Template))
            ElseIf evRevit.newQue <> "" Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.NEW_PROJECT_TEMPLATE, FILENAME:=e.Template, MARKET:=arrM, LANGUAGE:=arrL, NOTES:="Template = " & IO.Path.GetFileName(e.Template))
                evRevit.newQue = ""
            ElseIf evRevit.newQue = "" Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.NEW_PROJECT_TEMPLATE, FILENAME:=e.Template, MARKET:=arrM, LANGUAGE:=arrL, NOTES:="Template = " & IO.Path.GetFileName(e.Template))
            End If
        End If
    End Sub

    Public Shared Sub evAppC_DocumentOpened(sender As Object, e As DocumentOpenedEventArgs) Handles evAppC.DocumentOpened
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentOpened-->Inicicado")
        'Dim mensaje As String = "ControlledApplication.DocumentOpened-->Inicicado" & vbCrLf
        'If e.Document IsNot Nothing Then mensaje &= "PathName = " & e.Document.PathName & vbCrLf
        'mensaje &= e.Status.ToString
        anularWARNINGS = True
        '
        If e.Status = RevitAPIEventStatus.Succeeded Then
            anularWARNINGS = False
            'System.Windows.MessageBox.Show(mensaje)
        End If
        '
        If evAppUI Is Nothing Then
            Try
                evAppUI = New UIApplication(TryCast(sender, Autodesk.Revit.ApplicationServices.Application))
                'evRevit.SubscribeAll()
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        End If
        ' Si se ha cancelado en Opening, salimos.
        If e.IsCancelled Then Exit Sub
        ' Solo si no estamos en proceso Automático de Traducción y Codificación.
        If e.Status = RevitAPIEventStatus.Succeeded Then    ' AndAlso (frmT Is Nothing AndAlso frmC Is Nothing) Then ' e.Document.IsFamilyDocument = False Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If e.Document.PathName.ToUpper.EndsWith("RFA") Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_FAMILY, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf e.Document.PathName.ToUpper.EndsWith("RVT") Then
                ' Rellenar la configuración por defecto, cada vez que abrimos un fichero.
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_PROJECT, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf e.Document.PathName.ToUpper.EndsWith("RFT") Then
                If cLcsv IsNot Nothing AndAlso registraLog = True Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_FAMILY_TEMPLATE, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf e.Document.PathName.ToUpper.EndsWith("RTE") Then
                If ULMALGFree.clsBase._ultimaAccion <> ULMALGFree.ACTION.UCR_CODIFY.ToString AndAlso ULMALGFree.clsBase._ultimaAccion <> ULMALGFree.ACTION.UCR_TRANSLATE.ToString Then
                    If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_PROJECT_TEMPLATE, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
                End If
            ElseIf e.Document.PathName.ToUpper.EndsWith("IFC") Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("OPEN_IFC", FILENAME:=IO.Path.GetFileName(e.Document.PathName), MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf e.Document.PathName.ToUpper.EndsWith("ADSK") Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("OPEN_ADSK", FILENAME:=IO.Path.GetFileName(e.Document.PathName), MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf e.Document.PathName.ToUpper.EndsWith("SAT") Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("OPEN_SAT", FILENAME:=IO.Path.GetFileName(e.Document.PathName), MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            End If
        End If
        evRevit.SubscribeAll()
    End Sub
    Public Shared Sub evAppC_DocumentOpening(sender As Object, e As DocumentOpeningEventArgs) Handles evAppC.DocumentOpening
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentOpening-->Iniciado")
        If evAppUI Is Nothing Then
            Try
                evAppUI = New UIApplication(TryCast(sender, Autodesk.Revit.ApplicationServices.Application))
            Catch ex As Exception
                System.Windows.MessageBox.Show(ex.ToString)
            End Try
        End If
    End Sub

    Private Shared Sub evAppC_DocumentSaved(sender As Object, e As DocumentSavedEventArgs) Handles evAppC.DocumentSaved
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSaved")
        'If evRevit.evSave = False Then Exit Sub
        If e.Status = RevitAPIEventStatus.Succeeded Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.SAVE, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
        End If
    End Sub

    Private Shared Sub evAppC_DocumentSavedAs(sender As Object, e As DocumentSavedAsEventArgs) Handles evAppC.DocumentSavedAs
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSavedAs")
        'e.Documento    'Documento que ha generado el evento. El que acabamos de GuardarComo
        'e.OriginalPath 'El Path original desde el que se guardó esta nueva copia.
        'If evSave = False Then Exit Sub
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.SAVEAS,
        '                  "File : " & IO.Path.GetFileName(e.OriginalPath) & " SaveAs : " & IO.Path.GetFileName(e.Document.PathName))
        If e.Status = RevitAPIEventStatus.Succeeded Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If SaveAsLibrary = False Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.SAVEAS, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            ElseIf SaveAsLibrary = True Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.SAVEAS_LIBRARY_FAMILY, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.Document.PathName))
            End If
        End If
    End Sub

    Private Shared Sub evAppC_DocumentSaving(sender As Object, e As DocumentSavingEventArgs) Handles evAppC.DocumentSaving
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSaving")
        'If evRevit.evSave = False Then Exit Sub
    End Sub

    Private Shared Sub evAppC_DocumentSavingAs(sender As Object, e As DocumentSavingAsEventArgs) Handles evAppC.DocumentSavingAs
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSavingAs")
    End Sub

    Private Shared Sub evAppC_FailuresProcessing(sender As Object, e As FailuresProcessingEventArgs) Handles evAppC.FailuresProcessing
        ' System.Windows.MessageBox.Show("ControlledApplication.FailuresProcessing")
        '' Anularemos los mensajes de aviso de nuestra aplicación. Si enejecucion = True
        If anularWARNINGS = True OrElse FormulariosDesarrolloAbiertos() = True Then
            Dim oFa As FailuresAccessor = e.GetFailuresAccessor
            'PonLog(oFa.GetSeverity.ToString & ": " & oFa.GetDocument.PathName)
            If oFa.GetSeverity = FailureSeverity.Warning Then
                'PonLog("FailureSeverity.Warning : " & evDoc.PathName)
                oFa.DeleteAllWarnings()
            ElseIf oFa.GetSeverity = FailureSeverity.Error Then
                'PonLog("FailureSeverity.Error : " & evDoc.PathName)
                oFa.ResolveFailure(oFa.GetFailureMessages.FirstOrDefault)
                e.SetProcessingResult(FailureProcessingResult.ProceedWithCommit)
            End If
        End If
    End Sub

    Private Shared Sub evAppC_FamilyLoadingIntoDocument(sender As Object, e As FamilyLoadingIntoDocumentEventArgs) Handles evAppC.FamilyLoadingIntoDocument
        'System.Windows.MessageBox.Show("ControlledApplication.FamilyLoadingIntoDocument")
        'If ULMALGFree.clsBase._registraLoadInsert = True Then

        'End If
    End Sub

    Private Shared Sub evAppC_FamilyLoadedIntoDocument(sender As Object, e As FamilyLoadedIntoDocumentEventArgs) Handles evAppC.FamilyLoadedIntoDocument
        'System.Windows.MessageBox.Show("ControlledApplication.FamilyLoadedIntoDocument")
        ' ALBERTO. Para que no registre 2 veces lo que viene del Browser
        If e.Status = RevitAPIEventStatus.Succeeded AndAlso ULMALGFree.clsBase._registraLoadInsert = True Then
            Dim strPath As String = e.FamilyPath
            If strPath = "" AndAlso e.Document IsNot Nothing AndAlso e.Document.PathName <> "" Then
                strPath = e.Document.PathName
            End If
            If strPath <> "" Then   ' AndAlso IO.File.Exists(e.FamilyPath) Then
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.LOAD_FAMILY, strPath, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=e.FamilyName)
            Else
                If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.LOAD_FAMILY, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=e.FamilyName)
            End If
        End If
        ULMALGFree.clsBase._registraLoadInsert = True
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.UCREVIT
    End Sub
    Private Shared Sub evAppC_DocumentPrinted(sender As Object, e As DocumentPrintedEventArgs) Handles evAppC.DocumentPrinted
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentPrinted")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.PRINT_DOCUMENT, FILENAME:=IO.Path.GetFileName(e.Document.PathName), MARKET:=arrM, LANGUAGE:=arrL)
        End If
    End Sub
    Private Shared Sub evAppC_DocumentSynchronizedWithCentral(sender As Object, e As DocumentSynchronizedWithCentralEventArgs) Handles evAppC.DocumentSynchronizedWithCentral
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSynchronizedWithCentral")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.SYNCHRONIZE_DOCUMENT, FILENAME:=e.Document.PathName, MARKET:=arrM, LANGUAGE:=arrL)
        End If
    End Sub
    '
    'DWG DWG format  
    'DWF DWF format  
    'DWFX DWFX format  
    'GBXML GBXML format  
    'FBX FBX format  
    'Image Image format  
    'DGN DGN format  
    'Civil3D Civel3D format  
    'Inventor Inventor format  
    'DXF DXF format  
    'SAT SAT format  
    'IFC IFC format  
    'NWC Navisworks format  
    Private Shared Sub evAppC_FileExported(sender As Object, e As FileExportedEventArgs) Handles evAppC.FileExported
        'System.Windows.MessageBox.Show("ControlledApplication.FileExported")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("EXPORT_" & IO.Path.GetExtension(e.Path).ToUpper.Replace(".", ""), FILENAME:=e.Path, MARKET:=arrM, LANGUAGE:=arrL)   ' IO.Path.GetFileName(e.Document.PathName))
        End If
    End Sub
    Private Shared Sub evAppC_FileImported(sender As Object, e As FileImportedEventArgs) Handles evAppC.FileImported
        'System.Windows.MessageBox.Show("ControlledApplication.FileImported")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("IMPORT_" & IO.Path.GetExtension(e.Path).ToUpper.Replace(".", ""), FILENAME:=e.Path, MARKET:=arrM, LANGUAGE:=arrL)    ' IO.Path.GetFileName(e.Document.PathName))
        End If
    End Sub
    Private Shared Sub evAppC_LinkedResourceOpened(sender As Object, e As LinkedResourceOpenedEventArgs) Handles evAppC.LinkedResourceOpened
        'System.Windows.MessageBox.Show("ControlledApplication.LinkedResourceOpened")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.OPEN_FILELINK, FILENAME:=e.LinkedResourcePathName, MARKET:=arrM, LANGUAGE:=arrL, NOTES:=IO.Path.GetFileName(e.LinkedResourcePathName))   ' & " in Document : " & IO.Path.GetFileName(e.Document.PathName))
        End If
    End Sub
    Private Shared Sub evAppC_ViewPrinted(sender As Object, e As ViewPrintedEventArgs) Handles evAppC.ViewPrinted
        'System.Windows.MessageBox.Show("ControlledApplication.ViewPrinted")
        If e.Status = RevitAPIEventStatus.Succeeded Then
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.PRINT_VIEW, MARKET:=arrM, LANGUAGE:=arrL, NOTES:="View : " & e.View.Name) '& " in Document : " & IO.Path.GetFileName(e.Document.PathName))
        End If
    End Sub
#End Region

#Region "EVENTOS CONTROLLEDAPPLICATION DESACTIVADOS"

    Private Shared Sub evAppC_DocumentPrinting(sender As Object, e As DocumentPrintingEventArgs) Handles evAppC.DocumentPrinting
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentPrinting")
    End Sub

    Private Shared Sub evAppC_DocumentSynchronizingWithCentral(sender As Object, e As DocumentSynchronizingWithCentralEventArgs) Handles evAppC.DocumentSynchronizingWithCentral
        'System.Windows.MessageBox.Show("ControlledApplication.DocumentSynchronizingWithCentral")
    End Sub

    Private Shared Sub evAppC_ElementTypeDuplicated(sender As Object, e As ElementTypeDuplicatedEventArgs) Handles evAppC.ElementTypeDuplicated
        'System.Windows.MessageBox.Show("ControlledApplication.ElementTypeDuplicated")
    End Sub

    Private Shared Sub evAppC_ElementTypeDuplicating(sender As Object, e As ElementTypeDuplicatingEventArgs) Handles evAppC.ElementTypeDuplicating
        'System.Windows.MessageBox.Show("ControlledApplication.ElementTypeDuplicating")
    End Sub

    Private Shared Sub evAppC_FileExporting(sender As Object, e As FileExportingEventArgs) Handles evAppC.FileExporting
        'System.Windows.MessageBox.Show("ControlledApplication.FileExporting")
    End Sub


    Private Shared Sub evAppC_FileImporting(sender As Object, e As FileImportingEventArgs) Handles evAppC.FileImporting
        'System.Windows.MessageBox.Show("ControlledApplication.FileImporting")
    End Sub

    Private Shared Sub evAppC_LinkedResourceOpening(sender As Object, e As LinkedResourceOpeningEventArgs) Handles evAppC.LinkedResourceOpening
        'System.Windows.MessageBox.Show("ControlledApplication.LinkedResourceOpening")
    End Sub
    Public Shared Sub evAppCont_ProgressChanged(sender As Object, e As Autodesk.Revit.DB.Events.ProgressChangedEventArgs) Handles evAppC.ProgressChanged
        'If e.Position = e.LowerRange Then
        '    System.Windows.MessageBox.Show("ControlledApplication.ProgressChanged")
        'End If
        'If e.Caption.ToLower.StartsWith("guardar") OrElse e.Caption.ToLower.StartsWith("save") Then
        '    Debug.Print(e.Caption)
        'End If
    End Sub

    Private Shared Sub evAppC_ViewPrinting(sender As Object, e As ViewPrintingEventArgs) Handles evAppC.ViewPrinting
        'System.Windows.MessageBox.Show("ControlledApplication.ViewPrinting")
    End Sub

    Private Shared Sub evAppC_WorksharedOperationProgressChanged(sender As Object, e As WorksharedOperationProgressChangedEventArgs) Handles evAppC.WorksharedOperationProgressChanged
        'System.Windows.MessageBox.Show("ControlledApplication.WorksharedOperationProgressChanged")
    End Sub
#End Region
End Class