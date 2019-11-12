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
Imports uf = ULMALGFree.clsBase

Partial Public Class evRevit
#Region "EVENTOS UIAPPLICATION"
    Private Shared Sub evAppCUI_ApplicationClosing(sender As Object, e As ApplicationClosingEventArgs) Handles evAppUIC.ApplicationClosing
        'System.Windows.MessageBox.Show("UIControlledApplication.ApplicationClosing")
        CerrandoRevit = True
        ' Cerrar Proceso UCRevitFreeReport (Si esta abierto)
        ULMALGFree.clsBase.Process_Close("UCRevitFreeReport")
        'PonLog("***** Salimos de Revit *****" & vbCrLf & vbCrLf)
        If evApp IsNot Nothing Then
            UnsubscribeAll()
        End If
        ' Log, subir y borrar ficheros antes de salir.
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
        If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.CLOSE_REVIT)
        If log Then PonLog_BASICO(ULMALGFree.clsBase._appLogBaseFichero, "End REVIT")
        Try
            If cLcsv IsNot Nothing Then Call cLcsv.CompruebaConexionFTPUlma(True, False, True)
        Catch ex As Exception
            Debug.Print(ex.ToString)
        End Try
        ' ACTUALIZACIONES
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMAUPDATERADDIN
        While cLcsv.terminado = False
            'System.Windows.Forms.Application.DoEvents()
            ' Quedarse aquí hast aque termine todo. Para que no salga de Revit.
        End While
        'If actualizar = True Then
        '    actualizar = False
        '    Dim cUp As Dictionary(Of String, List(Of String)) = ULMALGFree.clsBase.FTP_DameActualizacionesTodas(todas:=False)
        '    If cUp("addins").Count > 0 OrElse cUp("families").Count > 0 OrElse cUp("xmls").Count > 0 Then
        '        frmU = New frmHayUpdates
        '        frmU.Interactivo = True
        'ULMALGFree.clsBase.Bat_CreaEjecuta(cerrarRevit:=False)
        '    End If
        'End If
    End Sub
    Private Shared Sub evAppCUI_DialogBoxShowing(sender As Object, e As DialogBoxShowingEventArgs) Handles evAppUIC.DialogBoxShowing
        'System.Windows.MessageBox.Show("UIControlledApplication.DialogBoxShowing-->Iniciado")
        Dim dialogId As String = e.DialogId     ' Nombre del Dialogo
        If anularWARNINGS = True Then
            If dialogId = "TaskDialog_Unresolved_References" Then
                e.OverrideResult(DialogResult.Cancel)
            Else
                e.OverrideResult(DialogResult.OK)
            End If
        End If
    End Sub
    Public Shared Sub evAppCUI_Idling_LlenaObjectos(sender As Object, e As IdlingEventArgs)
        RemoveHandler evAppUIC.Idling, AddressOf evAppCUI_Idling_LlenaObjectos
        If evAppUI Is Nothing Then evAppUI = TryCast(sender, UIApplication)
        If evApp Is Nothing AndAlso evAppUI IsNot Nothing Then evApp = evAppUI.Application
    End Sub

    Public Shared Sub evAppCUI_Idling(sender As Object, e As IdlingEventArgs) Handles evAppUIC.Idling
        'System.Windows.MessageBox.Show("UIControlledApplication.Idling-->Iniciado")
        If CerrandoRevit = True Then Exit Sub
        ' Llenar objetos 
        If evApp Is Nothing Then
            Try
                evApp = TryCast(sender, Autodesk.Revit.UI.UIApplication).Application
                evRevit.SubscribeAll()
            Catch ex As Exception
                '
            End Try
        End If
        '
        ' ***** COMPROBAR BOTONES Y ACTUALIZACIONES
        If (ultimacomprobacion = 0) Then ' OrElse DateTime.Now.Ticks >= New Date(ultimacomprobacion).AddSeconds(checkupdatesinseconds).Ticks) Then
            Call uf.INIUpdates_LeeTODO()
            uf._ActualizarFamilies = False
            uf._ActualizarAddIns = False
            uf._ActualizarXMLs = False                '
            If uf.cUp("addins").Count > 0 OrElse ULMALGFree.clsBase.cUp("families").Count > 0 OrElse ULMALGFree.clsBase.cUp("xmls").Count > 0 Then
                If uf.cUp("families").Count > 0 Then uf._ActualizarFamilies = True
                If uf.cUp("addins").Count > 0 Then uf._ActualizarAddIns = True
                If uf.cUp("xmls").Count > 0 Then uf._ActualizarXMLs = True                '
                ultimacomprobacion = DateTime.Now.Ticks
                ULMAStudioApplication.Botones_ActualizaEstadoActualizaciones(True)
            End If
        End If
        '' Quitamos esto. No queremos que ofrezca actualizaciones al inicio. Ni que abra el formulario.
        'If (ultimacomprobacion = 0) Then ' OrElse DateTime.Now.Ticks >= New Date(ultimacomprobacion).AddSeconds(checkupdatesinseconds).Ticks) Then
        '    ULMALGFree.clsBase.cUp = ULMALGFree.clsBase.FTP_DameActualizacionesTodas(todas:=False)
        '    If uf.cUp("addins").Count > 0 OrElse ULMALGFree.clsBase.cUp("families").Count > 0 OrElse ULMALGFree.clsBase.cUp("xmls").Count > 0 Then
        '        frmU = New frmHayUpdates
        '        'Dim mensaje As String = "There are updates to download ¿Close REVIT and update?" & vbCrLf & vbCrLf
        '        If uf.cUp("families").Count > 0 Then
        '            frmU.LblFamilies.Text = String.Join(vbCrLf, uf.cUp("families").ToArray) : frmU.LblFamilies.Refresh()
        '            uf._ActualizarFamilies = True
        '        End If
        '        If uf.cUp("addins").Count > 0 Then
        '            frmU.LblAddin.Text = String.Join(vbCrLf, uf.cUp("addins").ToArray) : frmU.LblAddin.Refresh()
        '            uf._ActualizarAddIns = True
        '        End If
        '        If uf.cUp("xmls").Count > 0 Then
        '            frmU.LblAddin.Text &= vbCrLf & "XML: " & String.Join(vbCrLf, uf.cUp("xmls").ToArray) : frmU.LblAddin.Refresh()
        '            uf._ActualizarXMLs = True
        '        End If
        '        ' Enviar los datos al formulario, antes de abrirlo.
        '        frmU.ActualizarFormulario(uf._ActualizarFamilies, (uf._ActualizarAddIns Or uf._ActualizarXMLs))
        '        '
        '        ultimacomprobacion = DateTime.Now.Ticks
        '        If frmU.ShowDialog(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle)) = System.Windows.Forms.DialogResult.Yes Then
        '            actualizar = False
        '            'frmU.Interactivo = False
        '            'ULMALGFree.clsBase.Bat_CreaEjecuta(cerrarRevit:=True)
        '        Else
        '            actualizar = True
        '        End If
        '        UCRevitFreeApplication.Botones_ActualizaEstadoActualizaciones()
        '    End If
        'End If

        '************* Salir en determinadas condiciones. *********
        ' No quedan ficheros abiertos. Borrar tmplog y salir
        If evRevit.evAppUI.ActiveUIDocument Is Nothing Then Exit Sub
        If evRevit.evAppUI.ActiveUIDocument.Document.IsFamilyDocument Then Exit Sub
        If evRevit.evAppUI.ActiveUIDocument.Document.IsReadOnly Then Exit Sub
        '**********************************************************
        '
        ' QUITADO. No haría falta en el Free. Poner ITEM_MARKET en las FamilyInstaces insertadas.
        'Dim errores As String = ""
        'If oIdsFamI.Count > 0 OrElse oIdsAnoI.Count > 0 Then
        '    'If oDoc.PathName = "" Then Exit Sub
        '    Try
        '        If (arrM(0) <> "" AndAlso IsNumeric(arrM(0))) Then
        '            Using transaction As New Transaction(evRevit.evAppUI.ActiveUIDocument.Document, "Write ITEM_MARKET in FamilyInstance")
        '                transaction.Start()
        '                For Each queId As ElementId In oIdsFamI
        '                    Try
        '                        Dim oEle As Element = evRevit.evAppUI.ActiveUIDocument.Document.GetElement(queId)
        '                        'PonLog("Poner ITEM_MARKET a " & CType(oEle, FamilyInstance).Symbol.Name)
        '                        errores &=
        '                        utilesRevit.ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, oEle, "ITEM_MARKET", CInt(arrM(0)))   ' Math.Round(CDbl(arrM123(0)), 0))
        '                    Catch ex As Exception
        '                        Continue For
        '                    End Try
        '                Next
        '                '
        '                ' Añadimos bucle para cambiar ITEM_MARKET también en las anotaciones.
        '                For Each queId As ElementId In oIdsAnoI
        '                    Try
        '                        Dim oEle As Element = evRevit.evAppUI.ActiveUIDocument.Document.GetElement(queId)
        '                        'PonLog("Poner ITEM_MARKET a " & CType(oEle, FamilyInstance).Symbol.Name)
        '                        errores &=
        '                        utilesRevit.ParametroElementEscribe(evRevit.evAppUI.ActiveUIDocument.Document, oEle, "ITEM_MARKET", CInt(arrM(0)))   ' Math.Round(CDbl(arrM123(0)), 0))
        '                    Catch ex As Exception
        '                        Continue For
        '                    End Try
        '                Next
        '                If errores = "" Then
        '                    transaction.Commit()
        '                Else
        '                    transaction.RollBack()
        '                End If
        '            End Using
        '        End If
        '    Catch ex As Exception
        '        'TaskDialog.Show("ATTENTION", ex.Message)
        '    End Try
        '    '' Vaciar la colección de ids para que no entre en este procedimiento.
        '    oIdsFamI.Clear()
        '    oIdsFamI = Nothing
        '    oIdsAnoI.Clear()
        '    oIdsAnoI = Nothing
        'End If
        '
        evApp.PurgeReleasedAPIObjects()
    End Sub
    Private Shared Sub evAppCUI_ViewActivated(sender As Object, e As ViewActivatedEventArgs) Handles evAppUIC.ViewActivated
        'System.Windows.MessageBox.Show("UIControlledApplication.ViewActivated")
        ' Para que cuando activemos la vista de un documento, suscribir los eventos de aplicación (Solo la primera vez)
        If evApp Is Nothing Then evApp = CType(sender, Autodesk.Revit.UI.UIApplication).Application
        If e.Document IsNot Nothing Then SubscribeToDoc(e.Document)
    End Sub
    Private Shared Sub evAppCUI_ViewActivating(sender As Object, e As ViewActivatingEventArgs) Handles evAppUIC.ViewActivating
        'System.Windows.MessageBox.Show("UIControlledApplication.ViewActivating")
        'SubscribeToDoc(e.Document)
    End Sub
#End Region
    '
#Region "EVENTOS UIAPPLICATION DESACTIVADOS"
    Private Shared Sub evAppCUI_DisplayingOptionsDialog(sender As Object, e As DisplayingOptionsDialogEventArgs) Handles evAppUIC.DisplayingOptionsDialog
        'System.Windows.MessageBox.Show("UIControlledApplication.DisplayingOptionsDialog-->Iniciado")
    End Sub

    Private Shared Sub evAppCUI_DockableFrameFocusChanged(sender As Object, e As DockableFrameFocusChangedEventArgs) Handles evAppUIC.DockableFrameFocusChanged
        'System.Windows.MessageBox.Show("UIControlledApplication.DockableFrameFocusChanged-->Iniciado")
    End Sub

    Private Shared Sub evAppCUI_DockableFrameVisibilityChanged(sender As Object, e As DockableFrameVisibilityChangedEventArgs) Handles evAppUIC.DockableFrameVisibilityChanged
        'System.Windows.MessageBox.Show("UIControlledApplication.DockableFrameVisibilityChanged-->Iniciado")
    End Sub

    Private Shared Sub evAppCUI_FabricationPartBrowserChanged(sender As Object, e As FabricationPartBrowserChangedEventArgs) Handles evAppUIC.FabricationPartBrowserChanged
        'System.Windows.MessageBox.Show("UIControlledApplication.FabricationPartBrowserChanged-->Iniciado")
    End Sub
    Private Shared Sub evAppCUI_TransferredProjectStandards(sender As Object, e As TransferredProjectStandardsEventArgs) Handles evAppUIC.TransferredProjectStandards
        'System.Windows.MessageBox.Show("UIControlledApplication.TransferredProjectStandards-->Iniciado")
    End Sub

    Private Shared Sub evAppCUI_TransferringProjectStandards(sender As Object, e As TransferringProjectStandardsEventArgs) Handles evAppUIC.TransferringProjectStandards
        'System.Windows.MessageBox.Show("UIControlledApplication.TransferringProjectStandards-->Iniciado")
    End Sub
#End Region
End Class
