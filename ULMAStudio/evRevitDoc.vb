Imports System
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
Imports Autodesk.Revit.DB.Analysis
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.UI.Events
'Imports Autodesk.Revit.Collections
Imports Autodesk.Revit.Exceptions
'Imports Autodesk.Revit.Utility
Imports Autodesk.Revit.ApplicationServices.Application
Imports Autodesk.Revit.DB.Document

Partial Public Class evRevit
#Region "EVENTOS DOCUMENTO DESACTIVADOS"
    Public Shared Sub evDoc_DocumentClosing(ByVal sender As Object, ByVal e As DocumentClosingEventArgs)
        'System.Windows.MessageBox.Show("evDoc.DocumentClosing")
        'AddHandler evRevit.evApp.DocumentClosed, AddressOf evRevit.evApp_DocumentClosed
        'If evRevit.evClose = False Then Exit Sub
        'clsCR.crUltimodoccerrado = e.Document.PathName
        ''RemoveHandler evApp.DocumentClosing, AddressOf ApplicationEvent_DocumentClosing_Handler
        'evRevit.evClose = False
        'Dim lLinks As New List(Of String)
        'LinksDocumento_RellenaList(e.Document, lLinks)
        'For Each oLink As String In lLinks
        '    clsCR.Fichero_Quita(oLink, borrar:=True)
        'Next
    End Sub

    Public Shared Sub evDoc_DocumentPrinted(ByVal sender As Object, ByVal e As DocumentPrintedEventArgs)
        'System.Windows.MessageBox.Show("evDoc.DocumentPrinted")
    End Sub

    Public Shared Sub evDoc_DocumentPrinting(ByVal sender As Object, ByVal e As DocumentPrintingEventArgs)
        'System.Windows.MessageBox.Show("evDoc.DocumentPrinting")
    End Sub

    Public Shared Sub evDoc_DocumentSaved(ByVal sender As Object, ByVal e As DocumentSavedEventArgs)
        'Exit Sub
        'System.Windows.MessageBox.Show("evDoc.DocumentSaved")
        ' Solo encriptar si es documento de ULMA y si crEncript = True
        'If evRevit.evSave = False Then Exit Sub
        'If clsCR.crUltimoguardado = "" OrElse clsCR.crEncrypt = False Then
        '    ' Lo guardamos y no hacemos nada mas. Esta en coleccion de fichero que no se encriptan o No activada encriptación.
        '    Exit Sub
        'End If
        ''
        'evRevit.evClose = False
        'If clsCR.crDicFicheros.ContainsKey(clsCR.crUltimoguardado) Then
        '    'PonLog("GUARDADO " & clsCR.crDicFicheros(clsCR.crUltimoguardado))
        '    ' Nueva encriptación en tmplog
        '    clsCR.Fichero_EncriptaConBackupsTemp(clsCR.crUltimoguardado, clsCR.crDicFicheros(clsCR.crUltimoguardado), clsCR.crSaveImagePreview, False)
        '    clsCR.DocumentoActual_EncriptaLinks()
        'Else
        '    clsCR.crfiAbrir = clsCR.Fichero_PonNuevo(clsCR.crUltimoguardado, True)
        '    clsComandos.Fichero_DesencriptaAbreTemp(clsCR.crfiAbrir)
        '    Call Documento_EstaAbierto(evApp, clsCR.crDicFicheros(clsCR.crUltimoguardado), True)
        'End If
        ''
        ''If clsCR.crSaveImagePreview Then
        ''    ImagenDocumentoGuarda(evRevit.evAppUI.ActiveUIDocument.Document, IO.Path.GetDirectoryName(evRevit.evAppUI.ActiveUIDocument.Document.PathName))
        ''End If
        'evRevit.evClose = True
    End Sub

    Public Shared Sub evDoc_DocumentSavedAs(ByVal sender As Object, ByVal e As DocumentSavedAsEventArgs)
        'System.Windows.MessageBox.Show("evDoc.DocumentSavedAs")
    End Sub

    Public Shared Sub evDoc_DocumentSaving(ByVal sender As Object, ByVal e As DocumentSavingEventArgs)
        'Exit Sub
        ''System.Windows.MessageBox.Show("evDoc.DocumentSaving")
        'If evRevit.evSave = False Then Exit Sub
        'clsCR.Fichero_ReadOnly(e.Document.PathName, sololectura:=False)
        'clsCR.crUltimoguardado = e.Document.PathName
        '' Codificamos todos los ficheros (No solo los de ULMA) Comentamos esto
        ''If Documento_EsDeULMA(e.Document) = False Then  ' OrElse e.Document.PathName.StartsWith(clsCR.crfolder_Temp) = False Then
        ''    '' No lleva familias ulma. No lo codificaremos.
        ''    clsCR.crUltimoguardado = ""
        ''    Exit Sub
        ''End If
        'If clsCR.crDicFicherosNO.Contains(e.Document.PathName) = True Then
        '    clsCR.crUltimoguardado = ""
        'End If
    End Sub

    Public Shared Sub evDoc_DocumentSavingAs(ByVal sender As Object, ByVal e As DocumentSavingAsEventArgs)
        'System.Windows.MessageBox.Show("evDoc.DocumentSavingAs")
    End Sub

    Public Shared Sub evDoc_ViewPrinted(ByVal sender As Object, ByVal e As ViewPrintedEventArgs)
        'System.Windows.MessageBox.Show("evDoc.ViewPrinted")
    End Sub

    Public Shared Sub evDoc_ViewPrinting(ByVal sender As Object, ByVal e As ViewPrintingEventArgs)
        'System.Windows.MessageBox.Show("evDoc.ViewPrinting")
    End Sub
#End Region
End Class