Imports adWin = Autodesk.Windows
Imports Autodesk.Internal.Windows
Imports Autodesk.Windows
Imports System.ComponentModel

Partial Public Class evRevit
#Region "EVENTOS COMPONENTMANAGER"
    Public Shared SaveAsLibrary As Boolean = False
    Public Shared newQue As String = ""
    '
    Private Shared Sub evCompM_UIElementActivated(sender As Object, e As adWin.UIElementActivatedEventArgs) Handles evCompM.UIElementActivated
        'System.Windows.MessageBox.Show("ComponentManager.UIElementActivated")
        If Environment.MachineName.ToUpper <> "ALBERTO-HP" Then
            Exit Sub
        End If
        Dim mensaje As String = ""
        If TypeOf e.Item Is adWin.RibbonItem Then
            On Error Resume Next
            mensaje &= "AutomationName = " & e.Item.AutomationName & vbCrLf
            mensaje &= "Name = " & e.Item.Name & vbCrLf
            mensaje &= "Text = " & e.Item.Text & vbCrLf
            mensaje &= "Description = " & e.Item.Description & vbCrLf
            mensaje &= "Id = " & e.Item.Id & vbCrLf
            mensaje &= "Cookie=" & e.Item.Cookie & vbCrLf
            mensaje &= "UId = " & e.Item.UID & vbCrLf
            mensaje &= "Description = " & e.Item.Description & vbCrLf
            mensaje &= "Image = " & e.Item.Image.ToString & vbCrLf
            mensaje &= StrDup(40, "*") & vbCrLf
            'e.UiElement.CommandBindings.Add()
        End If
        If mensaje <> "" Then
            'TaskDialog.Show("Datos del botón pulsado (UIElement)", mensaje)
            IO.File.AppendAllText(IO.Path.Combine("H:\DESARROLLO\REVIT_DOC\COMMANDS-EXECUTE-IEXTERNAL", "_comandos.txt"), mensaje & vbCrLf)
        End If
    End Sub
#End Region
    '
#Region "EVENTOS COMPONENTMANAGER DESACTIVADOS"
    Private Shared Sub evCompM_ItemAddedToQuickAccessToolBar(sender As Object, e As RibbonItemAddedToQuickAccessToolBarEventArgs) Handles evCompM.ItemAddedToQuickAccessToolBar
        'System.Windows.MessageBox.Show("ComponentManager.ItemAddedToQuickAccessToolBar")
    End Sub
    Private Shared Sub evCompM_ItemExecuted(sender As Object, e As RibbonItemExecutedEventArgs) Handles evCompM.ItemExecuted
        'System.Windows.MessageBox.Show("ComponentManager.ItemExecuted")
        evRevit.SaveAsLibrary = False
        evRevit.newQue = ""
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
    End Sub
    Private Shared Sub evCompM_ItemInitialized(sender As Object, e As RibbonItemEventArgs) Handles evCompM.ItemInitialized
        'System.Windows.MessageBox.Show("ComponentManager.ItemInitialized")
    End Sub
    Public Shared Sub evCompM_PreviewExecute(sender As Object, e As EventArgs) Handles evCompM.PreviewExecute
        'System.Windows.MessageBox.Show("ComponentManager.PreviewExecute")
        'Try
        '    Dim oItem As Autodesk.Windows.ApplicationMenuItem = CType(sender, Autodesk.Windows.ApplicationMenuItem)
        '    If oItem.Id = "ID_SAVE_FAMILY_ANY" Then
        '        evRevit.SaveAsLibrary = True
        '    ElseIf oItem.Id = "ID_FILE_NEW_CHOOSE_TEMPLATE" Then    ' New-Project
        '        evRevit.newQue = "NEW_PROJECT"
        '    ElseIf oItem.Id = "ID_FAMILY_NEW" Then                  ' New-Family
        '        evRevit.newQue = "NEW_PROJECT"
        '    ElseIf oItem.Id = "ID_NEW_REVIT_DESIGN_MODEL" Then      ' New-Conceptual Mass
        '        evRevit.newQue = "NEW_CONCEPTUAL_MASS"
        '    ElseIf oItem.Id = "ID_TITLEBLOCK_NEW" Then              ' New-Title Block
        '        evRevit.newQue = "NEW_TITLE_BLOCK"
        '    ElseIf oItem.Id = "ID_ANNOTATION_SYMBOL_NEW" Then       ' Nes-Annotation Symbol
        '        evRevit.newQue = "NEW_ANNOTATION_SYMBOL"
        '    End If
        'Catch ex As Exception
        'End Try
    End Sub
    Private Shared Sub evCompM_PropertyChanged(sender As Object, e As PropertyChangedEventArgs) Handles evCompM.PropertyChanged
        'System.Windows.MessageBox.Show("ComponentManager.PropertyChanged")
    End Sub
    Private Shared Sub evCompM_ToolTipClosed(sender As Object, e As EventArgs) Handles evCompM.ToolTipClosed
        'System.Windows.MessageBox.Show("ComponentManager.ToolTipClosed")
    End Sub
    Private Shared Sub evCompM_ToolTipOpened(sender As Object, e As EventArgs) Handles evCompM.ToolTipOpened
        'System.Windows.MessageBox.Show("ComponentManager.ToolTipOpened")
    End Sub
#End Region
End Class
