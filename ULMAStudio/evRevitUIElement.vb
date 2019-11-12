Imports adWin = Autodesk.Windows
Partial Public Class evRevit
    'in OnStartup 
    'AddHandler() Autodesk.Windows.ComponentManager.UIElementActivated, AddressOf MyUiElementActivated

    'in OnShutdown
    'RemoveHandler() Autodesk.Windows.ComponentManager.UIElementActivated, AddressOf MyUiElementActivated

    Public Shared Sub MyUiElementActivated(ByVal sender As Object, ByVal e As Autodesk.Windows.UIElementActivatedEventArgs)
        'do your thing
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
            'e.UiElement.CommandBindings.Add()
        End If
        If mensaje <> "" Then
            'TaskDialog.Show("Datos del botón pulsado (UIElement)", mensaje)
            IO.File.AppendAllText(IO.Path.Combine("H:\DESARROLLO\REVIT_DOC\COMMANDS-EXECUTE-IEXTERNAL", "_comandos.txt"), mensaje & vbCrLf)
        End If
    End Sub
End Class
