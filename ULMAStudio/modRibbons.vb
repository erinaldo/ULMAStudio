#Region "Imported Namespaces"
Imports System
Imports System.Collections.Generic
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
'Imports UR2ACAD2016.btnWeb2aCAD
Imports adWin = Autodesk.Windows
#End Region
Module modRibbons
    Public fiRevitRibbons As String = My.Computer.FileSystem.SpecialDirectories.Temp ' IO.Path.GetTempPath
    '
    Public Function RibbonTab_ActivarVisualizar(idRibbon As String) As Boolean
        Dim nombreigual As Boolean = True
REPITE:
        Dim resultado As Boolean = False
        Dim r As Autodesk.Windows.RibbonControl = Autodesk.Windows.ComponentManager.Ribbon
        For Each t As Autodesk.Windows.RibbonTab In r.Tabs
            If nombreigual Then
                If t.Id <> idRibbon Then Continue For
            Else
                If t.Id.Contains(idRibbon) = False Then Continue For
            End If
            Try
                If t.IsVisible = False Then t.IsActive = True
                If t.IsEnabled = False Then t.IsEnabled = True
                If t.IsActive = False Then t.IsActive = True
                't.RibbonControl.Focus()
                r.ActiveTab = t
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
            Exit For
        Next
FINAL:
        If resultado = False And nombreigual = True Then
            nombreigual = False
            GoTo REPITE
        End If
        '
        Return resultado
    End Function
    Public Sub Ribbons_ListaAFichero(Optional fiOut As String = "")
        If fiOut = "" Then fiOut = fiRevitRibbons
        Dim mensaje As String = ""
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        ''
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            'If tab.AutomationName = _uiLabName Or tab.Name = _uiLabName Then Continue For
            '' RIBBONTABS
            mensaje &= NombreSinSaltos(tab.AutomationName) & " / " & tab.Id & " (RibbonTab)" & vbCrLf
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                mensaje &= vbTab & If(oPanel.Source.Name = "", NombreSinSaltos(oPanel.Source.AutomationName), NombreSinSaltos(oPanel.Source.Name)) & " / " & oPanel.Source.Id & " (RibbonPanel)" & vbCrLf
                ''
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    ''
                    If TypeOf oRi Is adWin.RibbonSplitButton Then
                        Dim oBtn As adWin.RibbonSplitButton = CType(oRi, adWin.RibbonSplitButton)
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonSplitButton / " & oRi.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonSplitButton
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonSplitButton).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem / " & oRi1.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ") " & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonRowPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonRowPanel / " & oRi.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonRowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonRowPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem / " & oRi1.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFlowPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonFlowPanel / " & oRi.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFlowPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem / " & oRi1.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    ElseIf TypeOf oRi Is adWin.RibbonFoldPanel Then
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonFoldPanel / " & oRi.Id & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        '' RibbonItem de cada RibbonFlowPanel
                        For Each oRi1 As adWin.RibbonItem In CType(oRi, adWin.RibbonFoldPanel).Items
                            If oRi1.Text <> "" Then mensaje &= vbTab & vbTab & vbTab & NombreSinSaltos(oRi1.Text) & " (RibbonItem / " & oRi1.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                        Next
                    Else
                        '' RibbonButton
                        If oRi.Text <> "" Then mensaje &= vbTab & vbTab & NombreSinSaltos(oRi.Text) & " (RibbonItem / " & oRi.Id & " / " & oRi.Name & " / " & oRi.AutomationName & ")" & vbCrLf ' & "(" & oRi1.Cookie & ")" & vbCrLf
                    End If
                Next
            Next
        Next
        'MsgBox(mensaje)
        Dim ficherofin As String = "C:\DESARROLLO\RevitRibbons.txt"
        IO.File.WriteAllText(ficherofin, mensaje)
        Process.Start(ficherofin)
    End Sub
    '
    Public Function RibbonItem_BuscaDame(idRibbonBusco As String,
                              idRibbonPanelBusco As String,
                              idRibbonItemBusco As String) As adWin.RibbonItem
        Dim encontrado As Boolean = False
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        On Error Resume Next
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            'On Error Resume Next
            If idRibbonBusco <> "" AndAlso tab.Id <> idRibbonBusco Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                'On Error Resume Next
                If idRibbonPanelBusco <> "" AndAlso oPanel.Source.Id <> idRibbonPanelBusco Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    '' Si no es un RibbonSplitButton, continuar
                    'On Error Resume Next
                    If oRi.Id = idRibbonItemBusco Then
                        'ritem = oRi.Clone
                        'ritem.LargeImage = oRi.LargeImage
                        'ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        'If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        oRi.IsEnabled = True : oRi.IsVisible = True
                        ritem = oRi
                        encontrado = True
                        Exit For
                    End If
                Next
                If encontrado = True Then Exit For
            Next
            If encontrado = True Then Exit For
        Next
        '
        Return ritem
    End Function
    '
    Public Sub RibbonButton_ChangeImage(oBtn As PushButton, fullPathImage As String)
        If oBtn Is Nothing OrElse IO.File.Exists(fullPathImage) = False Then Exit Sub
        Dim RibbonControl As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim RibbonTab As Autodesk.Windows.RibbonTab = RibbonControl.FindTab(_uiLabName)
        If RibbonTab Is Nothing Then RibbonTab = RibbonControl.FindTab(_uiLabNameULMA)
        Dim ASPanel As Autodesk.Windows.RibbonPanel
        Dim rButton As Autodesk.Windows.RibbonButton

        For Each panel In RibbonTab.Panels
            'If panel.Source.Id = "Button" Then
            ASPanel = panel
            For Each item In ASPanel.Source.Items
                If item.Id.Contains(oBtn.Name) Then      ' Button
                    rButton = CType(item, Autodesk.Windows.RibbonButton)
                    rButton.LargeImage = New System.Windows.Media.Imaging.BitmapImage(New Uri(fullPathImage, UriKind.Absolute))
                    rButton.Image = rButton.LargeImage
                    rButton.ShowImage = False
                    rButton.ShowImage = True
                    oBtn.Image = rButton.Image
                    oBtn.LargeImage = rButton.LargeImage
                End If
            Next
            'End If
        Next
    End Sub
    Public Sub RibbonButton_ChangeImage(oBtn As PushButton, resourceImage As System.Windows.Media.Imaging.BitmapImage)
        If oBtn Is Nothing OrElse resourceImage Is Nothing Then Exit Sub
        Dim RibbonControl As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim RibbonTab As Autodesk.Windows.RibbonTab = RibbonControl.FindTab(_uiLabName)
        If RibbonTab Is Nothing Then RibbonTab = RibbonControl.FindTab(_uiLabNameULMA)
        Dim ASPanel As Autodesk.Windows.RibbonPanel
        Dim rButton As Autodesk.Windows.RibbonButton

        For Each panel In RibbonTab.Panels
            'If panel.Source.Id = "Button" Then
            ASPanel = panel
            For Each item In ASPanel.Source.Items
                If item.Id.Contains(oBtn.Name) Then      ' Button
                    rButton = CType(item, Autodesk.Windows.RibbonButton)
                    rButton.Image = resourceImage  'System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                    rButton.LargeImage = resourceImage
                End If
            Next
            'End If
        Next
    End Sub
    Public Sub RibbonItem_PonEnMiPanel(ByRef RibbonPanelDestino As adWin.RibbonPanel,
                              idRibbonBusco As String,
                              idRibbonPanelBusco As String,
                              idRibbonItemBusco As String)
        Dim encontrado As Boolean = False
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        Dim ritem As adWin.RibbonItem = Nothing
        ''
        On Error Resume Next
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON queRibbon, continuar
            'On Error Resume Next
            If idRibbonBusco <> "" AndAlso tab.Id <> idRibbonBusco Then Continue For
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                '' Si no es el RIBBONPANEL, continuar
                'On Error Resume Next
                If idRibbonPanelBusco <> "" AndAlso oPanel.Source.Id <> idRibbonPanelBusco Then Continue For
                '' RibbonItem de cada RibbonPanel
                For Each oRi As adWin.RibbonItem In oPanel.Source.Items
                    '' Si no es un RibbonSplitButton, continuar
                    'On Error Resume Next
                    If oRi.Id = idRibbonItemBusco Then
                        'ritem = oRi.Clone
                        'ritem.LargeImage = oRi.LargeImage
                        'ritem.Size = Autodesk.Windows.RibbonItemSize.Large
                        'If contexto = False Then ritem.ResizeStyle = Autodesk.Windows.RibbonItemResizeStyles.HideText
                        oRi.IsEnabled = True : oRi.IsVisible = True
                        RibbonPanelDestino.Source.Items.Add(oRi)
                        encontrado = True
                        Exit For
                    End If
                Next
                If encontrado = True Then Exit For
            Next
            If encontrado = True Then Exit For
        Next
    End Sub
#Region "UTILIDADES"
    Public Function NombreSinSaltos(queNombre As String) As String
        If queNombre Is Nothing Then queNombre = ""
        If queNombre <> "" Then
            queNombre = queNombre.Replace(vbCrLf, "[vbCrLf]")
            queNombre = queNombre.Replace(vbCr, "[vbCr]")
            queNombre = queNombre.Replace(vbLf, "[vbLf]")
        End If
        Return queNombre
    End Function
#End Region
End Module
