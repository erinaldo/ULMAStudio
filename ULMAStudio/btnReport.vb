#Region "Imported Namespaces"
Imports System
Imports System.Collections.Generic
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports uf = ULMALGFree.clsBase
#End Region

<Transaction(TransactionMode.Manual)>
Public Class btnReport
    Implements Autodesk.Revit.UI.IExternalCommand

    Public dFilas As New Dictionary(Of String, filaDatos)

    ''' <summary>
    ''' The one and only method required by the IExternalCommand interface, the main entry point for every external command.
    ''' </summary>
    ''' <param name="commandData">Input argument providing access to the Revit application, its documents and their properties.</param>
    ''' <param name="message">Return argument to display a message to the user in case of error if Result is not Succeeded.</param>
    ''' <param name="elements">Return argument to highlight elements on the graphics screen if Result is not Succeeded.</param>
    ''' <returns>Cancelled, Failed or Succeeded Result code.</returns>
    Public Function Execute(
      ByVal commandData As ExternalCommandData,
      ByRef message As String,
      ByVal elements As ElementSet) _
    As Result Implements IExternalCommand.Execute
        'TODO: Add your code here
        '
        Dim oDoc As Document = commandData.Application.ActiveUIDocument.Document
        enejecucion = True
        Dim resultado As Result = Result.Succeeded
        '
        'If evRevit.evAppUI.ActiveUIDocument.Document.PathName = "" Then
        '    'TaskDialog.Show("ATTENTION", "Save the document first and try again")
        '    MsgBox("Save the document first and try again", MsgBoxStyle.Critical, "ATTENTION")
        '    resultado = Result.Cancelled : GoTo FINAL
        '    Exit Function
        'End If

        If evRevit.evAppUI.ActiveUIDocument.Document.IsFamilyDocument Then
            'TaskDialog.Show("ATTENTION", "Utility only for project document")
            MsgBox("Utility only for project document", MsgBoxStyle.Critical, "ATTENTION")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If
        '
        ' Solo para .rvt
        'If IO.Path.GetExtension(evRevit.evAppUI.ActiveUIDocument.Document.PathName).ToLower.Contains("rvt") = False Then
        '    'TaskDialog.Show("ATTENTION", "Utility only for .rvt files")
        '    MsgBox("Utility only for .rvt files", MsgBoxStyle.Critical, "ATTENTION")
        '    resultado = Result.Cancelled : GoTo FINAL
        '    Exit Function
        'End If

        'If oDoc.IsModified = True Then
        '    TaskDialog.Show("ATTENTION", "First save the document and try again")
        '    resultado = Result.Cancelled : GoTo FINAL
        '    Exit Function
        'End If

        ' Solicitar si solo vista actual o todo el documento.
        ' Solicitar si solo ULMA o todas las familias.
        Dim frmO As New frmReportOptions()

        If frmO.ShowDialog(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle)) = System.Windows.Forms.DialogResult.Cancel Then
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If
        '
        ULMALGFree.clsBase._ultimaApp = ULMALGFree.queApp.ULMASTUDIO
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("REPORT", evRevit.evDoc.PathName)
        ' Show frmGeneratingReport
        If fWait IsNot Nothing Then
            fWait.Close()
            fWait = Nothing
        End If
        fWait = New frmGeneratingReport
        fWait.titulo = "*** Generating Report ***"
        'fWait.TopMost = True
        fWait.Show(New WindowWrapper(System.Diagnostics.Process.GetCurrentProcess.MainWindowHandle))
        fWait.Update()
        'fWait.Show()
        '
        ' Todas las familias del documento o de la vista, solo ulma o todas.
        Dim lFi As List(Of FamilyInstance) = Nothing
        lFi = utilesRevit.FamilyInstance_DameULMA_FAMILYINSTANCE(oDoc,,,, vistaActual:=onlyactiveview, soloulma:=onlyulma)
        If (lFi Is Nothing OrElse lFi.Count = 0) Then   ' And frmO.RbUlma.Checked = True Then
            fWait.Close()
            fWait = Nothing
            If Not oDoc.ActiveView.CanBePrinted Then
                TaskDialog.Show("ATTENTION", "Please activate a view before generating the report.")
            Else
                TaskDialog.Show("ATTENTION", "There are no families to generate the report.")
            End If

            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If
        '
        Dim lFinal As New List(Of FamilyInstance)

        For Each oFi As FamilyInstance In lFi
            ' FILTROS A APLICAR. Según Vista/Proyecto o ULMA/Todos
            ' ** Filtro Supercomponente. Solo añadimos el FamilyInstance Supercomponente (No las sub-familias que contenga)
            If oFi.SuperComponent IsNot Nothing Then Continue For
            ' ** Filtramos los que están configurados como invisibles en todo el proyecto
            If oFi.Invisible Then Continue For
            ' ** Filtrar las invisibles en la vista actual, si onlyactiveview
            If onlyactiveview AndAlso oFi.IsHidden(evRevit.evAppUI.ActiveUIDocument.ActiveView) Then Continue For
            ' ** Filtro onlyulma
            If onlyulma AndAlso FamilyInstance_EsDeULMA(oDoc, oFi) = False Then Continue For
            If lFinal.Contains(oFi) = False Then lFinal.Add(oFi)
            System.Windows.Forms.Application.DoEvents()
        Next
        '
        fWait.pb1.Maximum = lFinal.Count
        'RellenaDatosInforme(oDoc.ActiveView.Name, lFinal)
        Dim nombreVistaAllProject As String
        If oDoc.PathName = "" Then
            nombreVistaAllProject = oDoc.Title
        Else
            nombreVistaAllProject = IO.Path.GetFileName(oDoc.PathName)
        End If

        ''
        Dim units As Units = oDoc.GetUnits()
        Dim fOptions As FormatOptions = units.GetFormatOptions(UnitType.UT_Mass)
        Dim unidadPeso As String = ""
        If (fOptions.DisplayUnits = DisplayUnitType.DUT_KILOGRAMS_MASS) Then
            unidadPeso = "Kg"
        ElseIf (fOptions.DisplayUnits = DisplayUnitType.DUT_POUNDS_MASS) Then
            unidadPeso = "lbm"
        ElseIf (fOptions.DisplayUnits = DisplayUnitType.DUT_TONNES_MASS) Then
            unidadPeso = "t"
        ElseIf (fOptions.DisplayUnits = DisplayUnitType.DUT_USTONNES_MASS) Then
            unidadPeso = "Tons"
        End If

        Dim nombre As String = IIf(onlyactiveview = True, oDoc.ActiveView.Name, nombreVistaAllProject).ToString
        RellenaDatosInforme(nombre, lFinal, unidadPeso)
        '
        enejecucion = False

        If dFilas Is Nothing OrElse dFilas.Count = 0 Then
            fWait.Close()
            fWait = Nothing
            'TaskDialog.Show("ATTENTION", "There are no families to generate the report.")
            MsgBox("There are no families to generate the report.", MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Critical, "ATTENTION")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If
        ' Poner Log
        Dim notas As String = ""
        If onlyactiveview = True Then
            notas = "Active View-->" & nombre
        Else
            notas = "All Project"
        End If
        notas &= " (Rows = " & dFilas.Count & " / FamilyInstances = " & lFinal.Count & ")"
        If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("REPORT", FILENAME:=oDoc.PathName, NOTES:=IO.Path.GetFileName(oDoc.PathName) & " " & notas)
        '
        lFinal = Nothing
        oDoc = Nothing
        lFi = Nothing
        If fWait IsNot Nothing Then
            fWait.Close()
            fWait = Nothing
        End If
        If IO.File.Exists(_reportExe) AndAlso IO.File.Exists(ULMALGFree.clsBase._ULMAStudioReport) Then
            Dim pInf As New ProcessStartInfo(_reportExe)
            pInf.UseShellExecute = False
            Using oProc = Process.Start(pInf)
                oProc.WaitForExit(15000)
                'While p.HasExited = False
                '    System.Windows.Forms.Application.DoEvents()
                'End While
            End Using
            pInf = Nothing
        Else
            MsgBox("ReportGenerator not found or error", MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Critical, "ATTENTION")
        End If
        '
FINAL:
        Return resultado
    End Function

    Private Sub RellenaDatosInforme(VistaActual As String, lFi As List(Of FamilyInstance), unidadPeso As String)
        If lFi Is Nothing OrElse lFi.Count = 0 Then
            Exit Sub
        End If
        ' Cerrar Proceso UCRevitFreeReport (Si esta abierto)
        For Each oProc As Process In Process.GetProcessesByName("UCRevitFreeReport")
            oProc.Kill()
        Next
        '
        Dim resultado As New List(Of filaDatos)
        For Each oFi As FamilyInstance In lFi
            Try
                ' ¿Es una familia de ULMA?
                Dim esulma As Boolean = FamilyInstance_EsDeULMA(evRevit.evDoc, oFi)
                '
                ' ImagePath temporal
                Dim imgPath As String = ""
                Dim name As String = oFi.Name
                imgPath = IO.Path.Combine(path_families_base_images, name & ".png")
                If IO.File.Exists(imgPath) = False Then imgPath = ""
                '
                ' NAME
                If name = "" Then
                    name = oFi.Symbol.Name
                    imgPath = IO.Path.Combine(path_families_base_images, name & ".png")
                    If IO.File.Exists(imgPath) = False Then imgPath = ""
                End If
                If name = "" Then
                    name = oFi.Symbol.FamilyName
                    If imgPath = "" Then
                        imgPath = IO.Path.Combine(path_families_base_images, name & ".png")
                        If IO.File.Exists(imgPath) = False Then imgPath = ""
                    End If
                End If
                '
                ' IMAGE Path (Por si aun no lo hemos cogido)
                If imgPath = "" OrElse IO.File.Exists(imgPath) = False Then
                    If esulma Then
                        imgPath = IO.Path.Combine(_dirApp & "\images", "LogoU.png")
                    Else
                        imgPath = IO.Path.Combine(_dirApp & "\images", "Revit.png")
                    End If
                End If
                ' NAMEINFORME
                Dim nameinforme As String = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi, "ITEM_DESCRIPTION")
                If nameinforme = "" Then
                    nameinforme = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi.Symbol, "ITEM_DESCRIPTION")
                End If
                If nameinforme = "" Then
                    nameinforme = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi.Symbol.Family, "ITEM_DESCRIPTION")
                End If
                If nameinforme = "" And esulma = False Then
                    nameinforme = oFi.Name
                End If
                '
                ' CODE
                Dim code As String = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi, "ITEM_CODE")
                If code = "" Then
                    code = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi.Symbol, "ITEM_CODE")
                End If
                If code = "" Then
                    code = utilesRevit.ParametroElementLeeString(evRevit.evDocUI.Document, oFi.Symbol.Family, "ITEM_CODE")
                End If
                If code <> "" And nameinforme = "" Then
                    If uf.colArticulos.ContainsKey(code) Then
                        '2019/10/24 Xabier Calvo: Si no hay description se coge la descripcion en ingles desde el XML
                        nameinforme = CType(uf.colArticulos(code), ULMALGFree.clsArticulos).colDescritions("en").ToString
                    End If
                End If
                ' No incluir los que code = ""
                If includewithoutcode = False AndAlso code = "" Then
                    Continue For
                End If
                ' WEIGHT
                Dim weightTemp As Object = utilesRevit.ParametroElementLeeObjeto(evRevit.evDocUI.Document, oFi, "ITEM_WEIGHT")  'name.Length
                If weightTemp Is Nothing Then
                    weightTemp = utilesRevit.ParametroElementLeeObjeto(evRevit.evDocUI.Document, oFi.Symbol, "ITEM_WEIGHT")
                End If
                If weightTemp Is Nothing Then
                    weightTemp = utilesRevit.ParametroElementLeeObjeto(evRevit.evDocUI.Document, oFi, "WEIGHT")
                End If
                If weightTemp Is Nothing Then
                    weightTemp = utilesRevit.ParametroElementLeeObjeto(evRevit.evDocUI.Document, oFi.Symbol, "WEIGHT")
                End If
                If weightTemp IsNot Nothing Then
                    weightTemp = uf.CheckDecimal_Value(weightTemp.ToString)
                End If
                Dim weight As Double = 0
                If weightTemp IsNot Nothing Then
                    weight = Convert.ToDouble(IIf(IsNumeric(weightTemp), weightTemp.ToString.Split(" "c)(0), 0))
                End If
                If weight = 0 AndAlso code <> "" Then
                    If uf.colArticulos.ContainsKey(code) Then
                        Dim wTemp As String = CType(uf.colArticulos(code), ULMALGFree.clsArticulos).weight
                        wTemp = uf.CheckDecimal_Value(wTemp.ToString)
                        If wTemp <> "" AndAlso IsNumeric(wTemp) Then
                            wTemp = wTemp.Split(" "c)(0)
                            weight = CDbl(wTemp)
                        End If
                    End If
                End If
                Dim convertedWeight As Double = convertirPeso(weight, unidadPeso)

                ' Clave para ordenar por solo "Name"
                Dim clave As String = IIf(nameinforme = "", name, nameinforme).ToString '& code
                ' Clave para ordenar por "Category" y "Name" (Lo quitamos)
                'Dim clave As String = oFi.Category.Name & IIf(nameinforme = "", name, nameinforme).ToString '& code
                If dFilas.ContainsKey(clave) = False Then
                    dFilas.Add(clave, New filaDatos(imgPath, IIf(nameinforme = "", "", nameinforme).ToString, code, convertedWeight, 1, esulma, oFi.Category.Name))
                Else
                    If dFilas(clave).Code = "" AndAlso code <> "" Then dFilas(clave).Code = code
                    dFilas(clave).Quantity += 1
                    ' Cambiamos para que sólo ponga el peso unitario (No los sume)
                    'dFilas(clave).Weight += weight
                    If dFilas(clave).Weight = 0 And convertedWeight > 0 Then
                        dFilas(clave).Weight = convertedWeight
                    End If
                End If
                System.Windows.Forms.Application.DoEvents()
            Catch ex As Exception
                Continue For
            End Try
        Next
        '
        ' Rellenar UCRevitFreeReport.txt
        Dim datos As String = VistaActual & vbCrLf  ' Primera linea es el nombre de la vista.
        datos &= unidadPeso & vbCrLf 'Segunda linea es la unidad de Peso del proyecto
        ' Ordenar colección de filaDatos por "Name"        
        Dim filas = From x In dFilas.Values
                    Order By x.Name
                    Select x
        '' Ordenar colección de filaDatos por "Category" y "Name" (Lo quitamos)
        'Dim filas = From x In dFilas.Values
        '            Order By x.categoria, x.Name
        '            Select x
        ' 1.- Primero los que si son de ULMA
        For Each oFD As filaDatos In filas
            ' Si no es ULMA, continuar
            If oFD.esulma = False Then Continue For
            fWait.pb1_Pon()
            datos &= oFD.ImagePath & ";" & oFD.Name & ";" & oFD.Code & ";" & oFD.Weight & ";" & oFD.Quantity & ";" & oFD.esulma.ToString & vbCrLf
            System.Windows.Forms.Application.DoEvents()
        Next
        ' 2.- Segundo los que no son de ULMA
        For Each oFD As filaDatos In filas
            ' Si es ULMA, continuar
            If oFD.esulma = True Then Continue For
            fWait.pb1_Pon()
            datos &= oFD.ImagePath & ";" & oFD.Name & ";" & oFD.Code & ";" & oFD.Weight & ";" & oFD.Quantity & ";" & oFD.esulma.ToString & vbCrLf
            System.Windows.Forms.Application.DoEvents()
        Next
        Try
            If IO.File.Exists(ULMALGFree.clsBase._ULMAStudioReport) Then
                IO.File.Delete(ULMALGFree.clsBase._ULMAStudioReport)
            End If

            If datos <> "" AndAlso filas.Count > 0 Then
                datos = datos.Substring(0, datos.Length - 2)
                IO.File.WriteAllText(ULMALGFree.clsBase._ULMAStudioReport, datos, System.Text.Encoding.UTF8)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Function convertirPeso(ByVal valor As Double, ByVal unidad As String) As Double
        Dim resultado As Double = valor
        Select Case unidad
            Case "lbm"
                resultado = Math.Round(valor * 2.204623, 3)
            Case "t"
                resultado = Math.Round(valor / 1000, 3)
            Case "Tons"
                resultado = Math.Round(valor / 1000, 3)
            Case Else
        End Select
        Return resultado
    End Function
End Class
'End Namespace
