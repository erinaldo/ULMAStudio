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
        If evRevit.evAppUI.ActiveUIDocument.Document.PathName = "" Then
            'TaskDialog.Show("ATTENTION", "Save the document first and try again")
            MsgBox("Save the document first and try again", MsgBoxStyle.Critical, "ATTENTION")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If

        If evRevit.evAppUI.ActiveUIDocument.Document.IsFamilyDocument Then
            'TaskDialog.Show("ATTENTION", "Utility only for project document")
            MsgBox("Utility only for project document", MsgBoxStyle.Critical, "ATTENTION")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If

        ''
        '' Solo para .rvt
        If IO.Path.GetExtension(evRevit.evAppUI.ActiveUIDocument.Document.PathName).ToLower.Contains("rvt") = False Then
            'TaskDialog.Show("ATTENTION", "Utility only for .rvt files")
            MsgBox("Utility only for .rvt files", MsgBoxStyle.Critical, "ATTENTION")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If

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
            TaskDialog.Show("ATTENTION", "There are no families to generate the report.")
            resultado = Result.Cancelled : GoTo FINAL
            Exit Function
        End If
        '
        Dim lFinal As New List(Of FamilyInstance)

        For Each oFi As FamilyInstance In lFi
            'fWait.pb1_Pon()
            ' Filtramos los que están configurados como invisibles en todo el proyecto
            If oFi.Invisible Then Continue For
            '
            ' Filtrar las invisibles en la vista actual, si onlyactiveview = True
            If onlyactiveview Then
                ' ALBERTO. Quieren que si salgan los ocultos en la vista actual.
                If oFi.IsHidden(evRevit.evAppUI.ActiveUIDocument.ActiveView) Then Continue For
            End If

            ' Si tiene Supercomponente, continuar. Solo añadimos el FamilyInstance Supercomponente.
            If FamilyInstance_EsDeULMA(oDoc, oFi) AndAlso oFi.SuperComponent IsNot Nothing Then Continue For
            If lFinal.Contains(oFi) = False Then lFinal.Add(oFi)
            System.Windows.Forms.Application.DoEvents()
        Next
        '
        fWait.pb1.Maximum = lFinal.Count
        'RellenaDatosInforme(oDoc.ActiveView.Name, lFinal)
        Dim nombre As String = IIf(onlyactiveview = True, oDoc.ActiveView.Name, IO.Path.GetFileName(oDoc.PathName)).ToString
        RellenaDatosInforme(nombre, lFinal)
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
        If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("REPORT", FILENAME:=oDoc.PathName, NOTES:=notas)
        '
        lFinal = Nothing
        oDoc = Nothing
        lFi = Nothing
        If fWait IsNot Nothing Then
            fWait.Close()
            fWait = Nothing
        End If
        If IO.File.Exists(_reportExe) Then
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
            MsgBox("ReportGenerator not found", MsgBoxStyle.ApplicationModal Or MsgBoxStyle.Critical, "ATTENTION")
        End If
        '
FINAL:
        Return resultado
    End Function

    Private Sub RellenaDatosInforme(VistaActual As String, lFi As List(Of FamilyInstance))
        If lFi Is Nothing OrElse lFi.Count = 0 Then
            Exit Sub
        End If
        ' Cerrar Proceso UCRevitFreeReport (Si esta abierto)
        For Each oProc As Process In Process.GetProcessesByName("UCRevitFreeReport")
            oProc.Kill()
        Next
        '
        Dim resultado As New List(Of filaDatos)
        'Dim carpeta As String = IO.File.ReadAllText(_appOptions)
        For Each oFi As FamilyInstance In lFi
            '
            'fWait.pb1_Pon()
            If oFi.Invisible Then Continue For
            If oFi.IsHidden(evRevit.evAppUI.ActiveUIDocument.ActiveView) Then Continue For
            '
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
            'Dim quantity As Integer = 1
            '
            'If nameinforme <> "" Then name = nameinforme
            Dim clave As String = IIf(nameinforme = "", name, nameinforme).ToString '& code
            If dFilas.ContainsKey(clave) = False Then
                dFilas.Add(clave, New filaDatos(imgPath, IIf(nameinforme = "", "", nameinforme).ToString, code, weight, 1, esulma))
            Else
                If dFilas(clave).Code = "" AndAlso code <> "" Then dFilas(clave).Code = code
                dFilas(clave).Quantity += 1
                ' Cambiamos para que sólo ponga el peso unitario (No los sume)
                'dFilas(clave).Weight += weight
                If dFilas(clave).Weight = 0 And weight > 0 Then
                    dFilas(clave).Weight = weight
                End If
            End If
            System.Windows.Forms.Application.DoEvents()
        Next
        '
        ' Rellenar UCRevitFreeReport.txt
        Dim datos As String = VistaActual & vbCrLf  ' Primera linea es el nombre de la vista.
        ' Ordenar colección de filaDatos por "Name"
        Dim filas = From x In dFilas.Values
                    Select x
                    Order By x.Name
        ' 1.- Primero los que si son de ULMA
        For Each oFD As filaDatos In filas
            ' Si no es ULMA, continuar
            If oFD.esulma = False Then
                Continue For
            End If
            fWait.pb1_Pon()
            datos &= oFD.ImagePath & ";" & oFD.Name & ";" & oFD.Code & ";" & oFD.Weight & ";" & oFD.Quantity & ";" & oFD.esulma.ToString & vbCrLf
            System.Windows.Forms.Application.DoEvents()
        Next
        ' 2.- Segundo los que no son de ULMA
        For Each oFD As filaDatos In filas
            ' Si es ULMA, continuar
            If oFD.esulma = True Then
                Continue For
            End If
            fWait.pb1_Pon()
            datos &= oFD.ImagePath & ";" & oFD.Name & ";" & oFD.Code & ";" & oFD.Weight & ";" & oFD.Quantity & ";" & oFD.esulma.ToString & vbCrLf
            System.Windows.Forms.Application.DoEvents()
        Next
        datos = datos.Substring(0, datos.Length - 2)
        IO.File.WriteAllText(ULMALGFree.clsBase._ULMAStudioReport, datos, System.Text.Encoding.UTF8)
    End Sub
End Class
'End Namespace
