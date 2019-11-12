#Region "Imported Namespaces"
Imports System
Imports System.Collections.Generic
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
#End Region

<Transaction(TransactionMode.Manual)>
Public Class btnAbout
    Implements Autodesk.Revit.UI.IExternalCommand

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
        'PonLog("btnAbout (About)")
        ''
        enejecucion = True
        Dim resultado As Result = Result.Succeeded
        '
        ' Poner la imagen de actualización
        'Dim queImg As System.Windows.Media.Imaging.BitmapSource = Nothing
        'Dim nombre As String = "LogoUupdate32.png"
        'queImg = Imagen_DesdeResources(nombre)
        '' Add an icon 
        'btnAboutBoton.Image = queImg
        'btnAboutBoton.LargeImage = queImg
        'btnAboutBoton.ToolTipImage = queImg
        ''
        'System.Windows.Forms.Application.DoEvents()
        'Threading.Thread.Sleep(5000)
        ''
        '' Volver a poner la imagen que había
        'nombre = "LogoU32.png"
        'queImg = Imagen_DesdeResources(nombre)
        '' Add an icon 
        'btnAboutBoton.Image = queImg
        'btnAboutBoton.LargeImage = queImg
        'btnAboutBoton.ToolTipImage = queImg
        'Return resultado
        'Exit Function
        '*********************************************

        frmA = New frmAbout
        'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UCREVIT_ABOUT,,, arrM123, arrL123, ultimaTraduccion)
        ''
        ''
        If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.ABOUT)
        If frmA.ShowDialog(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle)) = System.Windows.Forms.DialogResult.OK Then
            resultado = Result.Succeeded
        Else
            resultado = Result.Cancelled
        End If
        ''
        '' Vaciamos las variables
        'evRevit.evAppUI = Nothing
        'utApp1 = Nothing
        frmA = Nothing
        ''
        enejecucion = False
        Return resultado
    End Function
End Class
'End Namespace
