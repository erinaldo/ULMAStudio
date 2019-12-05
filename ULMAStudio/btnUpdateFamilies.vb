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
Public Class btnUpdateFamilies
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
        Dim resultado As Result
        Dim pruebaCon As String
        pruebaCon = uf.EstadoRed_String
        If pruebaCon = "" Then
            enejecucion = True
            resultado = Result.Succeeded
            '
            'If uf.cUp("families").Count = 0 Then
            '    TaskDialog.Show("ATTENTION", "Not updates of ULMA families")
            '    Return Result.Failed
            '    Exit Function
            'End If
            '
            uf.frmUFam = New ULMALGFree.frmUpdater
            'If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA(ULMALGFree.ACTION.UCREVIT_ABOUT,,, arrM123, arrL123, ultimaTraduccion)
            ''
            ''
            If cLcsv IsNot Nothing Then cLcsv.PonLog_ULMA("DOWNLOAD_OPEN")
            If uf.frmUFam.ShowDialog(New WindowWrapper(Process.GetCurrentProcess.MainWindowHandle)) = System.Windows.Forms.DialogResult.Cancel Then
                'resultado = Result.Cancelled
                'Else
                ULMAStudioApplication.Botones_ActualizaEstadoActualizaciones()
                ULMAStudioApplication.BotonBrowserReport()
                resultado = Result.Succeeded
            End If
            enejecucion = False
            ''
        Else
            TaskDialog.Show("No network Connection", "Please check your network connection")
            resultado = Result.Cancelled
        End If
        Return resultado
    End Function
End Class
'End Namespace
