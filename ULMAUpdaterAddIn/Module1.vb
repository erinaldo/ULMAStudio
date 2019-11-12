Imports uf = ULMALGFree.clsBase
Module Module1
    Public frmU As frmUpdateAddin
    Public _appFull As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    Public _appFolder As String = IO.Path.GetDirectoryName(_appFull)
    'Public _appFolderUpdates As String = IO.Path.Combine(_appFolder, "Updates\")
    'Public _iniFull As String = IO.Path.ChangeExtension(_appFull, ".ini")
    '
    Private queAccion As String = action.CHECK.ToString
    '
    Public cIni As clsINI
    Public cLg As ULMALGFree.clsBase
    Public releerFTP As Boolean = False

    Sub Main()
        'Threading.Thread.Sleep(3000)
        If cLg Is Nothing Then cLg = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)
        If cIni Is Nothing Then cIni = New clsINI()
        '
        'Console.CursorVisible = False
        Dim parametros() As String = Environment.GetCommandLineArgs
        ' Ejemplo: ULMAUpdaterAddin.exe REVIT True .\Updates\UCRevitFree_20191005.zip %appdata%\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree\
        ' 0 = Camino completo del .exe
        ' 1 = KILLUPDATE / UPDATE / CHECK
        ' 2 = REVIT     (Es opcional. Será el proceso que tenemos que parar o esperar a que pare)
        ' 3 = TRUE o FALSE (O NADA) Para que consulte el FTP y escribar ULMAUpdaterAddIn.in
        ' ** Leeremos ULMAUpdaterAddIn.ini (Ahi estará la configuración de actualización)
        '
        'For Each queP As String In parametros
        '    Console.WriteLine(queP)
        'Next
        'Console.ReadLine()
        '
        Dim fiBorra As String = IO.Path.Combine(IO.Path.GetDirectoryName(parametros(0)), "ejecuta.bat")
        Try
            IO.File.Delete(fiBorra)
        Catch ex As Exception
        End Try
        ' Comprobar y coger los parámetros de la linea de comandos.
        If parametros.Count = 1 Then End
        If parametros.Count > 1 Then
            queAccion = parametros(1)
            '    Try
            '        Call [Enum].TryParse(Of action)(parametros(1), queAccion)
            '    Catch ex As Exception
            '        'End
            '    End Try
        End If
        Dim nombreproceso As String = ULMALGFree.clsBase.nProcess
        If parametros.Count > 2 Then
            nombreproceso = parametros(2)
        End If
        ' Releer INIs y FTP
        'If parametros.Count > 3 Then
        'ULMALGFree.clsBase.cUp = ULMALGFree.clsBase.FTP_DameActualizacionesTodas()
        Dim resultado As String() = INICargar()
        Dim cmd As String = uf.Process_DameFullPath ' ULMALGFree.clsBase.RevitFull
        'MsgBox(cmd)
        'End If
        '
        If IO.Directory.Exists(ULMALGFree.clsBase._updatesFolder) = False Then
            Try
                IO.Directory.CreateDirectory(ULMALGFree.clsBase._updatesFolder)
            Catch ex As Exception
                'MsgBox(ex.ToString, MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
                End
            End Try
        End If
        '
        ' Leer .ini y mostrar errores/salir o continuar si mensaje=""
        Dim mensaje As String = ULMALGFree.clsBase.INIUpdates_LeeTODO
        If mensaje <> "" Then
            'MsgBox(mensaje.Replace("·", vbCrLf), MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
            End
        End If
        ' CHECK o UPDATE
        If action.KILLUPDATE.ToString = queAccion OrElse action.UPDATE.ToString = queAccion Then
            ' KILLUPDATE. Terminar proceso Revit
            Dim p As Process() = Nothing
            Do
                p = System.Diagnostics.Process.GetProcessesByName(nombreproceso)
                If action.KILLUPDATE.ToString = queAccion Then  ' AndAlso p IsNot Nothing AndAlso p.Count > 0 Then
                    ULMALGFree.clsBase.Process_Close(nombreproceso)
                End If
                System.Windows.Forms.Application.DoEvents()
            Loop While p IsNot Nothing AndAlso p.Count > 0
            Threading.Thread.Sleep(2000)
            '
            Dim oD As ULMALGFree.Datos = Nothing
            If ULMALGFree.clsBase.LUp.Count > 0 Then
                For Each d As ULMALGFree.Datos In ULMALGFree.clsBase.LUp
                    ' Solo actualizaremos UCRevitFree
                    If d.ClaveIni.ToUpper <> ULMALGFree.clsBase._AddInName.ToUpper Then
                        Continue For
                    End If
                    If ULMALGFree.clsBase.CLast.Keys.Contains(d.ClaveIni) = True Then
                        If ULMALGFree.clsBase.CLast(d.ClaveIni).ToUpper = d.Local_File.ToUpper Then
                            ' No hay que actualizarlo. Ya estaba actualizad0
                            Continue For
                        End If
                    End If
                    oD = d
                    Exit For
                Next
                frmU = New frmUpdateAddin
                frmU.d = oD
                If oD IsNot Nothing Then
                    'MsgBox(frmU.d.ClaveIni)
                    frmU.Show()
                    frmU.Visible = True
                    While frmU.d IsNot Nothing
                        'System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(1)
                    End While
                    frmU.Close()
                    frmU = Nothing
                End If
                '
                ULMALGFree.clsBase.Bat_Borra()
                'If action.KILLUPDATE.ToString = queAccion Then
                ' Abrir siempre Revit, sin preguntar.
                'If MsgBox("¿Open Revit?", MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Question & MsgBoxStyle.YesNo, "ATTENTION") = MsgBoxResult.Yes Then
                'Dim cmd As String = uf.Process_DameFullPath ' ULMALGFree.clsBase.RevitFull
                If IO.File.Exists(cmd) Then
                    ULMALGFree.clsBase.Process_Run(cmd)
                    System.Threading.Thread.Sleep(2000)
                End If
                'End If
                'End If
            End If
        End If
    End Sub

    Public Enum action
        CHECK
        KILLUPDATE
        UPDATE
    End Enum
End Module

