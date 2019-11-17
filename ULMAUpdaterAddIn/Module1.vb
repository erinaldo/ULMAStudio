Imports uf = ULMALGFree.clsBase
Module Module1
    Public frmU As frmUpdateAddin
    Public FullPathApp As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    Public FolderApp As String = IO.Path.GetDirectoryName(FullPathApp)
    Public FullPathZip As String = ""
    Public NameZip As String = ""
    Public KeyZip As String = ""
    Public FullPathIni As String = ""
    Public FullPathRevit As String = ""
    Public ConDescarga As Boolean = False
    Public correcto As Boolean = False
    'Public _appFolderUpdates As String = IO.Path.Combine(_appFolder, "Updates\")
    'Public _iniFull As String = IO.Path.ChangeExtension(_appFull, ".ini")
    '
    Public cIni As clsINI
    Public cLcsv As ULMALGFree.clsBase
    Public releerFTP As Boolean = False

    Sub Main()
        'If cLcsv Is Nothing Then cLcsv = New ULMALGFree.clsBase(System.Reflection.Assembly.GetExecutingAssembly)
        If cIni Is Nothing Then cIni = New clsINI()
        '
        'Console.CursorVisible = False
        Dim parametros() As String = Environment.GetCommandLineArgs
        ' Ejemplo: ULMAUpdaterAddin.exe REVIT True .\Updates\UCRevitFree_20191005.zip %appdata%\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree\
        ' 0 = Camino completo del .exe
        ' 1 = FullPath del fichero a descomprimir (Subcarpeta updates\[fichero.zip]
        ' 2 = FullPath del fichero ULMAUpdaterAddIn.ini a cambiar
        ' 3 = FullPath de Revit. Para volver a abrirlo al terminar
        ' 4 = True (Con descarga) False (Solo descomprime, no descarga)
        Try
            IO.File.Delete(uf._BatUpdateFull)
        Catch ex As Exception
        End Try
        ' Comprobar y coger los parámetros de la linea de comandos.
        If parametros.Count < 4 Then
            MsgBox("Parameter missing or error", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "ATTENTION")
            End
        End If
        '
        FullPathZip = parametros(1)
        NameZip = IO.Path.GetFileName(FullPathZip)
        KeyZip = NameZip.Split("_")(0)
        FullPathIni = parametros(2)
        FullPathRevit = parametros(3)
        If parametros.Count = 5 Then
            Try
                ConDescarga = Convert.ToBoolean(parametros(4))
            Catch ex As Exception
            End Try
        End If
        '
        Dim resultado As String() = INICargar()
        '
        If IO.File.Exists(FullPathZip) = False Then
            MsgBox("No update file available", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "ATTENTION")
            End
        End If
        '
        ' KILLUPDATE. Terminar proceso Revit
        Dim p As Process() = System.Diagnostics.Process.GetProcessesByName(ULMALGFree.clsBase.nProcess)
        If p Is Nothing OrElse p.Count = 0 Then
            End
        End If
        '
        p = Nothing
        Do
            ULMALGFree.clsBase.Process_Close(ULMALGFree.clsBase.nProcess)
            Threading.Thread.Sleep(2000)
            System.Windows.Forms.Application.DoEvents()
            p = System.Diagnostics.Process.GetProcessesByName(ULMALGFree.clsBase.nProcess)
        Loop While p IsNot Nothing AndAlso p.Count > 0
        'Threading.Thread.Sleep(2000)
        '
        DescomprimirAddIn()
        '
        ' Finalmente. Volver a abrir Revit
        'MsgBox(FullPathRevit)
        If IO.File.Exists(FullPathRevit) Then
            Dim pInf As New ProcessStartInfo(FullPathRevit)
            Call Process.Start(pInf)
            System.Threading.Thread.Sleep(2000)
        Else
            MsgBox("Error opening Revit", MsgBoxStyle.Critical, "ATTENTION")
        End If
    End Sub
    Private Sub DescomprimirAddIn()
        ' DESCARGAR
        'If ConDescarga = True Then
        '    ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
        '    If IO.Directory.Exists(IO.Path.GetDirectoryName(FullPathZip)) = False Then
        '        Try
        '            IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(FullPathZip))
        '        Catch ex As Exception
        '            MsgBox("Error creating updates folder", MsgBoxStyle.Critical, "ATTENTION")
        '            End
        '        End Try
        '    End If
        '    If IO.File.Exists(FullPathZip) = False Then
        '        ' Descargar el fichero si no está en la carpeta Updates
        '        ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(ULMALGFree.FOLDERWEB.Addins, uf._updatesFolder, NameZip)
        '    End If
        'End If
        '
        ' DESCOMPRIMIR
        If IO.File.Exists(FullPathZip) Then
            ' Descomprimir
            Try
                modZIP.DescomprimeZIP_UnoAUno(FullPathZip, FolderApp, renombraExes:=True)
                correcto = True
            Catch ex As Exception
                correcto = False
            End Try
        Else
            correcto = False
        End If
        '
        If correcto = True Then
            cIni.IniWrite(uf._IniUpdaterFull, "LAST", KeyZip, NameZip)
            cIni.IniDeleteKey(uf._IniUpdaterFull, "UPDATES", KeyZip)
            '
            ' Borrar el fichero .zip 
            Try
                IO.File.Delete(FullPathZip)
            Catch ex As Exception

            End Try
        End If
    End Sub
    'Private Sub ActualizarAddIn()
    '    If d Is Nothing Then
    '        'MsgBox("D IS NOTHING")
    '        Exit Sub
    '    End If
    '    ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
    '    If IO.Directory.Exists(d.TargetFolder) = False AndAlso d.TargetFolder <> "" Then
    '        Try
    '            IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(d.TargetFolder))
    '        Catch ex As Exception
    '            'Debug.Print(ex.ToString)
    '            'MsgBox(ex.ToString, MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
    '            'Continue For
    '            Me.Close()
    '        End Try
    '    End If

    '    If IO.File.Exists(d.Local_FileFull) = False Then
    '        ' Descargar el fichero si no está en la carpeta Updates
    '        LblAccion.Text = "Downloading... " & IO.Path.GetFileNameWithoutExtension(d.Local_File).Replace("_", " ") : LblAccion.Refresh()
    '        ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(d.Tipo, IO.Path.GetDirectoryName(d.Local_FileFull), d.Local_File, Pb1)
    '    End If

    '    If IO.File.Exists(d.Local_FileFull) Then
    '        ' Descomprimir
    '        Try
    '            If d.TargetFolder.Contains("ULMACONSTRUCCION") Then d.TargetFolder &= "\DOWNLOAD"
    '            modZIP.DescomprimeZIP_UnoAUno(d.Local_FileFull, d.TargetFolder, Pb1, LblAccion, renombraExes:=True)
    '            correcto = True
    '        Catch ex As Exception
    '            correcto = False
    '        End Try
    '    Else
    '        correcto = False
    '    End If
    '    '
    '    If correcto = True Then
    '        cIni.IniWrite(uf._IniUpdaterFull, "LAST", d.ClaveIni, d.Local_File)
    '        cIni.IniDeleteKey(uf._IniUpdaterFull, "UPDATES", d.ClaveIni)
    '        '
    '        ' Borrar el fichero .zip 
    '        Try
    '            IO.File.Delete(d.Local_FileFull)
    '        Catch ex As Exception

    '        End Try
    '    End If
    '    Me.d = Nothing
    'End Sub
End Module

