Imports System.Windows.Forms
Imports uf = ULMALGFree.clsBase

Public Class frmUpdateAddin
    Public correcto As Boolean = False
    Public d As ULMALGFree.Datos = Nothing
    Private Sub FrmUpdateAddin_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' No detectar llamadas a controles desde otros procesos (False) Si queremos detectarlas sería True.
        'Control.CheckForIllegalCrossThreadCalls = False
        ActualizarAddIn()
    End Sub

    Private Sub ActualizarAddIn()
        If d Is Nothing Then
            'MsgBox("D IS NOTHING")
            Exit Sub
        End If
        ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
        If IO.Directory.Exists(d.TargetFolder) = False AndAlso d.TargetFolder <> "" Then
            Try
                IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(d.TargetFolder))
            Catch ex As Exception
                'Debug.Print(ex.ToString)
                'MsgBox(ex.ToString, MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
                'Continue For
                Me.Close()
            End Try
        End If

        If IO.File.Exists(d.Local_FileFull) = False Then
            ' Descargar el fichero si no está en la carpeta Updates
            LblAccion.Text = "Downloading... " & IO.Path.GetFileNameWithoutExtension(d.Local_File).Replace("_", " ") : LblAccion.Refresh()
            ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(d.Tipo, IO.Path.GetDirectoryName(d.Local_FileFull), d.Local_File, Pb1)
        End If

        If IO.File.Exists(d.Local_FileFull) Then
            ' Descomprimir
            Try
                If d.TargetFolder.Contains("ULMACONSTRUCCION") Then d.TargetFolder &= "\DOWNLOAD"
                modZIP.DescomprimeZIP_UnoAUno(d.Local_FileFull, d.TargetFolder, Pb1, LblAccion, renombraExes:=True)
                correcto = True
            Catch ex As Exception
                correcto = False
            End Try
        Else
            correcto = False
        End If
        '
        If correcto = True Then
            cIni.IniWrite(uf._IniUpdaterFull, "LAST", d.ClaveIni, d.Local_File)
            cIni.IniDeleteKey(uf._IniUpdaterFull, "UPDATES", d.ClaveIni)
            '
            ' Borrar el fichero .zip 
            Try
                IO.File.Delete(d.Local_FileFull)
            Catch ex As Exception

            End Try
        End If
        Me.d = Nothing
    End Sub
End Class