Imports System.Windows.Forms
Imports System.Drawing

Partial Public Class clsBase
    Public Shared ReadOnly _updatesFolder As String = IO.Path.Combine(ULMALGFree.clsBase._LgFullFolder, "updates")
    Public Shared ReadOnly _unziptempFolder As String = IO.Path.Combine(ULMALGFree.clsBase._LgFullFolder, "_unziptemp")
    '

    Public Shared UltimoGrupo As Group
    Public Shared UltimoSuperGrupo As SuperGroup
    ''Private kill As Boolean = False
    'Private year As String = "R2019"
    Public Shared LUp As List(Of ULMALGFree.Datos)
    Public Shared CLast As Dictionary(Of String, String)
    Public Const nProcess As String = "REVIT"
    Public Shared cUp As Dictionary(Of String, List(Of String))
    '
    Public Shared lNumeros As ImageList = Nothing
    Public Shared lupdate As ImageList = Nothing
    'Public Shared fSuperGroups As FlowLayoutPanel
    'Public Shared fGroups As FlowLayoutPanel = Nothing
    'Public Shared pImagen As PictureBox = Nothing
    '
    Private Const _ExeUpdaterName As String = "ULMAUpdaterAddIn.exe"
    Private Const _IniUpdaterName As String = "ULMAUpdaterAddIn.ini"
    Private Const _BatUpdateName As String = "up.bat"
    Public Shared ReadOnly _IniUpdaterFull As String = IO.Path.Combine(_LgFullFolder, _IniUpdaterName)
    Public Shared ReadOnly _BatUpdateFull As String = IO.Path.Combine(_LgFullFolder, _BatUpdateName)
    Public Shared ReadOnly _ULMAUpdaterAddIn As String = IO.Path.Combine(_LgFullFolder, _ExeUpdaterName)
    Public Shared _ULMAUpdaterAddInActivo As Boolean = False

    Public Shared Sub Bat_CreaEjecuta(cerrarRevit As Boolean)
        Dim argumentos As String = ""
        If cerrarRevit Then
            'ULMALGFree.clsBase.Process_Run(_BatUpdateFull, "KILLUPDATE REVIT True", visible:=False)
            argumentos = " KILLUPDATE REVIT True"
        Else
            'ULMALGFree.clsBase.Process_Run(_BatUpdateFull, "UPDATE REVIT True", visible:=False)
            argumentos = " UPDATE REVIT True"
        End If
        'IO.File.WriteAllText(_BatUpdateFull, comillas & _ULMAUpdaterAddIn & comillas & " %1 %2 %3")
        IO.File.WriteAllText(_BatUpdateFull, "call " & comillas & _ULMAUpdaterAddIn & comillas & IIf(argumentos.StartsWith(" ") = True, argumentos, " " & argumentos))
        '
        If IO.File.Exists(_BatUpdateFull) = False OrElse argumentos = "" Then Exit Sub
        '
        Dim pInf As New ProcessStartInfo(_BatUpdateFull)
        'pInf.Arguments = argumentos
        pInf.WindowStyle = ProcessWindowStyle.Hidden


        Using p = Process.Start(pInf)
            'While p.HasExited = False
            '    System.Windows.Forms.Application.DoEvents()
            'End While
            'p.WaitForExit(15000)
        End Using
        '
        'Dim pRevit As Process() = System.Diagnostics.Process.GetProcessesByName(nProcess)
        'While pRevit IsNot Nothing AndAlso pRevit.Count > 0
        '    System.Windows.Forms.Application.DoEvents()
        '    pRevit = System.Diagnostics.Process.GetProcessesByName(nProcess)
        'End While
        'System.Threading.Thread.Sleep(2000)
    End Sub

    Public Shared Sub Bat_Borra()
        If IO.File.Exists(_BatUpdateFull) Then IO.File.Delete(_BatUpdateFull)
    End Sub

    Public Shared Sub INIUpdates_EscribeUPDATES(lDir As List(Of String), Optional borrarupdate As Boolean = False)
        ' Borrar primero los UPDATES que haya
        If borrarupdate Then cIni.IniDeleteSection(_IniUpdaterFull, "UPDATES")
        If lDir Is Nothing OrElse lDir.Count = 0 Then Exit Sub
        If cIni Is Nothing Then cIni = New clsINI
        If CLast Is Nothing OrElse CLast.Count = 0 Then Call INIUpdates_LeeLAST()
        '[UPDATES]
        'UCRevitFree = UCRevitFree_20191006.zip
        ';Ultimas actualzaciones realizadas de que cada AddIn.
        '[LAST]
        'UCRevitFree = UCRevitFree_20191006.zip
        For Each queFi As String In lDir
            Dim partes() As String = queFi.Split("_")
            Dim key As String = partes(0)
            '' Solo escribimos UPDATES si no está en LAST o es diferentes a LAST
            If CLast.ContainsKey(key) = False Then
                cIni.IniWrite(_IniUpdaterFull, "UPDATES", key, queFi)
            ElseIf CLast.ContainsKey(key) = True AndAlso CLast(key).ToUpper <> queFi.ToUpper Then
                cIni.IniWrite(_IniUpdaterFull, "UPDATES", key, queFi)
            End If
            'End If
        Next
    End Sub

    Public Shared Function INIUpdates_LeeTODO() As String
        If cIni Is Nothing Then cIni = New clsINI
        Dim resultado As String = ""
        resultado &= INIUpdates_LeeLAST()
        resultado = INIUpdates_LeeUPDATES()
        '' Actualizar las colecciones de fichero a descargar del FTP
        ULMALGFree.clsBase.cUp = ULMALGFree.clsBase.FTP_DameActualizacionesTodas()
        ' Poner los valores boolean que indican si hay actualizaciones en cada tipo.
        _ActualizarFamilies = False
        _ActualizarAddIns = False
        _ActualizarXMLs = False                '
        If cUp("addins").Count > 0 OrElse cUp("families").Count > 0 OrElse cUp("xmls").Count > 0 Then
            If cUp("families").Count > 0 Then _ActualizarFamilies = True
            If cUp("addins").Count > 0 Then _ActualizarAddIns = True
            If cUp("xmls").Count > 0 Then _ActualizarXMLs = True                '
        End If
        Return resultado
    End Function

    Public Shared Function INIUpdates_LeeLAST() As String
        If cIni Is Nothing Then cIni = New clsINI
        Dim resultado As String = ""
        ' CLast
        Try
            CLast = New Dictionary(Of String, String)
            Dim last() As String = cIni.IniGetSection(_IniUpdaterFull, "LAST")
            If last Is Nothing OrElse last.Count = 0 Then
                ' No hacemos nada.
            Else
                For i As Integer = 0 To UBound(last) - 1 Step 2
                    Dim Clave As String = last(i)
                    Dim nFile As String = last(i + 1)   ' Es una nombre de fichero, con la extensión .zip
                    If CLast.ContainsKey(Clave) = False Then
                        CLast.Add(Clave, nFile)
                    Else
                        CLast(Clave) = nFile
                    End If
                Next
            End If
        Catch ex As Exception
        End Try
        ' Actualizar las colecciones de ficheros del FTP que tenemos
        Return resultado
    End Function
    Public Shared Function INIUpdates_LeeUPDATES() As String
        If cIni Is Nothing Then cIni = New clsINI
        Dim resultado As String = ""
        ' LDatos
        LUp = New List(Of ULMALGFree.Datos)
        Dim updates() As String = cIni.IniGetSection(_IniUpdaterFull, "UPDATES")
        If updates Is Nothing OrElse updates.Count = 0 Then
            ' No hacemos nada
        Else
            For i As Integer = 0 To UBound(updates) - 1 Step 2
                Dim Clave As String = updates(i)
                Dim nFile As String = updates(i + 1)   ' Es una nombre de fichero, con la extensión .zip
                Dim d As New ULMALGFree.Datos(Clave, nFile)
                If LUp.Contains(d) = False Then LUp.Add(d)
                d = Nothing
            Next
        End If
        Return resultado
    End Function

    Public Shared Function LDatos_DameSoloTipoX(tipo As ULMALGFree.FOLDERWEB) As List(Of ULMALGFree.Datos)
        Dim res = From x In LUp
                  Where x.Tipo = tipo
                  Select x
        '
        Return res.ToList
    End Function
    Public Shared Function FTP_DameListaNombresFicheros(queDir As String) As List(Of String)
        Dim resultado As New List(Of String)
        Dim lista() As String = cFtp.FTP_ListaDir(queDir, fullpath:=False)
        Dim ListaFin As IEnumerable(Of String) = From x In lista
                                                 Where x <> ""
                                                 Select x
        If ListaFin IsNot Nothing AndAlso ListaFin.Count > 0 Then
            resultado = ListaFin.ToList
        End If
        Return resultado
    End Function

    Public Shared Function FTP_DameActualizacionesTodas() As Dictionary(Of String, List(Of String))
        Dim resultado As New Dictionary(Of String, List(Of String))
        Dim lDir As New List(Of String)
        ' Leer LAST
        If CLast Is Nothing OrElse CLast.Count = 0 Then INIUpdates_LeeLAST()
        ' addins
        Dim lista() As String = cFtp.FTP_ListaDir(ULMALGFree.clsBase.FTP1_dirAddins, fullpath:=False)
        If lista.Count > 0 Then
            lista = ArrayFullWeb_DameSoloNombreExtension(lista)
            lDir.AddRange(lista)
        End If
        INIUpdates_EscribeUPDATES(lDir, borrarupdate:=True)
        resultado.Add("addins", lDir)
        lDir = New List(Of String)
        '
        ' families
        lista = cFtp.FTP_ListaDir(ULMALGFree.clsBase.FTP1_dirFamilies)
        If lista.Count > 0 Then
            lista = ArrayFullWeb_DameSoloNombreExtension(lista)
            lDir.AddRange(lista)
        End If
        INIUpdates_EscribeUPDATES(lDir, borrarupdate:=False)
        resultado.Add("families", lDir)
        lDir = New List(Of String)
        '
        ' xml
        lista = cFtp.FTP_ListaDir(ULMALGFree.clsBase.FTP1_dirXml)
        If lista.Count > 0 Then
            lista = ArrayFullWeb_DameSoloNombreExtension(lista)
            lDir.AddRange(lista)
        End If
        INIUpdates_EscribeUPDATES(lDir, borrarupdate:=False)
        resultado.Add("xmls", lDir)
        '
        Return resultado
    End Function
    '
    Public Shared Function ArrayFullWeb_DameSoloNombreExtension(arr1 As String()) As String()
        Dim resultado As New List(Of String)
        For Each x As String In arr1
            Dim partes() As String = Nothing
            If x.Contains("\") Then
                partes = x.Split("\"c)
            ElseIf x.Contains("/") Then
                partes = x.Split("/")
            Else
                partes = {x}
            End If
            Dim solonombre As String = ""
            solonombre = partes(UBound(partes))
            If solonombre.ToUpper.StartsWith("ERROR") = False AndAlso FicheroFTP_YaEstaDescargado(solonombre) = False Then
                resultado.Add(solonombre)
            End If
        Next
        Return resultado.ToArray
    End Function

    Public Shared Function FicheroFTP_YaEstaDescargado(fiName As String) As Boolean
        If cIni Is Nothing Then cIni = New clsINI
        Dim resultado As Boolean = False
        If IO.File.Exists(ULMALGFree.clsBase._IniUpdaterFull) = False Then
            Return resultado
            Exit Function
        End If
        '
        '' Leer LAST
        Call INIUpdates_LeeLAST()
        If CLast Is Nothing OrElse CLast.Count = 0 Then
            Return resultado
            Exit Function
        End If
        '
        Dim partes() As String = fiName.Split("_")
        Dim key As String = partes(0)
        '
        If CLast.ContainsKey(key) Then
            If CLast(key) <> fiName Then
                resultado = False
            ElseIf CLast(key) = fiName Then
                resultado = True
            End If
        Else
            resultado = False
        End If

        Return resultado
    End Function

    Public Shared Sub DescargaFicheroFTPUCRevitFree(fWeb As FOLDERWEB, dirLocal As String, ByVal Fichero As String,
            Optional ByRef quePb As System.Windows.Forms.ProgressBar = Nothing)

        Dim terminado As Boolean = False
        If cFtp Is Nothing Then cFtp = New clsFTP(FTP1_host, FTP1_dirLog, FTP1_user, FTP1_userpass)
        If IO.Directory.Exists(dirLocal) = False Then
            IO.Directory.CreateDirectory(dirLocal)
        End If
        Select Case fWeb
            Case FOLDERWEB.Addins
                Call cFtp.DescargarFicheroFTP(FTP1_dirAddins, FTP1_user, FTP1_userpass, dirLocal, Fichero, terminado, quePb)
            Case FOLDERWEB.Families
                Call cFtp.DescargarFicheroFTP(FTP1_dirFamilies, FTP1_user, FTP1_userpass, dirLocal, Fichero, terminado, quePb)
            Case FOLDERWEB.XML
                Call cFtp.DescargarFicheroFTP(FTP1_dirXml, FTP1_user, FTP1_userpass, dirLocal, Fichero, terminado, quePb)
        End Select

        While terminado = False
            System.Windows.Forms.Application.DoEvents()
        End While
    End Sub

    Public Shared Sub FTP_DescargarYDescomprimir(ByRef oGroup As Group,
                                           d As Datos,
                                          ByRef LblAccion As System.Windows.Forms.Label,
                                          ByRef Pb1 As System.Windows.Forms.ProgressBar)
        '
        Dim correcto As Boolean = False
        ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
        If IO.Directory.Exists(d.TargetFolder) = False AndAlso d.TargetFolder <> "" Then
            Try
                IO.Directory.CreateDirectory(d.TargetFolder)
            Catch ex As Exception
                'Debug.Print(ex.ToString)
                MsgBox(ex.ToString, MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
                'Continue For
                Exit Sub
            End Try
        End If

        If IO.File.Exists(d.Local_FileFull) = False Then
            ' Descargar el fichero si no está en la carpeta Updates
            LblAccion.Text = "Downloading... " & d.Local_File.Split("_"c)(1) : LblAccion.Refresh()
            ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(d.Tipo, IO.Path.GetDirectoryName(d.Local_FileFull), d.Local_File, Pb1)
        End If

        If IO.File.Exists(d.Local_FileFull) Then
            ' Descomprimir
            Try
                LblAccion.Text = "UnZip... " & d.Local_File : LblAccion.Refresh()
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
            cIni.IniWrite(_IniUpdaterFull, "LAST", d.ClaveIni, d.Local_File)
            cIni.IniDeleteKey(_IniUpdaterFull, "UPDATES", d.ClaveIni)
            oGroup.gButton2.Image = oGroup.Grupo_PonImageAction()
            UltimoSuperGrupo.nActualizaciones -= 1
            '
            ' Borrar el fichero .zip 
            Try
                IO.File.Delete(d.Local_FileFull)
            Catch ex As Exception

            End Try
        End If
    End Sub
    '
    Public Shared Sub FTP_DescargarYDescomprimirXML(d As Datos)
        '
        Dim correcto As Boolean = False
        ' Crear el directorio destino, si no existe. Para que lo podamos descargar ahí.
        If IO.Directory.Exists(d.TargetFolder) = False AndAlso d.TargetFolder <> "" Then
            Try
                IO.Directory.CreateDirectory(d.TargetFolder)
            Catch ex As Exception
                'Debug.Print(ex.ToString)
                MsgBox(ex.ToString, MsgBoxStyle.MsgBoxSetForeground Or MsgBoxStyle.Critical, "ATTENTION")
                'Continue For
                Exit Sub
            End Try
        End If

        If IO.File.Exists(d.Local_FileFull) = False Then
            ' Descargar el fichero si no está en la carpeta Updates
            ULMALGFree.clsBase.DescargaFicheroFTPUCRevitFree(d.Tipo, IO.Path.GetDirectoryName(d.Local_FileFull), d.Local_File)
        End If

        If IO.File.Exists(d.Local_FileFull) Then
            ' Descomprimir
            Try
                modZIP.DescomprimeZIP_UnoAUno(d.Local_FileFull, d.TargetFolder)
                correcto = True
            Catch ex As Exception
                correcto = False
            End Try
        Else
            correcto = False
        End If
        '
        If correcto = True Then
            ' Log del XML
            If yo Is Nothing Then yo = New clsBase(Reflection.Assembly.GetExecutingAssembly)
            yo.PonLog_ULMA(ULMALGFree.ACTION.UPDATE_XML,,,,,, d.Local_File)
            '
            cIni.IniWrite(_IniUpdaterFull, "LAST", d.ClaveIni, d.Local_File)
            cIni.IniDeleteKey(_IniUpdaterFull, "UPDATES", d.ClaveIni)
            ' Borrar el fichero .zip 
            Try
                IO.File.Delete(d.Local_FileFull)
            Catch ex As Exception

            End Try
            '
            Dim t As Task = New Task(AddressOf BorraFamiliasGroup_SoloResiduales) : t.Start()
        End If
    End Sub
    '
    Public Shared Sub BorraFamiliasGroup_SoloResiduales()
        If colGroups Is Nothing OrElse colGroups.Count = 0 Then Exit Sub
        '
        Dim LBorrarEnd As New List(Of String)
        Dim fiRFAs() As String = IO.Directory.GetFiles(path_families_base, "*.rfa", IO.SearchOption.TopDirectoryOnly)
        ' Recorrer todos los Grupos, rellenar LstFilenameOnly
        Dim estaengrupo As Boolean = False
        For Each oGTemp As clsGroup In colGroups.Values
            ' Rellenar LstFilenameOnly de este Grupo
            oGTemp.RellenaFilenameOnly()
            System.Threading.Thread.Sleep(1)
        Next
        '
        For Each fiRFA In fiRFAs
            For Each oGTemp As clsGroup In colGroups.Values
                If oGTemp.LstFilenameOnly.Contains(IO.Path.GetFileNameWithoutExtension(fiRFA)) = True Then
                    estaengrupo = True
                    Exit For
                End If
                System.Threading.Thread.Sleep(1)
            Next
            ' Si no está en otros grupos, añadir a la lista definitiva de borrado
            If estaengrupo = False Then
                LBorrarEnd.Add(fiRFA)
            End If
            System.Threading.Thread.Sleep(1)
        Next
        ' Finalmente. Borrar los que no estaban en ningún grupo
        For Each fiBorra In LBorrarEnd
            Dim fiImg As String = IO.Path.Combine(path_families_base_images, IO.Path.GetFileNameWithoutExtension(fiBorra) & ".png")
            Try
                If IO.File.Exists(fiBorra) Then IO.File.Delete(fiBorra)
                If IO.File.Exists(fiImg) Then IO.File.Delete(fiImg)
            Catch ex As Exception

            End Try
            System.Threading.Thread.Sleep(1)
        Next
        LBorrarEnd = Nothing
    End Sub
End Class