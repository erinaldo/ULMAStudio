Module modVar
    Public Function INICargar() As String()
        If cIni Is Nothing Then cIni = New clsINI
        Dim mensaje(1) As String
        ' Mensaje(0) contendrá los errores.
        ' Mensaje(1) contendrá los valores leidos del .INI
        mensaje(0) = "" : mensaje(1) = ""
        '
        ' [PATHS] CARGAR DIRECTORIOS Y PLANTILLAS
        ' path_offlineBDIdata
        Dim path_offlineBDIdata As String = cIni.IniGet(ULMALGFree.clsBase._IniFull, "PATHS", "path_offlineBDIdata")
        ULMALGFree.clsBase.path_offlineBDIdata = ULMALGFree.clsBase.DameCaminoAbsolutoUsuarioFichero(path_offlineBDIdata)
        If IO.Directory.Exists(path_offlineBDIdata) Then
            mensaje(1) &= "path_offlineBDIdata = " & path_offlineBDIdata & vbCrLf
        Else
            mensaje(0) &= "The directory (path_offlineBDIdata) not exist... --> " & path_offlineBDIdata & vbCrLf
        End If
        '
        ' Directorio Librerias/Imagenes Delegation
        Dim path_families_base As String = cIni.IniGet(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_base")
        ULMALGFree.clsBase.path_families_base = ULMALGFree.clsBase.DameCaminoAbsolutoUsuarioFichero(path_families_base)
        If IO.Directory.Exists(path_families_base) Then
            mensaje(1) &= "path_families_base = " & path_families_base & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_base) not exist... --> " & path_families_base & vbCrLf
        End If
        Dim path_families_base_images As String = cIni.IniGet(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_base_images")
        ULMALGFree.clsBase.path_families_base_images = ULMALGFree.clsBase.DameCaminoAbsolutoUsuarioFichero(path_families_base_images)
        If IO.Directory.Exists(path_families_base_images) Then
            mensaje(1) &= "path_families_base_images = " & path_families_base_images & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_base_images) not exist... --> " & path_families_base_images & vbCrLf
        End If
        '
        '
        ' Directorio Librerias/Imagenes CUSTOM
        Dim path_families_custom As String = cIni.IniGet(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_custom")
        ULMALGFree.clsBase.path_families_custom = ULMALGFree.clsBase.DameCaminoAbsolutoUsuarioFichero(path_families_custom)
        If IO.Directory.Exists(path_families_custom) Then
            mensaje(1) &= "path_families_custom = " & path_families_custom & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_custom) not exist... --> " & path_families_custom & vbCrLf
        End If
        Dim path_families_custom_images As String = cIni.IniGet(ULMALGFree.clsBase._IniFull, "PATHS", "path_families_custom_images")
        ULMALGFree.clsBase.path_families_custom_images = ULMALGFree.clsBase.DameCaminoAbsolutoUsuarioFichero(path_families_custom_images)
        If IO.Directory.Exists(path_families_custom_images) Then
            mensaje(1) &= "path_families_custom_images = " & path_families_custom_images & vbCrLf
        Else
            mensaje(0) &= "The directory (path_families_custom_images) not exist... --> " & path_families_custom_images & vbCrLf
        End If
        '
        Return mensaje
    End Function
End Module
