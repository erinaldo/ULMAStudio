Public Class Datos
    Public ClaveIni As String = ""
    Public Local_File As String = ""  ' UCRevitFree_20191005.zip      (Estará en /data/R201X/addins)
    Public Local_FileFull As String = ""
    Public TargetFolder As String = ""
    Public TargetFolderImg As String = ""
    Public Tipo As ULMALGFree.FOLDERWEB

    Public Sub New(claIni As String, nFile As String)
        ClaveIni = claIni
        If nFile.EndsWith(".zip") = False Then nFile &= ".zip"
        Local_File = nFile
        '
        ' TargetFolder y TargetFolderImg. Todos a ULMALGFree.clsBase._LgFullFolder (los .zip ya llevar los subdirectorios)
        If ClaveIni.StartsWith(ULMALGFree.clsBase._AddInName) Then
            ' addins
            Tipo = ULMALGFree.FOLDERWEB.Addins
            TargetFolder = ULMALGFree.clsBase._LgFullFolder
        ElseIf ClaveIni.ToUpper.StartsWith("XML") Then
            ' xmls
            Tipo = ULMALGFree.FOLDERWEB.XML
            TargetFolder = ULMALGFree.clsBase._LgFullFolder    'ULMALGFree.clsBase.path_offlineBDIdata
        Else
            ' families
            Tipo = ULMALGFree.FOLDERWEB.Families
            TargetFolder = ULMALGFree.clsBase._LgFullFolder     ' El .zip ya lleva families y families_images   'ULMALGFree.clsBase.path_families_base
            TargetFolderImg = ULMALGFree.clsBase._LgFullFolder  ' TargetFolder & "_images"                      'ULMALGFree.clsBase.path_families_base_images
        End If
        '
        '
        Local_FileFull = IO.Path.Combine(ULMALGFree.clsBase._updatesFolder, Local_File)
    End Sub
End Class