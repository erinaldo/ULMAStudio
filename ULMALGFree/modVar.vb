Imports System.Windows.Forms
Imports uf = ULMALGFree.clsBase

Module modVar
    'Public frmIni As frmUpdater
    'Public _appPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    'Public _appFolder As String = IO.Path.GetDirectoryName(_appPath)
    'Public _xmlPath As String = IO.Path.Combine(_appFolder, "120_publicStructure.xml")
    'Public _imgPath As String = IO.Path.Combine(_appFolder, "IMAGENES\")
    'Public _imgBase As String = IO.Path.Combine(_imgPath, "MEGAFORM_render.png")

    'Function Aleatorio_Dame(min As Integer, max As Integer) As Integer
    '    Static Random_Number As New Random()
    '    Return Random_Number.Next(min, max + 1) ' le sumamos 1
    'End Function

    '[UPDATES]
    'UCRevitFree = R2019·UCRevitFree_20191006.zip·C:\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    'E1904=R2019·E1904_(20190613_R2018_P_MEGAFORM).zip·C\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    '[UPDATESNOT]
    'UCRevitFree=R2019·UCRevitFree_20191007.zip·C:\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    'E1904=R2019·E1904_(20190613_R2018_P_MEGAFORM).zip·C\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    'X0004=R2019·X0004_(20190612_R2018_P_ENKOFLEX).zip·C:\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    'X0007=R2019·X0007_(20190612_R2018_P_MEGALITE).zip·C:\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    'R2019=R2019·R2019_UCP_XML_20190603.zip·C:\Users\alberto.ADA\AppData\Roaming\Autodesk\Revit\Addins\2019\UCRevitFree
    ';Ultimas actualzaciones realizadas de que cada AddIn.
    '[LAST]
    'UCRevitFree=20191007
    'E1904_MEGAFORM=20190612
    'X0004_ENKOFLEX=20190612
    'X0007_MEGALITE=20190612


    Function Aleatorio_Dame(min As Integer, max As Integer) As Integer
        Static Random_Number As New Random()
        Return Random_Number.Next(min, max + 1) ' le sumamos 1
    End Function
    Public Function Grupo_DameActionNumero(gCode As String, gSName As String) As Integer
        Dim resultado As Integer = 0
        Dim lupUno As ULMALGFree.Datos
        Call uf.INIUpdates_LeeTODO()
        '
        ' No hay actualizaciones. Salir con 0 (updated) Actualizado
        If uf.LUp.Count = 0 Then
            If uf.CLast.ContainsKey(gCode) Then
                lupUno = New ULMALGFree.Datos(gCode, uf.CLast(gCode))
            Else
                resultado = 1
                Return resultado
                Exit Function
            End If
            'End If
        Else
            '
            Dim LUpBusco As IEnumerable(Of ULMALGFree.Datos) = From x In uf.LUp
                                                               Where x.ClaveIni.ToUpper = gCode.ToUpper OrElse x.ClaveIni.ToUpper = gSName.ToUpper
                                                               Select x
            ' No está en actualizaciones
            If LUpBusco Is Nothing Then
                resultado = 0
                Return resultado
                Exit Function
            End If
            If LUpBusco.Count = 0 Then
                resultado = 0
                Return resultado
                Exit Function
            End If
            '
            ' Si estaba en UPDATE
            lupUno = LUpBusco.First
            '
        End If
        If uf.CLast.ContainsKey(lUpUno.ClaveIni) Then
            If uf.CLast(lUpUno.ClaveIni) = lUpUno.Local_File Then
                ' Ya se descargó y está actualizado
                resultado = 0
            Else
                ' Hay que descargarlo. Y ya se descargó anteriormente. Es una actualización porque el fichero es diferente.
                resultado = 2
            End If
        Else
            ' Hay que descargarlo. Y no se ha descargado antes.
            resultado = 1
        End If
        Return resultado
    End Function
End Module
