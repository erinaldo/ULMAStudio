'Imports ConsultarBDI
Imports System.Windows.Forms
Imports adWin = Autodesk.Windows
Imports Autodesk.Revit.DB
Imports uf = ULMALGFree.clsBase
Module modULMA
    ''
    Private Const categoryview As String = "VIEW TYPE"
    Public formHide As System.Windows.Forms.Form = Nothing
    Public compruebacambios As Boolean = False
    Public preInf As String = ""
    Public lstErrores As List(Of String)
    '

    Public Sub Cierra_DocumentosUlma(oApp As Autodesk.Revit.ApplicationServices.Application, Optional soloULMA As Boolean = True)
        If oApp.Documents.Size = 0 Then Exit Sub
        '
        Dim cerraractual As Boolean = False
        For Each queDoc As Autodesk.Revit.DB.Document In oApp.Documents
            Try
                If soloULMA Then
                    Dim colIds As List(Of ElementId) = utilesRevit.FamilySymbol_DameULMA_ID(queDoc, BuiltInCategory.OST_GenericModel)
                    If colIds IsNot Nothing AndAlso colIds.Count > 0 Then
                        If queDoc.PathName <> evRevit.evAppUI.ActiveUIDocument.Document.PathName Then
                            queDoc.Close(True)
                        Else
                            cerraractual = True
                        End If
                    End If
                    colIds = Nothing
                Else
                    If queDoc.PathName <> evRevit.evAppUI.ActiveUIDocument.Document.PathName Then
                        queDoc.Close(True)
                    Else
                        cerraractual = True
                    End If
                End If
            Catch ex As Exception
                Continue For
            End Try
        Next
        '
        If cerraractual Then
            Dim rvtIntPtr As IntPtr = Autodesk.Windows.ComponentManager.ApplicationWindow
            clsAPI.SetForegroundWindow(rvtIntPtr)
            Try
                clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 0, UIntPtr.Zero)
                clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 2, UIntPtr.Zero)
                clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 0, UIntPtr.Zero)
                clsAPI.keybd_event(CByte(clsAPI.eMensajes.VK_ESCAPE), 0, 2, UIntPtr.Zero)
                System.Windows.Forms.SendKeys.SendWait("^{F4}")
                'System.Windows.Forms.SendKeys.SendWait("^W")
            Catch ex As Exception
                Debug.Print(ex.ToString)
            End Try
        End If
    End Sub
    '
    '    Public Sub LeeEscribeDatos(doc As Autodesk.Revit.DB.Document, todo As Boolean, Optional auto As Boolean = False)
    '        enejecucion = True
    '        '
    '        Dim vs As ViewSchedule = utilesRevit.ScheduleDameUna(evRevit.evAppUI.ActiveUIDocument.Document, bom_base)    ' "#BOM-WK")
    '        '*** Iniciar la codificación
    '        Dim oTabla As Autodesk.Revit.DB.TableData = vs.GetTableData
    '        'Dim oHead As Autodesk.Revit.DB.TableSectionData = oTabla.GetSectionData(SectionType.Header)
    '        Dim oBody As Autodesk.Revit.DB.TableSectionData = oTabla.GetSectionData(SectionType.Body)
    '        'Dim def As IList(Of SchedulableField) = vs.Definition.GetSchedulableFields
    '        ''
    '        Dim mensaje As String = ""
    '        Dim errores As String = ""
    '        ''
    '        Dim filacabeceras As Integer = 0
    '        Dim filaIniDatos As Integer = 2
    '        ''
    '        Dim colTYPE As Integer = -1
    '        Dim colFamily As Integer = -1
    '        Dim colFAMILY_CODE As Integer = -1
    '        Dim colITEM_CODE As Integer = -1
    '        Dim colITEM_LENGTH As Integer = -1
    '        Dim colITEM_WIDTH As Integer = -1
    '        Dim colITEM_HEIGHT As Integer = -1
    '        Dim colFILTER_ID As Integer = -1
    '        Dim colITEM_WEIGHT As Integer = -1
    '        Dim colWEIGHT As Integer = -1
    '        Dim colITEM_DESCRIPTION As Integer = -1
    '        Dim colCount As Integer = -1
    '        Dim colW_MARKET As Integer = -1
    '        Dim colITEM_GENERIC As Integer = -1
    '        'Dim colPhaseCreated As Integer = 0
    '        ''
    '        Dim filaIni As Integer = oBody.NumberOfRows
    '        Dim columnas As Integer = oBody.NumberOfColumns
    '        ''
    '        '' Ponemos a 0, fijamos máximo y escribimos.
    '        PonInf(preInf & "Phase 1/4" & vbCrLf & "Processing Columns", True, oBody.NumberOfRows)
    '        ''
    '        '' ***** Recorremos la cabecera para coger los números de columna de cada dato que necesitamos.
    '        Dim colColumnas As New Hashtable
    '        For queCol As Integer = oBody.FirstColumnNumber To oBody.LastColumnNumber
    '            Dim valor As String = vs.GetCellText(SectionType.Body, filacabeceras, queCol)
    '            Select Case valor.ToUpper
    '                Case nType.ToUpper, nType     ',"TYPE"           ' FAMILY es como se llama originalmente. Lo renombramos a TYPE (Por si se les olvida renombrarlo)
    '                    colTYPE = queCol
    '                    'Case "TYPE"     ',"TYPE"           ' FAMILY es como se llama originalmente. Lo renombramos a TYPE (Por si se les olvida renombrarlo)
    '                    '    colTYPE = queCol
    '                    '    'nType = "TYPE"
    '                    'Case "TIPO"     ',"TYPE"           ' FAMILY es como se llama originalmente. Lo renombramos a TYPE (Por si se les olvida renombrarlo)
    '                    '    colTYPE = queCol
    '                    '    'nType = "TIPO"
    '                Case "FAMILY", nFamily.ToUpper    ' "FAMILY"
    '                    colFamily = queCol
    '                Case "FAMILY_CODE"
    '                    colFAMILY_CODE = queCol
    '                Case "ITEM_CODE"
    '                    colITEM_CODE = queCol
    '                Case "ITEM_LENGTH"
    '                    colITEM_LENGTH = queCol
    '                Case "ITEM_WIDTH"
    '                    colITEM_WIDTH = queCol
    '                Case "ITEM_HEIGHT"
    '                    colITEM_HEIGHT = queCol
    '                Case "FILTER_ID"
    '                    colFILTER_ID = queCol
    '                Case "ITEM_WEIGHT"
    '                    colITEM_WEIGHT = queCol
    '                Case "WEIGHT"
    '                    colWEIGHT = queCol
    '                Case "ITEM_DESCRIPTION"
    '                    colITEM_DESCRIPTION = queCol
    '                Case nCount.ToUpper, nCount
    '                    colCount = queCol
    '                Case "W_MARKET"
    '                    colW_MARKET = queCol
    '                Case "ITEM_GENERIC"
    '                    colITEM_GENERIC = queCol
    '                    'Case "Phase Created"
    '                    'colPhaseCreated = queCol
    '            End Select
    '            '' Poner las columnas COUNT y TYPE
    '            colCount = 3
    '            colTYPE = 5
    '            '' Llenamos colColumnas
    '            If colColumnas.ContainsKey(queCol) = False Then colColumnas.Add(queCol, valor.ToUpper)
    '            ''
    '            PonInf()
    '        Next
    '        '' 2017/03/24 Quitamos ITEM_MARKET de todo, menos de crear parametros proyecto desde fichero parámetros compartidos
    '        '' Y en insertar familias, donde debe rellenarlo en familias ULMA.
    '        '' *** Comprobar cabeceras, para ver si están todos los campos que necesitamos. Error y salir si no están.
    '        If colTYPE = -1 OrElse colFAMILY_CODE = -1 OrElse colITEM_CODE = -1 OrElse colITEM_LENGTH = -1 OrElse colITEM_WIDTH = -1 OrElse
    '                colITEM_HEIGHT = -1 OrElse colFILTER_ID = -1 OrElse colITEM_WEIGHT = -1 OrElse colITEM_DESCRIPTION = -1 OrElse colCount = -1 OrElse colW_MARKET = -1 OrElse
    '                colITEM_GENERIC = -1 Then
    '            ''
    '            errores &= bom_base & " must have the following columns and order:" & vbCrLf &
    '                "ITEM_CODE, ITEM_DESCRIPTION, ITEM_WEIGHT," & nCount & ", W_MARKET, " & nType & ", FAMILY_CODE, ITEM_LENGTH, ITEM_WIDTH, ITEM_HEIGHT, ITEM_GENERIC and FILTER_ID" & vbCrLf &
    '                "To be able to read, search and perform calculations..."
    '            '
    '            GoTo FINAL
    '            Exit Sub
    '        End If
    '        '
    '        ' Ponemos a 0, fijamos máximo y escribimos.
    '        PonInf(preInf & "Phase 2/4" & vbCrLf & "Saving Current Data", True, oBody.NumberOfRows * oBody.NumberOfColumns)
    '        ''
    '        '' Sacamos todos los datos actuales a un fichero.
    '        '' Y rellenamos el Hashtable colFilas con ULMALGFree.clsFamilia
    '        colFilas = New Hashtable
    '        For fi As Integer = oBody.FirstRowNumber To oBody.LastRowNumber
    '            Dim valorcoluno As String = vs.GetCellText(SectionType.Body, fi, oBody.FirstRowNumber)
    '            valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 1)
    '            valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 2)
    '            valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 3)
    '            '' Saltamos las filas vacias (la fila 2) y la ultima con los totales.
    '            If valorcoluno = "" OrElse valorcoluno.ToUpper.Contains("TOTAL") OrElse valorcoluno.Contains(":") Then
    '                ''
    '                PonInf()
    '                Continue For
    '            End If
    '            ''
    '            '' Crear la clase uf.cFam de clsFamilias
    '            If fi > oBody.FirstRowNumber AndAlso uf.cFam Is Nothing Then
    '                uf.cFam = New ULMALGFree.clsFamilia(fi, colColumnas)
    '                '' Añadimos la clase uf.cFam a colFilas
    '                colFilas.Add(fi, uf.cFam)
    '            End If
    '            ''
    '            '' Creamos la clase uf.cFam y la ponemos en la colección colFilas (key=nº fila, value=uf.cFam)
    '            '' No tenemos encuenta la fila de cabeceras.
    '            mensaje &= fi & "="
    '            For co As Integer = oBody.FirstColumnNumber To oBody.LastColumnNumber
    '                PonInf()
    '                Dim valor As String = vs.GetCellText(SectionType.Body, fi, co)
    '                ''
    '                If fi > oBody.FirstRowNumber AndAlso uf.cFam IsNot Nothing Then
    '                    CType(colFilas(fi), ULMALGFree.clsFamilia).PonDato(co, valor)
    '                End If
    '                ''
    '                mensaje &= valor & "|"
    '                ''
    '                ''
    '                PonInf()
    '            Next
    '            ''
    '            mensaje &= vbCrLf
    '            ''
    '            PonInf()
    '            uf.cFam = Nothing
    '        Next
    '        'MsgBox(mensaje)
    '        ''
    '        '' Escribimos el listado en un fichero temportal.
    '        Dim queFi As String = IO.Path.ChangeExtension(evRevit.evAppUI.ActiveUIDocument.Document.PathName, "_before.txt")
    '        Try
    '            IO.File.Delete(queFi)
    '        Catch ex As Exception
    '            ''
    '        End Try
    '        Try
    '            ' No escribiremos el fichero final, comentamos linea.
    '            ' IO.File.WriteAllText(queFi, mensaje)
    '        Catch ex As Exception
    '            '' El fichero no existía o estaba abierto.
    '        End Try
    '        mensaje = ""
    '        ''
    '        ''
    '        '' **** LLENAMOS LOS DATOS DEL MARKET ELEGIDO, BUSCAMOS Y SI NO LO ENCONTRAMOS, SIGUIENTE MARKET
    '        '' **** Buscaremos en las delegaciones, en orden de arrM123(M1,M2,M3)
    '        '' Si no tiene FAMILY_CODE, buscamos por ITEM_CODE.
    '        '
    '        '' Inicializamos arrEncontrados
    '        Dim arrEncontrados As New ArrayList                 '' Array de los que vayamos encontrando. Para no procesarlos otra vez.
    '        ''
    '        PonInf(preInf & "Phase 3/4" & vbCrLf & "Searching for CODED Data", True, arrM.Length * oBody.NumberOfRows * oBody.NumberOfColumns)
    '        ''
    '        Dim errortemp As String = ""        '' Para errores temporales.
    '        Dim arrBorrar As New ArrayList          '' ArrayList de los objectos clsFilas a borrar (Si tienen errores)
    '        '
    '        For x As Integer = 0 To UBound(arrM)
    '            ' Si ya está todo procesado. Salimos
    '            If arrEncontrados.Count = oBody.NumberOfRows Then Exit For
    '            '
    '            If arrM(x) = "" OrElse IsNumeric(arrM(x)) = False Then
    '                PonInf(, , , oBody.NumberOfRows * oBody.NumberOfColumns)
    '                ''
    '                Continue For
    '            End If
    '            ''
    '            '' Codigo de la compañia a utilizar (String)
    '            Dim CodMarket As String = arrM(x)
    '            ''
    '            '' Llenamos los objetos del mercado elegido. Sólo si no hemos encontrado los
    '            '' datos que necesitamos. Si los hemos encontrado, salimos fuera.
    '            If IsNumeric(CodMarket) AndAlso LlenaDatosMercados(CInt(CodMarket)) <> "" Then
    '                PonInf(preInf & "Phase 3/4" & vbCrLf & "Searching for CODED Data" & vbCrLf &
    '                           "Error in Market " & CodMarket, , , oBody.NumberOfRows * oBody.NumberOfColumns)
    '                Continue For
    '            End If
    '            ''
    '            PonInf(preInf & "Phase 3/4" & vbCrLf & "Searching for CODED Data" & vbCrLf &
    '                           "Searching in Market " & CodMarket, , , oBody.NumberOfRows * oBody.NumberOfColumns)
    '            ''
    '            '' Datos de mercado para escribir
    '            Dim W_MARKET As String = "M" & x + 1    'arrM123(x) ' "M" & x
    '            'Dim M1 As String = utilesRevit.ParametroProyectoLee(oDoc, "PROJECT_CONFIG", "M1")   '' Código del mercado 1
    '            ''
    '            '' **** Ahora volvemos a recorrer todas las filas para buscar y escribir valores.
    '            ''
    '            '' Sacamos todos los datos a un fichero.
    '            For fi As Integer = oBody.FirstRowNumber To oBody.LastRowNumber
    '                ''
    '                '' Si ya está todo procesado. Salimos
    '                '' If arrEncontrados.Count = oBody.NumberOfRows Then Exit For
    '                '' Si ya hemos encontrado esa línea, pasamos a la siguiente.
    '                '' O si es una fila de cabecera, vacia o de totales.
    '                If arrEncontrados.Contains(fi) Then Continue For
    '                ''
    '                ''
    '                Dim valorcoluno As String = vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber)
    '                valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 1)
    '                valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 2)
    '                valorcoluno &= vs.GetCellText(SectionType.Body, fi, oBody.FirstColumnNumber + 3)
    '                '' Saltamos las filas vacias (la fila 2), la fila de cabeceras (1ª col = ITEM_CODE) y la ultima con los totales.
    '                If valorcoluno = "" OrElse valorcoluno.ToUpper.Contains("TOTAL") OrElse fi = oBody.FirstRowNumber Then
    '                    If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                    ''
    '                    PonInf()
    '                    Continue For
    '                End If
    '                ''
    '                '' Objecto ULMALGFree.clsFamilia de la fila
    '                uf.cFam = CType(colFilas(fi), ULMALGFree.clsFamilia)
    '                ''
    '                '' Cogemos los datos que necesitamos para buscar en la Bases de Datos.
    '                Dim TYPE As String = ""
    '                If colTYPE > -1 Then TYPE = vs.GetCellText(SectionType.Body, fi, colTYPE)
    '                Dim FAMILY As String = ""
    '                If colFamily > -1 Then FAMILY = vs.GetCellText(SectionType.Body, fi, colFamily)
    '                Dim FAMILY_CODE As String = ""
    '                If colFAMILY_CODE > -1 Then FAMILY_CODE = vs.GetCellText(SectionType.Body, fi, colFAMILY_CODE)
    '                Dim ITEM_CODE As String = ""
    '                If colITEM_CODE > -1 Then ITEM_CODE = vs.GetCellText(SectionType.Body, fi, colITEM_CODE)
    '                Dim ITEM_LENGTH As String = ""
    '                If colITEM_LENGTH > -1 Then ITEM_LENGTH = vs.GetCellText(SectionType.Body, fi, colITEM_LENGTH)
    '                Dim ITEM_WIDTH As String = ""
    '                If colITEM_WIDTH > -1 Then ITEM_WIDTH = vs.GetCellText(SectionType.Body, fi, colITEM_WIDTH)
    '                Dim ITEM_HEIGHT As String = ""
    '                If colITEM_HEIGHT > -1 Then ITEM_HEIGHT = vs.GetCellText(SectionType.Body, fi, colITEM_HEIGHT)
    '                Dim ITEM_WEIGHT As String = ""
    '                If colITEM_WEIGHT > -1 Then ITEM_WEIGHT = vs.GetCellText(SectionType.Body, fi, colITEM_WEIGHT)
    '                Dim Count As String = ""
    '                If colCount > -1 Then Count = vs.GetCellText(SectionType.Body, fi, colCount)
    '                Dim ITEM_GENERIC As String = ""
    '                If colITEM_GENERIC > -1 Then ITEM_GENERIC = vs.GetCellText(SectionType.Body, fi, colITEM_GENERIC)
    '                ''
    '                '' Poner ITEM_LENGTH, ITEM_WIDTH e ITEM_HEIGHT en mm.
    '                ITEM_LENGTH = Unidades_DameMilimetros(ITEM_LENGTH)
    '                ITEM_WIDTH = Unidades_DameMilimetros(ITEM_WIDTH)
    '                ITEM_HEIGHT = Unidades_DameMilimetros(ITEM_HEIGHT)
    '                ''
    '                Try
    '                    If FAMILY_CODE = "" AndAlso ITEM_CODE = "" Then                 '' CONERROR
    '                        If lstErrores.Contains(preInf & "(" & Count & ") FAMILY=" & FAMILY & ", TYPE = " & TYPE & ", FAMILY_CODE = '', ITEM_CODE = ''") = False Then
    '                            lstErrores.Add(preInf & "(" & Count & ") FAMILY=" & FAMILY & ", TYPE = " & TYPE & ", FAMILY_CODE = '', ITEM_CODE = ''")
    '                        End If
    '                        ''
    '                        '' Esto sería un error
    '                        errores &= "***** ERRORS OF " & evRevit.evAppUI.ActiveUIDocument.Document.Title & " *****" & vbCrLf & vbCrLf
    '                        errores &= "Row " & fi & " error (TYPE=" & TYPE & " / Count=" & Count & ")" & vbCrLf
    '                        errores &= vbTab & "FAMILY = " & FAMILY & " FAMILY_CODE = '' And ITEM_CODE = ''" & vbCrLf & vbCrLf
    '                        'errores &= ParametrosEscribreEnFamiliasDocumento(oDoc, arrCriterios, colCV)
    '                        ''
    '                        '' Ponemos los datos en la clase (criterios y propiedades a escribir)
    '                        '' Aunque deberíamos borrarlo de la colección colFilas, para no procesarlo
    '                        'uf.cFam.PonCriterios(arrCriterios)
    '                        'uf.cFam.PonPropEscribir(arrPropEscribe)
    '                        'uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.CONERROR)
    '                        ''
    '                        '' Añadimos esta fila al ArrayList de filas a borrar (arrBorrar) después de iterar con todas las filas.
    '                        arrBorrar.Add(fi)
    '                        If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                        ''
    '                    ElseIf FAMILY_CODE = "" AndAlso ITEM_CODE <> "" Then            '' ESTATICO
    '                        '' Solo procesamos los estáticos si todo = true
    '                        Dim oArt As ULMALGFree.clsArticulos = Nothing
    '                        If todo = True Then
    '                            '' Buscamos por ITEM_CODE (Producto Generico)
    '                            If uf.colArticulos.ContainsKey(ITEM_CODE) Then
    '                                oArt = CType(uf.colArticulos(ITEM_CODE), ULMALGFree.clsArticulos)
    '                                Dim descripcion As String = ""
    '                                ''
    '                                For y As Integer = 0 To UBound(arrL)
    '                                    If arrL(y) = "" Then Continue For
    '                                    'Dim idioma As String = colIdiomas(arrL123(y)).ToString
    '                                    'descripcion = DameDescripcionIdioma(idioma, articulos(ITEM_CODE))
    '                                    descripcion = oArt.colDescritions(arrL(y)).ToString
    '                                    ''
    '                                    If descripcion <> "" Then
    '                                        Exit For
    '                                    End If
    '                                Next
    '                                ''
    '                                '' Volvemos a buscar en idiomadef, si no lo hemos encontrado en ninguno de los idiomas seleccionados.
    '                                If descripcion = "" Then
    '                                    '' descripcion en el idioma por defecto "en"
    '                                    'descripcion = DameDescripcionIdioma(colIdiomas(idiomadef).ToString, articulos(ITEM_CODE))
    '                                    descripcion = oArt.colDescritions(DEFAULT_PROGRAM_LANGUAGE).ToString
    '                                    If descripcion <> "" Then
    '                                        errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE & "--> Description not found in languages. Default language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                    Else
    '                                        descripcion = oArt.colDescritions("local").ToString
    '                                        errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE & "--> Description not found in languages. local language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                    End If
    '                                End If
    '                                ''
    '                                Try
    '                                    ' *** Poner los parametros
    '                                    '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                                    uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.ESTATICO)
    '                                    ''
    '                                    '' Ponemos los datos en la clase (criterios y propiedades a escribir)
    '                                    uf.cFam.ITEM_GENERIC_UNIT = oArt.genericArticleUnit
    '                                    uf.cFam.colCV("ITEM_CODE_L") = oArt.articleCodeL
    '                                    uf.cFam.colCV("ITEM_WEIGHT_L") = oArt.weightL
    '                                    uf.cFam.colCV("ITEM_MARKET") = ITEM_MARKET
    '                                    'uf.cFam.PonDato(colTYPE, TYPE, True)
    '                                    'uf.cFam.PonDato(colFAMILY_CODE, FAMILY_CODE, True)
    '                                    'uf.cFam.PonDato(colITEM_CODE, ITEM_CODE, True)
    '                                    'uf.cFam.PonDato(colITEM_LENGTH, ITEM_LENGTH, True)
    '                                    'uf.cFam.PonDato(colITEM_WIDTH, ITEM_WIDTH, True)
    '                                    'uf.cFam.PonDato(colITEM_HEIGHT, ITEM_HEIGHT, True)
    '                                    uf.cFam.PonDato(colITEM_WEIGHT, oArt.weightEnd_Dame(False), True)    ' oArt.weightAll, True)
    '                                    'uf.cFam.colCV("WEIGHT") = oArt.weightEnd_Dame(True)
    '                                    'uf.cFam.colCV("ITEM_WEIGHT") = oArt.weightAll
    '                                    '
    '                                    uf.cFam.PonDato(colITEM_DESCRIPTION, descripcion, True)
    '                                    'uf.cFam.PonDato(colITEM_CODEL, itemcodel, True)
    '                                    uf.cFam.PonDato(colCount, Count, True)
    '                                    uf.cFam.PonDato(colW_MARKET, W_MARKET, True)
    '                                    uf.cFam.PonDato(colITEM_GENERIC, "", True)     'ITEM_GENERIC = 0 (o "")
    '                                    ''
    '                                    uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.ENCONTRADO
    '                                    uf.cFam.fila = fi
    '                                    '' Calculamos el PESOFIN del artículo
    '                                    uf.cFam.PonPesoFin()
    '                                    uf.cFam.PonTotalesGenerales()
    '                                    ''
    '                                    If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                                    ''
    '                                    '' Si había dado error en un Mercado y se encuentra en otro.
    '                                    If errortemp.Contains(
    '                                        ITEM_CODE & "--> Not in XML of Articles..." & vbCrLf & vbCrLf) Then
    '                                        errortemp = errortemp.Replace(
    '                                             ITEM_CODE & "--> Not in XML of Articles..." & vbCrLf & vbCrLf, "")
    '                                    End If
    '                                Catch ex As Exception
    '                                    errores &= "LeeEscribeDatos --> " & "FAMILY = " & FAMILY & " FAMILY_CODE = '' And ITEM_CODE <> '' --> colArticulos.ContainsKey(ITEM_CODE)" & vbCrLf _
    '                                        & ex.Source & vbCrLf & ex.Message & vbCrLf
    '                                    '
    '                                    If lstErrores.Contains(preInf & "(" & Count & ")  FAMILY=" & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE) = False Then
    '                                        lstErrores.Add(preInf & "(" & Count & ")  FAMILY=" & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE)
    '                                    End If
    '                                End Try
    '                            Else
    '                                uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.ESTATICO)
    '                                uf.cFam.PonDato(colITEM_WEIGHT, "", True)
    '                                uf.cFam.PonDato(colITEM_DESCRIPTION, "", True)
    '                                uf.cFam.PonDato(colW_MARKET, "", True)
    '                                uf.cFam.colCV("ITEM_MARKET") = 0
    '                                uf.cFam.PonDato(colITEM_GENERIC, "", True)
    '                                uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.NOENCONTRADO
    '                                uf.cFam.fila = fi
    '                                '' No encontrado en Artículos del mercado especificado
    '                                errortemp &= ITEM_CODE & "--> Not in XML of Articles..." & vbCrLf & vbCrLf
    '                                '
    '                                If lstErrores.Contains(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE & " --> Not in XML of Articles...") = False Then
    '                                    lstErrores.Add(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE & " --> Not in XML of Articles...")
    '                                End If
    '                            End If
    '                        ElseIf todo = False Then
    '                            '' Lo quitaremos de la lista de filas a procesar.
    '                            arrBorrar.Add(fi)
    '                            ''Y lo damos por procesado, para que no vuelva a  buscarlo en otra delegación.
    '                            If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                        End If
    '                        oArt = Nothing
    '                    ElseIf FAMILY_CODE <> "" AndAlso (ITEM_LENGTH <> "" OrElse ITEM_WIDTH <> "" OrElse ITEM_HEIGHT <> "") Then          '' DINAMICO (Tiene alguna medida en ITEM_LENGTH o ITEM_WIDTH o ITEM_HEIGHT)
    '                        '' Buscamos por FAMILY_CODE para localizar ITEM_CODE y otros datos (Producto Especial)
    '                        '' Habrá que buscar el grupo y las medidas para saber el ITEM_CODE correcto
    '                        ''
    '                        '' ***** Cadena para buscar en fam_cod (key=cadenabuscar, value=clsFamCode
    '                        Dim queFam As ULMALGFree.clsFamCode = CadenaBuscardynamicFamilies(FAMILY_CODE, ITEM_LENGTH, ITEM_WIDTH, ITEM_HEIGHT)
    '                        Dim ITEM_CODE_buscar As String = ""
    '                        Dim oArt As ULMALGFree.clsArticulos = Nothing
    '                        ''**************************
    '                        ''
    '                        If queFam IsNot Nothing Then
    '                            ITEM_CODE_buscar = queFam.ICODE
    '                            '' ***** Ponermos otros datos leidos de XXX_revit_dynamiuf.cFamilies.xml
    '                            ITEM_GENERIC = queFam.IGENERIC
    '                        Else
    '                            ''
    '                            '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                            uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.DINAMICO)
    '                            ''
    '                            '' No encontrado en fam_code.txt. Habrá que borrar los datos que tenga.
    '                            '' Ponemos los datos en la clase (criterios y propiedades a escribir)
    '                            uf.cFam.PonDato(colITEM_CODE, "", True)                  '' Borrar Dato
    '                            Try
    '                                uf.cFam.PonDato(colITEM_WEIGHT, "", True)                  '' Borrar Dato
    '                            Catch ex As Exception
    '                                Debug.Print("")
    '                            End Try
    '                            uf.cFam.PonDato(colITEM_DESCRIPTION, "", True)                  '' Borrar Dato
    '                            uf.cFam.PonDato(colW_MARKET, "", True)                  '' Borrar Dato
    '                            uf.cFam.colCV("ITEM_MARKET") = 0
    '                            uf.cFam.PonDato(colITEM_GENERIC, "", True)   '' O ""
    '                            ''
    '                            uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.NOENCONTRADO
    '                            uf.cFam.fila = fi
    '                            ''
    '                            'errortemp = errortemp.Replace(FAMILY_CODE & "--> Not in fam_code.txt" & vbCrLf & vbCrLf, "")
    '                            errortemp &= FAMILY_CODE & "--> Not in fam_code.txt" & vbCrLf & vbCrLf
    '                            If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                            'If arrBorrar.Contains(fi) = False Then arrBorrar.Add(fi)
    '                            '' Pasamos a la siguiente fila. Este ya no se procesará más veces.
    '                            Continue For
    '                        End If
    '                        ''
    '                        ''
    '                        ''
    '                        '' Ahora localizamos en Artículos el ITEM_CODE y sacamos los datos que necesitamos.
    '                        'If colArticulos.ContainsKey(ITEM_CODE_buscar) AndAlso CType(colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos).active = True Then
    '                        If uf.colArticulos.ContainsKey(ITEM_CODE_buscar) Then
    '                            oArt = CType(uf.colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos)
    '                            Dim descripcion As String = ""
    '                            ''
    '                            For y As Integer = 0 To UBound(arrL)
    '                                If arrL(y) = "" Then Continue For
    '                                'Dim idioma As String = colIdiomas(arrL123(y)).ToString
    '                                'descripcion = DameDescripcionIdioma(idioma, articulos(ITEM_CODE_buscar))
    '                                descripcion = oArt.colDescritions(arrL(y)).ToString
    '                                ''
    '                                If descripcion <> "" Then
    '                                    Exit For
    '                                End If
    '                            Next
    '                            '' Volvemos a buscar en idiomadef, si no lo hemos encontrado en ninguno de los idiomas seleccionados.
    '                            If descripcion = "" Then
    '                                '' descripcion en el idioma por defecto "en"
    '                                'descripcion = DameDescripcionIdioma(colIdiomas(idiomadef).ToString, articulos(ITEM_CODE_buscar))
    '                                descripcion = oArt.colDescritions(DEFAULT_PROGRAM_LANGUAGE).ToString
    '                                If descripcion <> "" Then
    '                                    errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE_buscar & "--> Description not found in languages. Default language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                Else
    '                                    descripcion = oArt.colDescritions("local").ToString
    '                                    errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE_buscar & "--> Description not found in languages. Local language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                End If
    '                            End If
    '                            ''
    '                            Try
    '                                '
    '                                '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                                uf.cFam.ITEM_GENERIC_UNIT = oArt.genericArticleUnit
    '                                uf.cFam.ITEM_WEIGHT_UNIT = oArt.weightUnit
    '                                uf.cFam.colCV("ITEM_CODE_L") = oArt.articleCodeL
    '                                uf.cFam.colCV("ITEM_WEIGHT_L") = oArt.weightL
    '                                uf.cFam.colCV("ITEM_MARKET") = ITEM_MARKET
    '                                '
    '                                uf.cFam.colCV("AREA") = oArt.formArea
    '                                uf.cFam.colCV("AREA_UNID") = oArt.formAreaUnit
    '                                '
    '                                uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.DINAMICO)
    '                                '
    '                                uf.cFam.PonDato(colITEM_CODE, ITEM_CODE_buscar, True)
    '                                'uf.cFam.PonDato(colITEM_CODEL, itemcodel, True)
    '                                uf.cFam.PonDato(colITEM_WEIGHT, oArt.weightEnd_Dame(False), True)    ' oArt.weightAll, True)
    '                                'uf.cFam.colCV("WEIGHT") = oArt.weightEnd_Dame(True)
    '                                uf.cFam.PonDato(colITEM_DESCRIPTION, descripcion, True)
    '                                uf.cFam.PonDato(colCount, Count, True)
    '                                uf.cFam.PonDato(colW_MARKET, IIf(queFam.W_MARKET <> W_MARKET, queFam.W_MARKET, W_MARKET).ToString, True)
    '                                uf.cFam.PonDato(colITEM_GENERIC, ITEM_GENERIC, True)
    '                                ''
    '                                uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.ENCONTRADO
    '                                uf.cFam.fila = fi
    '                                '' Calculamos el PESOFIN del artículo
    '                                uf.cFam.PonPesoFin()
    '                                uf.cFam.PonTotalesGenerales()
    '                                ''
    '                                If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                                ''
    '                                '' Si había dado error en un Mercado y se encuentra en otro.
    '                                errortemp = errortemp.Replace(ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf, "")
    '                            Catch ex As Exception
    '                                errores &= "LeeEscribeDatos --> " & "FAMILY = " & FAMILY & " FAMILY_CODE <> '' And (ITEM_LENGTH <> '' OrElse ITEM_WIDTH <> '' OrElse ITEM_HEIGHT <> '' --> colArticulos.ContainsKey(ITEM_CODE_buscar)" & vbCrLf _
    '                                    & ex.Source & vbCrLf & ex.Message & vbCrLf
    '                                '
    '                                If lstErrores.Contains(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE_buscar & " --> ERROR...") = False Then
    '                                    lstErrores.Add(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE_buscar & " --> ERROR...")
    '                                End If
    '                            End Try
    '                        Else    'If colArticulos.ContainsKey(ITEM_CODE_buscar) = False Then
    '                            '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                            uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.DINAMICO)
    '                            uf.cFam.PonDato(colITEM_CODE, "", True)
    '                            Try
    '                                uf.cFam.PonDato(colITEM_WEIGHT, "", True)
    '                            Catch ex As Exception
    '                                Debug.Print("LeeEscribeDatos / PonDato --> " & ex.Message)
    '                            End Try
    '                            uf.cFam.PonDato(colITEM_DESCRIPTION, "", True)
    '                            uf.cFam.PonDato(colW_MARKET, "", True)
    '                            uf.cFam.colCV("ITEM_MARKET") = ""
    '                            uf.cFam.PonDato(colITEM_GENERIC, ITEM_GENERIC, True)
    '                            ''
    '                            uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.NOENCONTRADO
    '                            uf.cFam.fila = fi
    '                            '' No encontrado en Artículos del mercado especificado
    '                            errortemp = errortemp.Replace(ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf, "")
    '                            errortemp &= ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf
    '                            '
    '                            If lstErrores.Contains(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE_buscar & " --> Not in XML of Articles...") = False Then
    '                                lstErrores.Add(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE_buscar & " --> Not in XML of Articles...")
    '                            End If
    '                        End If
    '                        ''
    '                    ElseIf FAMILY_CODE <> "" AndAlso ITEM_LENGTH = "" AndAlso ITEM_WIDTH = "" AndAlso ITEM_HEIGHT = "" Then          '' GENERICO (No tiene ninguna medida en ITEM_LENGTH o ITEM_WIDTH o ITEM_HEIGHT)
    '                        '' Buscamos por FAMILY_CODE para localizar ITEM_CODE y otros datos (Producto Especial)
    '                        '' Habrá que buscar el grupo y las medidas para saber el ITEM_CODE correcto
    '                        ''
    '                        '' ***** Cadena para buscar en fam_cod (key=cadenabuscar, value=clsFamCode
    '                        Dim queFam As ULMALGFree.clsFamCode = CadenaBuscardynamicFamilies(FAMILY_CODE, ITEM_LENGTH, ITEM_WIDTH, ITEM_HEIGHT)
    '                        Dim ITEM_CODE_buscar As String = ""
    '                        Dim oArt As ULMALGFree.clsArticulos = Nothing
    '                        ''**************************
    '                        ''
    '                        If queFam IsNot Nothing Then
    '                            ITEM_CODE_buscar = queFam.ICODE
    '                            '' ***** Ponermos otros datos leidos de dynamicsfamilicode.txt
    '                            ITEM_GENERIC = queFam.IGENERIC
    '                        Else
    '                            ''
    '                            '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                            uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.GENERICO)
    '                            ''
    '                            '' No encontrado en fam_code.txt. Habrá que borrar los datos que tenga.
    '                            '' Ponemos los datos en la clase (criterios y propiedades a escribir)
    '                            uf.cFam.PonDato(colITEM_CODE, "", True)                  '' Borrar Dato
    '                            uf.cFam.PonDato(colITEM_WEIGHT, "", True)                  '' Borrar Dato
    '                            uf.cFam.PonDato(colITEM_DESCRIPTION, "", True)                  '' Borrar Dato
    '                            uf.cFam.PonDato(colW_MARKET, "", True)                  '' Borrar Dato
    '                            uf.cFam.colCV("ITEM_MARKET") = ""
    '                            uf.cFam.PonDato(colITEM_GENERIC, "", True)
    '                            ''
    '                            uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.NOENCONTRADO
    '                            uf.cFam.fila = fi
    '                            ''
    '                            errortemp &= FAMILY_CODE & "--> Not in fam_code.txt" & vbCrLf & vbCrLf
    '                            If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                            'If arrBorrar.Contains(fi) = False Then arrBorrar.Add(fi)
    '                            '' Pasamos a la siguiente fila. Este ya no se procesará más veces.
    '                            Continue For
    '                        End If
    '                        ''
    '                        ''
    '                        ''
    '                        '' Ahora localizamos en Artículos el ITEM_CODE y sacamos los datos que necesitamos.
    '                        'If colArticulos.ContainsKey(ITEM_CODE_buscar) AndAlso CType(colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos).active = True Then
    '                        If uf.colArticulos.ContainsKey(ITEM_CODE_buscar) Then
    '                            oArt = CType(uf.colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos)
    '                            Dim descripcion As String = ""
    '                            ''
    '                            For y As Integer = 0 To UBound(arrL)
    '                                If arrL(y) = "" Then Continue For
    '                                'Dim idioma As String = colIdiomas(arrL123(y)).ToString
    '                                'descripcion = DameDescripcionIdioma(idioma, articulos(ITEM_CODE_buscar))
    '                                descripcion = oArt.colDescritions(arrL(y)).ToString
    '                                ''
    '                                If descripcion <> "" Then
    '                                    Exit For
    '                                End If
    '                            Next
    '                            '' Volvemos a buscar en idiomadef, si no lo hemos encontrado en ninguno de los idiomas seleccionados.
    '                            If descripcion = "" Then
    '                                '' descripcion en el idioma por defecto "en"
    '                                'descripcion = DameDescripcionIdioma(colIdiomas(idiomadef).ToString, articulos(ITEM_CODE_buscar))
    '                                descripcion = oArt.colDescritions(DEFAULT_PROGRAM_LANGUAGE).ToString
    '                                If descripcion <> "" Then
    '                                    errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE_buscar & "--> Description not found in languages. Default language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                Else
    '                                    descripcion = oArt.colDescritions("local").ToString
    '                                    errores &= "Line " & fi & " / ITEM_CODE: " & ITEM_CODE_buscar & "--> Description not found in languages. Local language used '" & DEFAULT_PROGRAM_LANGUAGE & "'" & vbCrLf & vbCrLf
    '                                End If
    '                            End If
    '                            ''
    '                            Try
    '                                'Dim peso As String = CType(colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos).peso
    '                                ''
    '                                '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                                uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.GENERICO)
    '                                ''
    '                                '' Ponemos los datos en la clase (criterios y propiedades a escribir)
    '                                uf.cFam.ITEM_GENERIC_UNIT = oArt.genericArticleUnit
    '                                uf.cFam.ITEM_WEIGHT_UNIT = oArt.weightUnit
    '                                uf.cFam.colCV("ITEM_CODE_L") = oArt.articleCodeL
    '                                uf.cFam.colCV("ITEM_WEIGHT_L") = oArt.weightL
    '                                uf.cFam.colCV("ITEM_MARKET") = ITEM_MARKET
    '                                '
    '                                uf.cFam.colCV("AREA") = CType(uf.colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos).formArea
    '                                uf.cFam.colCV("AREA_UNID") = CType(uf.colArticulos(ITEM_CODE_buscar), ULMALGFree.clsArticulos).formAreaUnit

    '                                uf.cFam.PonDato(colITEM_CODE, ITEM_CODE, True)
    '                                uf.cFam.PonDato(colITEM_WEIGHT, oArt.weightEnd_Dame(False), True)    ' oArt.weightAll, True)
    '                                uf.cFam.PonDato(colITEM_DESCRIPTION, descripcion, True)
    '                                uf.cFam.PonDato(colCount, Count, True)
    '                                uf.cFam.PonDato(colW_MARKET, IIf(queFam.W_MARKET <> W_MARKET, queFam.W_MARKET, W_MARKET).ToString, True)
    '                                uf.cFam.PonDato(colITEM_GENERIC, ITEM_GENERIC, True)
    '                                ''
    '                                uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.ENCONTRADO
    '                                uf.cFam.fila = fi
    '                                ''
    '                                uf.cFam.PonPesoFin()
    '                                uf.cFam.PonTotalesGenerales()
    '                                ''
    '                                If arrEncontrados.Contains(fi) = False Then arrEncontrados.Add(fi)
    '                                ''
    '                                '' Si había dado error en un Mercado y se encuentra en otro.
    '                                errortemp = errortemp.Replace(ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf, "")
    '                            Catch ex As Exception
    '                                errores &= "LeeEscribeDatos --> " & "FAMILY = " & FAMILY & " FAMILY_CODE <> '' And (ITEM_LENGTH = '' OrElse ITEM_WIDTH = '' OrElse ITEM_HEIGHT = '' --> colArticulos.ContainsKey(ITEM_CODE_buscar)" & vbCrLf _
    '                                                                        & ex.Source & vbCrLf & ex.Message & vbCrLf

    '                                '
    '                                If lstErrores.Contains(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & " --> ERROR...") = False Then
    '                                    lstErrores.Add(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & " --> ERROR...")
    '                                End If
    '                            End Try
    '                        Else    'If colArticulos.ContainsKey(ITEM_CODE_buscar) = False Then
    '                            '' Esta es fundamental. Ya que rellena arrCriterios y arrPropEscribir en la clase ULMALGFree.clsFamilia (uf.cFam) según Tipo.
    '                            uf.cFam.PonTipo(ULMALGFree.clsFamilia.Tipo.GENERICO)
    '                            uf.cFam.PonDato(colITEM_CODE, "", True)
    '                            uf.cFam.PonDato(colITEM_WEIGHT, "", True)
    '                            uf.cFam.PonDato(colITEM_DESCRIPTION, "", True)
    '                            uf.cFam.PonDato(colW_MARKET, "", True)
    '                            uf.cFam.colCV("ITEM_MARKET") = ""
    '                            uf.cFam.PonDato(colITEM_GENERIC, ITEM_GENERIC, True)
    '                            ''
    '                            uf.cFam.QUEESTADO = ULMALGFree.clsFamilia.ESTADO.NOENCONTRADO
    '                            uf.cFam.fila = fi
    '                            '' No encontrado en Artículos del mercado especificado
    '                            errortemp = errortemp.Replace(ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf, "")
    '                            errortemp &= ITEM_CODE_buscar & "--> Not in XML of Articles..." & vbCrLf & vbCrLf
    '                            '
    '                            If lstErrores.Contains(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE & " --> Not in XML of Articles...") = False Then
    '                                lstErrores.Add(preInf & "(" & Count & ") FAMILY = " & FAMILY & ", TYPE = " & TYPE & ", ITEM_CODE = " & ITEM_CODE & " --> Not in XML of Articles...")
    '                            End If
    '                        End If
    '                        ''
    '                    End If
    '                Catch ex As Exception
    '                    errores &= "LeeEscribeDatos --> " & ex.Source & vbCrLf & ex.Message & vbCrLf
    '                End Try
    '                ''
    '                uf.cFam = Nothing
    '                PonInf()
    '                ''
    '                System.Windows.Forms.Application.DoEvents()
    '                '' FIN DE FILAS BOM
    '            Next
    '            ''
    '            PonInf()
    '            '' FIN DE MERCADOS
    '        Next
    '        ''
    '        '' ***** Borramos de colFilas los que hayan dado error. Para no procesarlos (Buscar, rellenar parámetros, etc.)
    '        For Each fi As Integer In arrBorrar
    '            colFilas.Remove(fi)
    '        Next
    '        ''
    '        ''
    '        PonInf(preInf & "Phase 4/4" & vbCrLf & "SEARCHING and WRITING properties in Families", True, colFilas.Count)
    '        ''
    '        errores &= ParametrosEscribeEnFamiliasDocumento(evRevit.evAppUI.ActiveUIDocument.Document, colFilas)
    '        ''
    '        '' Volvemos a filtrar, por si hay varios W_MARKET
    '        ''
    '        '' Campos de agrupación. No usar ninguno de los que vayamos a escribir.
    '        ' Dim agruparpor1 As String() = New String() { _
    '        'nType.ToUpper, "FAMILY_CODE", "FILTER_ID", "W_MARKET"}
    '        '' Crear y agrupar la tabla bomBase (#BOM-WK)
    '        'ScheduleAgrupaFields(vs.Document, vs, agruparpor1)
    '        ''
    '        errores &= errortemp
    '        ''
    '        ''
    '        '' Abrimos el fichero con #BOM-WK anterior (Para comparar)
    '        Try
    '            'If log = True And IO.File.Exists(queFi) Then Call Process.Start(queFi)
    '        Catch ex As Exception
    '            ''
    '        End Try

    'FINAL:
    '        ''
    '        'Dim queFi1 As String = IO.Path.Combine(IO.Path.ChangeExtension(oDoc.PathName, "_error.txt"))
    '        Dim nFilog As String = IO.Path.ChangeExtension(evRevit.evAppUI.ActiveUIDocument.Document.PathName, "_error.txt")     '' Path completo, si lo dejamos en la misma carpeta que el .RVT
    '        '' Como los dejamos en el directorio LOGS, componer aquí el camino completo a LOGS + solo nombre.
    '        Dim queFi1 As String = IO.Path.Combine(ULMALGFree.clsBase._appLogBaseFolder, IO.Path.GetFileName(nFilog))
    '        '' Si hay errores, escribimos log de errores. Si no hay y existe el fichero, lo borramos.
    '        Try
    '            If IO.File.Exists(queFi1) Then IO.File.Delete(queFi1)
    '        Catch ex As Exception
    '            ''
    '        End Try
    '        Try
    '            If IO.Directory.Exists(ULMALGFree.clsBase._appLogBaseFolder) = False Then IO.Directory.CreateDirectory(ULMALGFree.clsBase._appLogBaseFolder)
    '            'If IO.Directory.Exists(_dirApplogs) Then PermisosTodoCarpeta(_dirApplogs)
    '        Catch ex As Exception
    '            '' No hacemos nada. No existe el directorio.
    '        End Try
    '        '
    '        Try
    '            If errores <> "" Then
    '                ' No generamos el fichero log. Comentar
    '                'IO.File.WriteAllText(queFi1, errores)
    '                'If IO.File.Exists(queFi1) And log = True Then Call Process.Start(queFi1)
    '            End If
    '        Catch ex As Exception
    '            '' El fichero no existía o estaba abierto.
    '        End Try
    '    End Sub
    '
    Public Sub PonInf(Optional queInf As String = "", Optional valorCero As Boolean = False, Optional maximo As Integer = 0, Optional suma As Integer = 1)
        'If frmC Is Nothing Then Exit Sub
        'If frmC IsNot Nothing AndAlso frmC.Visible = False Then Exit Sub
        '''
        'If queInf <> "" Then
        '    frmC.lblInformacion.Text = queInf
        '    frmC.lblInformacion.Refresh()
        'End If
        '''
        'If maximo > 0 Then frmC.pb1.Maximum = maximo
        'If valorCero = True Then
        '    frmC.pb1.Value = 0
        '    frmC.pb1.Refresh()
        'Else
        '    If frmC.pb1.Value + suma <= frmC.pb1.Maximum Then frmC.pb1.Value += suma
        '    frmC.pb1.Refresh()
        'End If
    End Sub
    ''
    'Public Function ParametrosEscribeEnFamiliasDocumento(ByRef queDoc As Autodesk.Revit.DB.Document,
    '                                                 queColFi As Hashtable) As String
    '    ''
    '    Dim errores As String = ""
    '    Dim contador As Integer = 0
    '    ''
    '    'Dim dMode As WorksharingDisplayMode = oDoc.ActiveView.GetWorksharingDisplayMode
    '    'oDoc.ActiveView.SetWorksharingDisplayMode(WorksharingDisplayMode.Off)
    '    ''
    '    Using trans1 As Autodesk.Revit.DB.Transaction = New Autodesk.Revit.DB.Transaction(queDoc, "Write family parameters")
    '        trans1.Start()
    '        ''
    '        Try
    '            modULMA.PonInf("", True, queColFi.Values.Count)
    '            Dim total As Integer = queColFi.Values.Count
    '            Dim FamCambiadas As New ArrayList      '' ArrayList de ID de FamiliInstances ya cambiadas.
    '            ''
    '            For Each queFaBusco As ULMALGFree.clsFamilia In queColFi.Values
    '                ''
    '                '' Crearemos un IList(Of FamilyInstance) de cada queFaBusco.colCV("TYPE" (nType.ToUpper)).ToString
    '                ''
    '                '' Tendremos que sacar las propiedades de la familia de arrCriterios
    '                '' para compararlas con los valores de ULMALGFree.clsFamilia. Si son iguales
    '                '' hemos localizado la familia a cambiar y pondremos sus valores.
    '                '' ** Concatenar los valores para compararlos.
    '                ''
    '                contador += 1
    '                ''
    '                Dim nombreFamilia As String = queFaBusco.colCV("TYPE").ToString
    '                ''
    '                PonInf("Phase 4/4" & vbCrLf & "SEARCHING and WRITING properties in Families" & vbCrLf & vbCrLf &
    '                           "(" & contador & " of " & total & ") " & "Family=" & nombreFamilia)
    '                ''
    '                '' Si queFaBusco tiene modificado = false. No tiene modificaciones de los datos
    '                '' por lo que no necesitamos escribir los mismos datos.
    '                If queFaBusco.modificado = False Then
    '                    PonInf()
    '                    Continue For
    '                End If
    '                ''
    '                '' **** Crear el IList si sólo queremos recorrer objetos (Sin modificarlos)
    '                '' **** Crear el ICollection si queremos recorrer los objetos (Modificándolos)
    '                'Dim colF As IList(Of Element) = _
    '                'utilesRevit.FamilyInstanceName_DameICollection(oDoc, nombreFamilia)
    '                Dim colF As ICollection(Of Element) =
    '                    utilesRevit.FamilyInstanceName_DameICollection(evRevit.evAppUI.ActiveUIDocument.Document, nombreFamilia)
    '                ''
    '                '' Si no hay nada en IList, continuar
    '                If colF Is Nothing OrElse (colF IsNot Nothing AndAlso colF.Count = 0) Then
    '                    PonInf()
    '                    Continue For
    '                End If
    '                ''
    '                Dim oFamEncontrada As FamilyInstance = Nothing
    '                ''
    '                '' ***** Iterar con cada elemento del IList
    '                For Each oFamI As FamilyInstance In colF
    '                    '' Si ya ha sido cambiada esta familia, continuar
    '                    If FamCambiadas.Contains(oFamI.Id) Then
    '                        Continue For
    '                    End If
    '                    'Dim oFamI As FamilyInstance = CType(oDoc.GetElement(oFamIa.Id), FamilyInstance)
    '                    ''
    '                    ''
    '                    Dim comparaCla As String = ""       '' Cadena de comparación en ULMALGFree.clsFamilia
    '                    Dim comparaFam As String = ""       '' Cadena de comparación en FamilyInstance
    '                    '' ***** Buscamos si cumple los criterios.
    '                    For Each criterio As String In queFaBusco.arrCriterios
    '                        ''
    '                        '' Concatenamos criterio de ULMALGFree.clsFamilia
    '                        comparaCla &= queFaBusco.colCV(criterio).ToString
    '                        Dim queValor As String = ""
    '                        '' PRUEBAS: Leer parametro de utilesRevit
    '                        queValor = utilesRevit.ParametroFamilyInstanceLee(queDoc, oFamI, criterio)
    '                        ''
    '                        If queValor <> "" Then

    '                            '' Comprobar las unidades
    '                            If (criterio = "ITEM_LENGTH" OrElse
    '                                criterio = "ITEM_WIDTH" OrElse
    '                                criterio = "ITEM_HEIGHT") AndAlso
    '                                (queValor.Trim <> "" AndAlso IsNumeric(queValor.Split(" "c)(0).Trim)) = True Then
    '                                ''
    '                                queValor = queValor.Split(" "c)(0).Trim
    '                                If unidLDoc.DisplayUnits = DisplayUnitType.DUT_MILLIMETERS Then
    '                                    '' Ponemos en milímetros los centímetros leídos.
    '                                    queValor = CDbl(FormatNumber(queValor.Trim, 2)).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_CENTIMETERS Then
    '                                    '' Ponemos en milímetros los centímetros leídos.
    '                                    queValor = (CDbl(FormatNumber(queValor.Trim, 2)) * 10).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_DECIMETERS Then
    '                                    '' Ponemos en milímetros los decímetros leídos.
    '                                    queValor = (CDbl(FormatNumber(queValor.Trim, 2)) * 100).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_METERS Then
    '                                    '' Ponemos en milímetros los metros leídos.
    '                                    queValor = (CDbl(FormatNumber(queValor.Trim, 2)) * 1000).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_DECIMAL_INCHES Then
    '                                    '' Ponemos en milímetros las pulgadas decimales leídas.
    '                                    queValor = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_DECIMAL_INCHES, DisplayUnitType.DUT_MILLIMETERS).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_DECIMAL_FEET Then
    '                                    '' Ponemos en milímetros los pies decimales leídos.
    '                                    queValor = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_DECIMAL_FEET, DisplayUnitType.DUT_MILLIMETERS).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES Then
    '                                    '' Ponemos en milímetros los pies y pulgados decimales leídos.
    '                                    queValor = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES, DisplayUnitType.DUT_MILLIMETERS).ToString
    '                                ElseIf unidLDoc.DisplayUnits = DisplayUnitType.DUT_FRACTIONAL_INCHES Then
    '                                    '' Ponemos en milímetros las pulgadas fraccionales leídas.
    '                                    queValor = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_FRACTIONAL_INCHES, DisplayUnitType.DUT_MILLIMETERS).ToString
    '                                End If
    '                                '' Si las unidades de longitud son diferentes de milímetros, convierte a milimetros.
    '                                'If unidLDoc.DisplayUnits <> DisplayUnitType.DUT_MILLIMETERS Then
    '                                'queValor = UnitUtils.ConvertFromInternalUnits(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_MILLIMETERS).ToString
    '                                'End If
    '                                '' Ponemos los valores con 2 decimales.
    '                                queValor = FormatNumber(queValor, 2, TriState.True, TriState.False, TriState.False)
    '                            End If
    '                            ''
    '                            ''
    '                            comparaFam &= queValor
    '                        End If
    '                        ''
    '                        'System.Windows.Forms.Application.DoEvents()
    '                        '' si no coinciden los criterios desde el principio, salimos.
    '                        If comparaCla <> comparaFam Then
    '                            Exit For
    '                        End If
    '                    Next
    '                    ''
    '                    If comparaCla = comparaFam Then
    '                        oFamEncontrada = oFamI ' CType(oDoc.GetElement(oFamI.Id), FamilyInstance) ' oFamI
    '                        'Exit For
    '                    End If
    '                    'modULMA.PonInf()     '' Avanzamos queColFi
    '                    ''
    '                    'System.Windows.Forms.Application.DoEvents()
    '                    ''
    '                    '' Si la hemos encontrado. Escribimos los valores de queFa.colCV
    '                    '' Siempre deberíamos encontrarlo. Ya que hemos filtrado por TYPE de BOM-WK
    '                    '' Si no hemos encontrado la familia, pasamo a la siguiente.
    '                    If oFamEncontrada Is Nothing Then
    '                        PonInf()
    '                        Continue For
    '                    End If
    '                    ''
    '                    ''
    '                    For Each nPar As String In queFaBusco.arrPropEscribir    ' queFa.colCV.Keys
    '                        'Dim trans1 As Autodesk.Revit.DB.Transaction = New Autodesk.Revit.DB.Transaction(queDoc, "Write family parameters of " & oFam.Name)
    '                        Dim valor As String = ""
    '                        valor = queFaBusco.colCV(nPar).ToString()
    '                        ' 2018/11/09 Cambiamos para que sólo ponga en ITEM_WEIGHT el peso unitario. Antes ponía el total.
    '                        'If nPar = "ITEM_WEIGHT" Then
    '                        '    valor = queFaBusco.colCV("PESOFIN").ToString()
    '                        'Else
    '                        'End If
    '                        '' FAMILYINSTANCE
    '                        Try
    '                            ' Coger el primer parámetros que sea de lectura/escritura
    '                            ' ITEM_CODE_L
    '                            Dim oPar As Parameter = Nothing
    '                            Dim oParS As Parameter = Nothing
    '                            'If nPar = "ITEM_CODE_L" Then
    '                            '    ' FAMILYSYMBOL (Por si no la hemos encontrado en en FamilyInsance) ** Solo estaticos
    '                            '    For Each oPar In oFamEncontrada.Symbol.GetParameters(nPar)
    '                            '        If Not (oPar.IsReadOnly) Then
    '                            '            Exit For
    '                            '        End If
    '                            '    Next
    '                            'Else
    '                            For Each oPar In oFamEncontrada.GetParameters(nPar)
    '                                If Not (oPar.IsReadOnly) Then
    '                                    Exit For
    '                                End If
    '                            Next

    '                            For Each oParS In oFamEncontrada.Symbol.GetParameters(nPar)
    '                                If Not (oParS.IsReadOnly) Then
    '                                    Exit For
    '                                End If
    '                            Next
    '                            'End If
    '                            ' ITEM_WEIGHT_L
    '                            'If nPar = "ITEM_WEIGHT_L" Then
    '                            '    ' FAMILYSYMBOL (Por si no la hemos encontrado en en FamilyInsance) ** Solo estaticos
    '                            '    For Each oPar In oFamEncontrada.Symbol.GetParameters(nPar)
    '                            '        If Not (oPar.IsReadOnly) Then
    '                            '            Exit For
    '                            '        End If
    '                            '    Next
    '                            'Else
    '                            '    For Each oPar In oFamEncontrada.GetParameters(nPar)
    '                            '        If Not (oPar.IsReadOnly) Then
    '                            '            Exit For
    '                            '        End If
    '                            '    Next
    '                            'End If

    '                            ' '' FAMILY (Por si no la hemos encontrado en en FamilySymbol) ** Solo estaticos
    '                            'For Each oPar In oFamEncontrada.Symbol.Family.GetParameters(nPar)
    '                            '    If Not (oPar.IsReadOnly) Then
    '                            '        Exit For
    '                            '    End If
    '                            'Next
    '                            ''
    '                            If oPar IsNot Nothing Then
    '                                If (nPar = "ITEM_WEIGHT") OrElse (nPar = "ITEM_WEIGHT_L") Then   ' AndAlso valor <> "" AndAlso valor.Contains(" ") Then
    '                                    oPar.SetValueString(valor.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT.ToString)
    '                                    ' FamilySymbol
    '                                    If oParS IsNot Nothing Then oParS.SetValueString(valor.Replace(",", "."))
    '                                ElseIf (nPar = "WEIGHT") Then   ' AndAlso valor <> "" AndAlso valor.Contains(" ") Then
    '                                    ' Calculo del peso (WEIGHT)
    '                                    Dim iGeneric As String = queFaBusco.colCV("ITEM_GENERIC").ToString
    '                                    Dim itemW As String = queFaBusco.colCV("ITEM_WEIGHT").ToString
    '                                    If iGeneric = "" OrElse iGeneric = "0" Then
    '                                        oPar.SetValueString(itemW.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                        ' FamilySymbol
    '                                        If oParS IsNot Nothing Then oParS.SetValueString(itemW.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                    ElseIf iGeneric = "1" Then
    '                                        ' Buscar Length o Height o Width (El primero que tenga valor)
    '                                        Dim length As Double = oFamEncontrada.GetParameters("ITEM_LENGTH").FirstOrDefault.AsDouble
    '                                        If length = 0 Then
    '                                            length = oFamEncontrada.GetParameters("ITEM_HEIGHT").FirstOrDefault.AsDouble
    '                                        End If
    '                                        If length = 0 Then
    '                                            length = oFamEncontrada.GetParameters("ITEM_WIDTH").FirstOrDefault.AsDouble
    '                                        End If
    '                                        ' ft = unidades internas de Revit para longitudes
    '                                        Dim lengthTemp As String = U_DameString__DesdeCualquiera(length.ToString, "ft", queFaBusco.ITEM_GENERIC_UNIT)
    '                                        'Dim lengthTemp As String = U_DameString__DesdeCualquiera(length, queFaBusco.ITEM_LENGTH_UNIT, queFaBusco.ITEM_GENERIC_UNIT)
    '                                        Dim peso As Double = CDbl(lengthTemp) * CDbl(itemW)
    '                                        oPar.SetValueString(FormatNumber(peso, 2).ToString.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                        ' FamilySymbol
    '                                        If oParS IsNot Nothing Then oParS.SetValueString(FormatNumber(peso, 2).ToString.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                    ElseIf iGeneric = "2" Then
    '                                        ' Buscar Area (parámetro calculado)
    '                                        Dim area As Double = oFamEncontrada.GetParameters("AREA").FirstOrDefault.AsDouble
    '                                        ' m2 unidades internas de Revit para areas
    '                                        Dim areaTemp As String = U_DameString__DesdeCualquiera(area.ToString, "ft2", queFaBusco.ITEM_GENERIC_UNIT)
    '                                        Dim peso As Double = CDbl(areaTemp) * CDbl(itemW)
    '                                        oPar.SetValueString(FormatNumber(peso, 2).ToString.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                        ' FamilySymbol
    '                                        If oParS IsNot Nothing Then oParS.SetValueString(FormatNumber(peso, 2).ToString.Replace(",", ".") & " " & queFaBusco.ITEM_WEIGHT_UNIT)
    '                                    End If
    '                                ElseIf (nPar = "ITEM_MARKET") Then
    '                                    oPar.SetValueString(CInt(ITEM_MARKET).ToString)
    '                                Else
    '                                    ' Solo para parámetros de Texto. Los numéricos en los If anteriores (SetValueString)
    '                                    oPar.Set(valor)
    '                                    If oParS IsNot Nothing Then oParS.Set(valor)
    '                                End If
    '                            Else
    '                                errores &= "ParametrosEscribeEnFamiliasDocumento --> " & nPar & " --> Not in Parameters of " & oFamEncontrada.Name & vbCrLf
    '                            End If
    '                        Catch ex As Exception
    '                            errores &= "ParametrosEscribeEnFamiliasDocumento --> " & ex.Message & vbCrLf & vbCrLf
    '                        End Try
    '                        ''
    '                        'System.Windows.Forms.Application.DoEvents()
    '                    Next
    '                    ''
    '                    'modULMA.PonInf()
    '                    If FamCambiadas.Contains(oFamEncontrada.Id) = False Then
    '                        FamCambiadas.Add(oFamEncontrada.Id)
    '                        If oFamEncontrada.Symbol IsNot Nothing Then FamCambiadas.Add(oFamEncontrada.Symbol.Id)
    '                    End If
    '                    oFamEncontrada = Nothing
    '                Next
    '            Next
    '            '' Vaciamos los valores del bucle For--Next
    '            total = 0
    '            FamCambiadas = Nothing
    '            ''
    '        Catch ex As Exception
    '            errores &= "ParametrosEscribeEnFamiliasDocumento --> (" & contador & ")" & ex.Message & vbCrLf & vbCrLf
    '            'trans0.RollBack()
    '        End Try
    '        ''
    '        If trans1 IsNot Nothing AndAlso trans1.GetStatus <> TransactionStatus.Committed Then trans1.Commit()
    '    End Using
    '    ''
    '    ''oDoc.Regenerate()         '' Que se vean los cambios.
    '    ''
    '    'oDoc.ActiveView.SetWorksharingDisplayMode(dMode)
    '    Return errores
    'End Function
    ''
    'Public Function CadenaBuscardynamicFamilies(fcode As String, il As String, iw As String, ih As String) As ULMALGFree.clsFamCode
    '    Dim resultado As ULMALGFree.clsFamCode = Nothing
    '    '
    '    ' M1, M2, M3 y luego [M1,M2,M3] genericos
    '    For X As Integer = 0 To arrM.Length - 1
    '        Dim W_MARKET As String = "M" & X + 1
    '        Dim queMark As String = arrM(X)
    '        '
    '        If queMark = "" Then Continue For
    '        ''
    '        '' Cadena para buscar en fam_cod (key=cadenabuscar, value=clsFamCode)
    '        Dim cadenabuscarM As String = queMark & fcode.Trim & il.Trim & iw.Trim & ih.Trim
    '        ''
    '        If colFamCode.ContainsKey(cadenabuscarM) Then
    '            resultado = CType(colFamCode(cadenabuscarM), ULMALGFree.clsFamCode)
    '            resultado.W_MARKET = W_MARKET
    '            Exit For
    '        End If
    '    Next
    '    '
    '    If resultado Is Nothing Then
    '        For X As Integer = 0 To arrM.Length - 1
    '            Dim W_MARKET As String = "M" & X + 1
    '            Dim queMark As String = arrM(X)
    '            '
    '            If queMark = "" Then Continue For
    '            ''
    '            '' Cadena para buscar en fam_cod (key=cadenabuscar, value=clsFamCode
    '            Dim cadenabuscarMGEN As String = queMark & fcode & ih
    '            Dim cadenabuscarMGEN1 As String = queMark & fcode
    '            ''
    '            If colFamCode.ContainsKey(cadenabuscarMGEN) Then
    '                resultado = CType(colFamCode(cadenabuscarMGEN), ULMALGFree.clsFamCode)
    '                resultado.W_MARKET = W_MARKET
    '                Exit For
    '            ElseIf colFamCode.ContainsKey(cadenabuscarMGEN1) Then
    '                resultado = CType(colFamCode(cadenabuscarMGEN1), ULMALGFree.clsFamCode)
    '                resultado.W_MARKET = W_MARKET
    '                Exit For
    '            End If
    '        Next
    '    End If
    '    '
    '    Return resultado
    'End Function
    ''
    '
    Public Sub AyudaAbrir()
        Dim fAyuda As String = ""
        If evRevit.evAppUI.Application.Language = LanguageType.Spanish Then
            fAyuda = IO.Path.Combine(_dirApp, "BRIO SET HELP SPANISH.pdf")
        Else
            'fAyuda = IO.Path.Combine(_dirApp, "BRIO SET HELP ENGLISH.pdf")
            fAyuda = IO.Path.Combine(_dirApp, "BRIO SET HELP.pdf")
        End If
        ''
        If fAyuda <> "" AndAlso IO.File.Exists(fAyuda) Then
            Process.Start(fAyuda)
        End If
    End Sub
    ' ''
    'public static string PrintOutRevitUnitInfo(UnitType ut, FormatOptions obj, string indent)
    '    {
    '        string msg = string.Format(indent + "{0} ({1}):" + Environment.NewLine, LabelUtils.GetLabelFor(ut), ut);

    '        msg += string.Format(indent + "\tAccuracy: {0}" + Environment.NewLine, obj.Accuracy);
    '        msg += string.Format(indent + "\tUnit display: {0} ({1})" + Environment.NewLine, LabelUtils.GetLabelFor(obj.DisplayUnits), obj.DisplayUnits);
    '        msg += string.Format(indent + "\tUnit symbol: {0}" + Environment.NewLine, obj.CanHaveUnitSymbol() ? string.Format("{0} ({1})", (obj.UnitSymbol== UnitSymbolType.UST_NONE? "": LabelUtils.GetLabelFor(obj.UnitSymbol)), obj.UnitSymbol) : "n/a");
    '        msg += string.Format(indent + "\tUse default: {0}" + Environment.NewLine, obj.UseDefault);
    '        msg += string.Format(indent + "\tUse grouping: {0}" + Environment.NewLine, obj.UseGrouping);
    '        msg += string.Format(indent + "\tUse digit grouping: {0}" + Environment.NewLine, obj.UseDigitGrouping);
    '        msg += string.Format(indent + "\tUse plus prefix: {0}" + Environment.NewLine, obj.CanUsePlusPrefix() ? obj.UsePlusPrefix.ToString() : "n/a");
    '        msg += string.Format(indent + "\tSuppress spaces: {0}" + Environment.NewLine, obj.CanSuppressSpaces() ? obj.SuppressSpaces.ToString() : "n/a");
    '        msg += string.Format(indent + "\tSuppress leading zeros: {0}" + Environment.NewLine, obj.CanSuppressLeadingZeros() ? obj.SuppressLeadingZeros.ToString() : "n/a");
    '        msg += string.Format(indent + "\tSuppress trailing zeros: {0}" + Environment.NewLine, obj.CanSuppressTrailingZeros() ? obj.SuppressTrailingZeros.ToString() : "n/a");

    '        return msg;
    '    }
    Public Enum NProp
        TYPE
        FAMILY_CODE
        ITEM_CODE
        ITEM_LENGTH
        ITEM_HEIGHT
        ITEM_WIDTH
        ITEM_WEIGHT
        ITEM_DESCRIPTION
        Count
        W_MARKET
        ITEM_GENERIC
    End Enum
    ''
    'Public Sub CargaDatosMercados()
    '    op = New CONSULTAS(False, False)  '(dirXml)
    '    ''
    '    Dim compañias As List(Of ConsultarBDI.Compania)
    '    compañias = op.getCompanias()
    '    ''
    '    '' Rellenamos Mercados y Lenguajes. Borramos las colecciones antes.
    '    If colMercadosCod IsNot Nothing Then colMercadosCod.Clear()
    '    If colMercadosDes IsNot Nothing Then colMercadosDes.Clear()
    '    If arrMercadosCod IsNot Nothing Then arrMercadosCod.Clear()
    '    If arrMercadosDes IsNot Nothing Then arrMercadosDes.Clear()
    '    ''
    '    colMercadosCod = New SortedList
    '    colMercadosDes = New SortedList
    '    arrMercadosCod = New ArrayList
    '    arrMercadosDes = New ArrayList
    '    ''
    '    For Each oC As ConsultarBDI.Compania In compañias
    '        Dim codigo As String = oC.codigo.ToString
    '        Dim descripcion As String = oC.descripcion
    '        'Dim idioma As String = oC.idioma
    '        'Dim pais As String = oC.pais
    '        ''
    '        If arrMercadosCod.Contains(codigo) = False Then arrMercadosCod.Add(codigo)
    '        If colMercadosCod.ContainsKey(codigo) = False Then colMercadosCod.Add(codigo, oC)
    '        If arrMercadosDes.Contains(descripcion) = False Then arrMercadosDes.Add(descripcion)
    '        If colMercadosDes.ContainsKey(descripcion) = False Then colMercadosDes.Add(descripcion, oC)
    '    Next
    '    arrMercadosCod.Sort()
    '    arrMercadosDes.Sort()
    'End Sub
    ''
    ''
    Public Sub RibbonPanelULMALlena()
        Dim encontrado As Boolean = False
        Dim ribbon As adWin.RibbonControl = adWin.ComponentManager.Ribbon
        ''
        '' *** Primero llenamos todos los RibbonPanel de UCRevit2018
        For Each tab As adWin.RibbonTab In ribbon.Tabs
            '' Si no es el RIBBON de PrefaBIM, continuar
            If (tab.AutomationName <> _uiLabName) Then
                Continue For
            End If
            '' RIBBONPANNELS de cada Ribbontab
            For Each oPanel As adWin.RibbonPanel In tab.Panels
                Select Case oPanel.Source.AutomationName
                    Case nombrePanelTools   ' "Tools"
                        panelToolsW = oPanel
                    Case nombrePanelAbout    ' "About"
                        panelAboutW = oPanel
                End Select
            Next
        Next
    End Sub
    ''
    Public Function ParametroProyectoLee(ByRef queDoc As Autodesk.Revit.DB.Document, quePar As String, Optional queParSub As String = "") As String
        Dim resultado As String = ""
        If queDoc Is Nothing Then queDoc = evRevit.evAppUI.ActiveUIDocument.Document
        ''
        '' Iterar con los parametros del proyecto.
        'For Each oPar As Parameter In oDoc.ProjectInformation.Parameters
        '    If oPar.Definition.Name = quePar Then
        '        resultado = oPar.AsString
        '        Exit For
        '    End If
        'Next
        ''
        '' Leer directamente por el nombre del parámetro.
        Try
            Dim pordefecto As String = "M1=" & DEFAULT_PROGRAM_MARKET & ",M2=,M3=|L1=" & DEFAULT_PROGRAM_LANGUAGE & ",L2=,L3="
            Dim oPar As Parameter = Nothing
            Try
                oPar = evRevit.evAppUI.ActiveUIDocument.Document.ProjectInformation.GetParameters(quePar).First
            Catch ex As Exception
                ' Si da error y no existe el parámetro.
                If quePar = "PROJECT_CONFIG" Then
                    resultado = DEFAULT_USER_CONFIG
                ElseIf quePar = "PROJECT_LANGUAGE" Then
                    resultado = DEFAULT_PROGRAM_LANGUAGE
                Else
                    resultado = ""
                End If
                GoTo ELFINAL
                Exit Function
            End Try
            '
            If oPar.HasValue = False AndAlso quePar = "PROJECT_CONFIG" Then
                'utilesRevit.ParametroProyectoEscribe(queDoc, quePar, pordefecto)
                If oPar.AsString Is Nothing OrElse oPar.AsString = "" Then
                    If DEFAULT_USER_CONFIG <> "" Then
                        resultado = DEFAULT_USER_CONFIG
                    Else
                        resultado = pordefecto
                    End If
                Else
                    resultado = oPar.AsString
                End If
            Else
                resultado = oPar.AsString
            End If
            'resultado = oDoc.ProjectInformation.GetParameters(varConfig).First.AsString
        Catch ex As Exception
            resultado = daError
            If queParSub = "" Then
                GoTo ELFINAL
                Exit Function
            End If
        End Try
        ''
        If queParSub <> "" Then
            Dim bloques() As String = Split(resultado, "|"c)
            Dim bl1(-1) As String
            For x As Integer = 0 To bloques.Length - 1
                If bloques(x).ToUpper.Contains(queParSub.ToUpper & "=") Then
                    bl1 = Split(bloques(x), ","c)
                    Exit For
                End If
            Next
            If bl1.Length > 0 Then
                For x As Integer = 0 To bl1.Length - 1
                    If bl1(x).ToUpper.Contains(queParSub & "=") Then
                        resultado = Split(bl1(x), "=")(1)
                    End If
                Next
            End If
        End If
        ''
ELFINAL:
        Return resultado
    End Function
#Region "TRANSLATE"
    Public Function DameMayMin(txtOri As String, txtFin As String) As String
        Dim resultado As String = ""
        '' Para realizar la conversión de cadenas a May, Min, Tipo titulo.
        Dim curCulture As System.Globalization.CultureInfo = System.Threading.Thread.CurrentThread.CurrentCulture
        Dim tInfo As System.Globalization.TextInfo = curCulture.TextInfo()
        '' Analizamos txtOri para saber si estaba en MAYUSCULAS o Tipo Titulo
        Dim tieneMay As Boolean = False
        Dim tieneMin As Boolean = False
        Dim minusculas() As String = New String() {"a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z"}
        Dim mayusculas() As String = New String() {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"}
        ''
        For x As Integer = 0 To txtOri.Length - 1
            If minusculas.Contains(txtOri.Chars(x)) Then
                tieneMin = True
            ElseIf mayusculas.Contains(txtOri.Chars(x)) Then
                tieneMay = True
            End If
        Next
        ''
        If tieneMin = True AndAlso tieneMay = False Then
            resultado = txtFin.ToLower
        ElseIf tieneMin = False AndAlso tieneMay = True Then
            resultado = txtFin.ToUpper
        ElseIf tieneMin = True AndAlso tieneMay = True Then
            resultado = tInfo.ToTitleCase(txtFin.ToLower)
        Else
            resultado = txtFin
        End If
        ''
        Return resultado.Trim
    End Function
#End Region
End Module
