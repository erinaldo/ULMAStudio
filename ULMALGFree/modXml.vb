Imports System.IO
Imports System.Xml
Imports System.Xml.Linq
'Imports System.Data
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.DB
Imports Microsoft.VisualBasic.FileIO
Imports uf = ULMALGFree.clsBase

Module modXml
    Public Const fijoDynFam As String = "_revit_dynamicfamilies"
    Public Const fijoArt As String = "_articles"
    Public Const fijoComp As String = "companies"
    Public Const fijoLang As String = "languages"
    Public Const fijoGroups As String = "_groups_subgroups_elements"

    Public Sub BorraViejosDynamicfamilies()
        '' Borrar todos los fichero viejos "XXX_dynamicfamilies.xml o XXX_dynamicfamilies.xmlc"
        Dim preViejos As String = "_dynamicfamilies"
        '
        For Each oFi As String In IO.Directory.GetFiles(uf.path_offlineBDIdata, "*.xml*", IO.SearchOption.TopDirectoryOnly)
            Dim nombre As String = IO.Path.GetFileName(oFi)
            If nombre.ToUpper.Contains(preViejos.ToUpper) AndAlso nombre.ToUpper.Contains("REVIT") = False Then
                IO.File.Delete(oFi)
            End If
        Next
    End Sub
    ' Función para poner los XML correctos (Sin codificar o codificados), si no existe el fichero
    ' copiar el de Central.
    Public Function PonFicheroXMLCorrecto(ByRef queFi As String) As String
        ' Hemos quitado la evaluación de los xmlc (Codificados)
        Dim resultado As String = ""
        '
        ' Si no enviamos nada
        If queFi = "" Then
            Return resultado
            Exit Function
        End If
        ' Si enviamos un fichero que ya existe. Salir sin copiar el de central
        If IO.File.Exists(queFi) Then
            resultado = queFi
            Return resultado
            Exit Function
        End If
        '
        '
        ' No existe el fichero. Evaluar y Copiar el de central
        Dim soloNombre As String = IO.Path.GetFileNameWithoutExtension(queFi)
        Dim queMarket As String = soloNombre.Split("_"c)(0)
        '
        ' PONER LOS NOMBRES CORRECTOS O COPIAR DE CENTRAL
        ' _articles
        ' _revit_
        Dim queFiCentral As String = ""   ' Sin codificar
        If queFi.Contains("_articles") Then
            queFiCentral = IO.Path.Combine(uf.path_offlineBDIdata, "\120" & fijoArt & ".xml")   ' Sin codificar
        ElseIf queFi.Contains("_revit_") Then
            queFiCentral = IO.Path.Combine(uf.path_offlineBDIdata, "\120" & fijoDynFam & ".xml")   ' Sin codificar
        End If
        '
        ' *** COPIAR EL DE CENTRAL
        ' Copiar el de central, si no existe el de la delegación.
        'If IsNumeric(queMarket) Then
        ' Es de delegación (120_XXX, etc)
        If IO.File.Exists(queFiCentral) Then
            'Call utilesRevit.Fichero_EsSoloLectura(queFi, True)
            IO.File.Copy(queFiCentral, queFi, False)
            ' resultado.Add("Not exist file " & IO.Path.GetFileName(queF) & ". Using " & IO.Path.GetFileName(fiCentral))
            resultado = queFi
        End If
        'Else
        'queFiC = queFiCentral
        'End If
        '
        ' *** COMPROBAR SI EXISTEN LOS FICHEROS A PROCESAR (COPIAR Y DESENCRIPTAR)
        'If IO.File.Exists(queFi) = False Then
        '    'MsgBox("Not exist files --> " & vbCrLf & IO.Path.GetFileName(queFi) & " OrElse " & IO.Path.GetFileName(queFiC))
        '    TaskDialog.Show("ERROR...", "Not exist files --> " & vbCrLf & IO.Path.GetFileName(queFi))
        '    resultado = ""
        'ElseIf IO.File.Exists(queFi) = True Then
        '    resultado = queFi
        'End If
        '
        Return resultado
    End Function
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>        (Primera línea de un fichero XML sin encriptar)
    Public Function LeeXMLLenguajes_Linq(queFichero As String) As SortedList
        queFichero = PonFicheroXMLCorrecto(queFichero)
        '<languages>
        '   <language>
        '       <languageCode>es</languageCode>
        '       <description>Español</description>
        '   </language>
        '</languages>

        Dim resultado As New SortedList
        ''
        Dim FeedXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        Dim Feeds = From Feed In FeedXML.Descendants("language")
                    Where (Feed.Attribute("status") Is Nothing) OrElse (Feed.Attribute("status").Value <> "disabled")
                    Select languageCode = Feed.Element("languageCode").Value,
                            description = Feed.Element("description").Value
        ''
        For Each item In Feeds
            If resultado.Contains(item.languageCode) = False Then resultado.Add(item.languageCode, item.description)
        Next
        ''
        'If encrypt = True Then
        '    If IO.Path.GetExtension(queFichero).ToUpper = ".XML" AndAlso IO.File.Exists(queFichero) Then
        '        IO.File.Delete(queFichero)
        '    End If
        'End If
        ''
        Return resultado
    End Function
    '
    Public Sub RellenaCompaniesXML_Linq(queFichero As String)
        queFichero = PonFicheroXMLCorrecto(queFichero)
        '<companies>
        '   <company>
        '       <companyCode>100</companyCode>
        '       <description>ULMA PORTUGAL - Cof. e And., Lda.</description>
        '       <languageCode>pt</languageCode>
        '       <countryCode>PT</countryCode>
        '   </company>
        '</companies>
        Dim companiesXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        Dim companies = From company In companiesXML.Descendants("company")
                        Where (company.Attribute("status") Is Nothing) OrElse (company.Attribute("status").Value <> "disabled")
                        Select companyCode = company.Element("companyCode").Value,
                                description = company.Element("description").Value,
                                languageCode = company.Element("languageCode").Value,
                                countryCode = company.Element("countryCode").Value
        ''
        uf.colMercadosCod = New SortedList
        uf.colMercadosDes = New SortedList
        uf.arrMercadosCod = New ArrayList
        uf.arrMercadosDes = New ArrayList
        ''
        For Each item In companies
            uf.cComs = New ULMALGFree.clsCompanies(item.companyCode, item.description, item.languageCode, item.countryCode)
            ''
            If uf.arrMercadosCod.Contains(item.companyCode) = False Then uf.arrMercadosCod.Add(item.companyCode)
            If uf.colMercadosCod.ContainsKey(item.companyCode) = False Then uf.colMercadosCod.Add(item.companyCode, uf.cComs)
            If uf.arrMercadosDes.Contains(item.description) = False Then uf.arrMercadosDes.Add(item.description)
            If uf.colMercadosDes.ContainsKey(item.description) = False Then uf.colMercadosDes.Add(item.description, uf.cComs)
            uf.arrMercadosCod.Sort()
            uf.arrMercadosDes.Sort()
        Next
        ''
        'If encrypt = True Then
        '    If IO.Path.GetExtension(queFichero).ToUpper = ".XML" AndAlso IO.File.Exists(queFichero) Then
        '        IO.File.Delete(queFichero)
        '    End If
        'End If
    End Sub
    '
    Public Sub RellenaArticulosXML_Linq(queFichero As String)
        queFichero = PonFicheroXMLCorrecto(queFichero)
        '
        uf.colArticulos = New Dictionary(Of String, ULMALGFree.clsArticulos)
        '
        Dim ArticulosXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        For Each articulo As System.Xml.Linq.XElement In ArticulosXML.Descendants("article")
            System.Threading.Thread.Sleep(1)
            Dim cArts As New ULMALGFree.clsArticulos
            '
            ' ARTICLECODE  (<articleCode>0000101</articleCode>)
            Try
                cArts.articleCode = articulo.Element("articleCode").Value
            Catch ex As Exception
                cArts.articleCode = ""
            End Try
            ''
            If uf.colArticulos.ContainsKey(cArts.articleCode) = True OrElse cArts.articleCode = "" Then
                Continue For
            End If
            '
            ' ARTICLECODEL  (<articleCodeL>0000101</articleCodeL>)
            Dim artCl As XElement = Nothing
            Try
                artCl = articulo.Element("articleCodeL")
                If artCl IsNot Nothing Then
                    Try
                        cArts.articleCodeL = artCl.Value   ' articulo.Element("articleCodeL").Value
                    Catch ex As Exception
                        cArts.articleCodeL = ""
                    End Try
                End If
            Catch ex As Exception
                ' No hacer nada
            End Try
            '
            ' DESCRIPTION
            Dim lstDescriptions As IEnumerable(Of XElement) = Nothing
            Try
                lstDescriptions = articulo.Elements("description")
            Catch ex As Exception
                '
            End Try
            cArts.colDescritions = New Hashtable
            If lstDescriptions IsNot Nothing AndAlso lstDescriptions.Count > 0 Then
                For Each des As XElement In lstDescriptions    ' articulo.Elements("description")
                    Dim queDes As String = des.Value
                    Try
                        Dim language As String = des.Attribute("language").Value
                        If language.ToUpper = "LOCAL" Then language = "local"
                        cArts.colDescritions.Add(language, queDes)
                    Catch ex As Exception
                        'Console.WriteLine(ex.Message & vbCrLf & queDes)
                    End Try
                Next
            End If
            '
            ' Para asegurarnos que están todos los idiomas
            'For Each queIdi As String In arrIdiomas
            'If uf.cArts.colDescritions.ContainsKey("local") Then
            '    uf.cArts.colDescritions.Add("local", uf.cArts.colDescritions("local"))
            'Else
            '    uf.cArts.colDescritions.Add("local", "")
            'End If
            'Next
            '
            ' ** adicionalData **
            '   <aditionalData>
            '       <material>Acero</material>
            '       <weight unit="kg">57</weight>
            '       <weightL unit="kg">57</weightL>
            '       <formArea unit="m2">1.024</formArea>
            '       <stockUnit>un</stockUnit>
            '       <genericArticleUnit>m</genericArticleUnit>
            '   </aditionalData>
            Dim lstAditionalData As IEnumerable(Of XElement) = Nothing
            Try
                lstAditionalData = articulo.Elements("aditionalData").Descendants
            Catch ex As Exception
                '
            End Try
            '
            If lstAditionalData IsNot Nothing AndAlso lstAditionalData.Count > 0 Then
                For Each aData As XElement In lstAditionalData  ' articulo.Elements("aditionalData").Descendants
                    Dim dato As String = aData.Value
                    'If dato <> "" Then
                    '    dato = dato
                    'End If
                    ''
                    If aData.Name = "material" Then
                        cArts.material = dato
                    ElseIf aData.Name = "weight" Then
                        cArts.weightUnit = aData.Attribute("unit").Value.ToLower
                        '' Poner unidades de Area para CSVs
                        If cArts.weightUnit.ToLower <> "kg" Then uf.unidPe = cArts.weightUnit.ToLower
                        ''
                        If IsNumeric(dato) AndAlso dato <> "" Then
                            dato = FormatNumber(uf.CheckDecimal_Value(dato.ToString))  ', 2)
                        End If
                        ''
                        cArts.weight = dato
                    ElseIf aData.Name = "weightL" Then
                        cArts.weightLUnit = aData.Attribute("unit").Value.ToLower
                        '' Poner unidades de Area para CSVs
                        If cArts.weightLUnit.ToLower <> "kg" Then uf.unidPe = cArts.weightLUnit.ToLower
                        ''
                        If IsNumeric(dato) AndAlso dato <> "" Then
                            dato = FormatNumber(dato.ToString)  ', 2)
                        End If
                        ''
                        cArts.weightL = dato
                    ElseIf aData.Name = "formArea" Then
                        cArts.formAreaUnit = aData.Attribute("unit").Value.ToLower
                        '' Poner unidades de Area para CSVs
                        If cArts.formAreaUnit.ToLower <> "m2" Then uf.unidAr = cArts.formAreaUnit.ToLower
                        ''
                        If IsNumeric(dato) AndAlso dato <> "" Then
                            dato = FormatNumber(dato.ToString)  ', 2)
                        End If
                        cArts.formArea = dato
                    ElseIf aData.Name = "stockUnit" Then
                        cArts.stockUnit = dato.ToLower
                    ElseIf aData.Name = "genericArticleUnit" Then
                        cArts.genericArticleUnit = dato.ToLower
                    End If
                Next
            End If
            ' Rellenar weightAll, weightLAll y formAreaAll con las unidades correspondientes.
            cArts.PesoArea_ConUnidades()
            ''
            '' Lo añadimos al Hashtable directamente. Ya que si existiera lo habríamos saltado antes.
            uf.colArticulos.Add(cArts.articleCode, cArts)
        Next
        ''
        'If encrypt = True Then
        '    If IO.Path.GetExtension(queFichero).ToUpper = ".XML" AndAlso IO.File.Exists(queFichero) Then
        '        IO.File.Delete(queFichero)
        '    End If
        'End If
        ' My.Computer.FileSystem.SpecialDirectories.Temp
        'Dim queDir As String = IO.Path.GetDirectoryName(queFichero)
        'If queDir.StartsWith(My.Computer.FileSystem.SpecialDirectories.Temp) = True AndAlso IO.File.Exists(queFichero) AndAlso enejecucion = False Then
        '    IO.File.Delete(queFichero)
        'End If
    End Sub

    'Public Sub RellenaGruposXML_Linq(queFichero As String)
    '    queFichero = PonFicheroXMLCorrecto(queFichero)
    '    ''
    '    uf.colGroups = New SortedList
    '    ''
    '    Dim GruposXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
    '    ''
    '    For Each grupo As System.Xml.Linq.XElement In GruposXML.Descendants("group")
    '        Dim cG As New ULMALGFree.clsGroup
    '        '
    '        ' <groupCode>E1851</groupCode>
    '        Try
    '            cG.groupCode = grupo.Element("groupCode").Value
    '        Catch ex As Exception
    '            cG.groupCode = ""
    '        End Try
    '        ''
    '        If uf.colGroups.ContainsKey(cG.groupCode) = True OrElse cG.groupCode = "" Then
    '            Continue For
    '        End If
    '        '
    '        ' description language="local">CILINDRO ESCUADRA CRI-35/80</description> * Hay varios
    '        Dim lstDescriptions As IEnumerable(Of XElement) = Nothing
    '        Try
    '            lstDescriptions = grupo.Elements("description")
    '        Catch ex As Exception
    '            '
    '        End Try
    '        cG.DicDescritions = New Dictionary(Of String, String)
    '        If lstDescriptions IsNot Nothing AndAlso lstDescriptions.Count > 0 Then
    '            For Each des As XElement In lstDescriptions
    '                Dim queDes As String = des.Value
    '                Try
    '                    Dim language As String = des.Attribute("language").Value
    '                    If language.ToUpper.Trim = "LOCAL" Then
    '                        cG.DefaultDescription = queDes
    '                    End If
    '                    If cG.DicDescritions.ContainsKey(language) = False Then cG.DicDescritions.Add(language, queDes)
    '                Catch ex As Exception
    '                    'Console.WriteLine(ex.Message & vbCrLf & queDes)
    '                End Try
    '            Next
    '        End If
    '        '
    '        ' <productType>EV</productType>
    '        Try
    '            cG.productType = grupo.Element("productType").Value
    '        Catch ex As Exception
    '            cG.productType = ""
    '        End Try
    '        '
    '        ' <shortName>COMAIN</shortName>
    '        Try
    '            cG.shortName = grupo.Element("shortName").Value
    '        Catch ex As Exception
    '            cG.shortName = ""
    '        End Try
    '        '
    '        ' <active>True</active>
    '        cG.active = False
    '        Try
    '            Dim activeTemp As String = grupo.Element("active").Value
    '            cG.active = (activeTemp.ToUpper.Trim = "TRUE")
    '        Catch ex As Exception
    '            cG.active = False
    '        End Try

    '        '
    '        ' ** SUBGROUPS **
    '        '   <subgroups>
    '        '       <subgroup>
    '        '           <subgroupCode>0510</subgroupCode>
    '        '           <description language = "local" > PANELES</description>
    '        Dim lstAditionalData As IEnumerable(Of XElement) = Nothing
    '        Try
    '            lstAditionalData = grupo.Elements("subgroups").Descendants("subgroup")
    '        Catch ex As Exception
    '            '
    '        End Try
    '        '
    '        If lstAditionalData IsNot Nothing AndAlso lstAditionalData.Count > 0 Then
    '            cG.DicSubGroups = New Dictionary(Of String, ULMALGFree.clsSubGroup)
    '            For Each subgroup As XElement In lstAditionalData  ' articulo.Elements("aditionalData").Descendants
    '                Dim ssG As New ULMALGFree.clsSubGroup
    '                '
    '                ' <subgroupCode>0510</subgroupCode>
    '                Try
    '                    ssG.subgroupCode = subgroup.Element("subgroupCode").Value
    '                Catch ex As Exception
    '                    ssG.subgroupCode = ""
    '                End Try
    '                '
    '                ' <description language="local">PANELES</description>
    '                ssG.DicDescritions = New Dictionary(Of String, String)
    '                For Each des As XElement In subgroup.Elements("description")
    '                    Dim queDes As String = des.Value
    '                    Try
    '                        Dim language As String = des.Attribute("language").Value
    '                        If language.ToLower = "LOCAL" Then
    '                            ssG.defaultDescription = queDes
    '                        End If
    '                        If ssG.DicDescritions.ContainsKey(language) = False Then ssG.DicDescritions.Add(language, queDes)
    '                    Catch ex As Exception
    '                        'Console.WriteLine(ex.Message & vbCrLf & queDes)
    '                    End Try
    '                Next
    '                '
    '                ' elements\block
    '                ssG.LstblockCodes = New List(Of String)
    '                For Each block As XElement In subgroup.Elements("elements").Descendants("block")
    '                    Try
    '                        Dim blockCode As String = block.Element("blockCode").Value
    '                        If blockCode <> "" AndAlso ssG.LstblockCodes.Contains(blockCode) = False Then ssG.LstblockCodes.Add(blockCode)
    '                    Catch ex As Exception

    '                    End Try
    '                Next
    '                '
    '                ' elements\article
    '                ssG.LstarticleCodes = New List(Of String)
    '                For Each article As XElement In subgroup.Elements("elements").Descendants("article")
    '                    Try
    '                        Dim articleCode As String = article.Element("articleCode").Value
    '                        If articleCode <> "" AndAlso ssG.LstarticleCodes.Contains(articleCode) = False Then ssG.LstarticleCodes.Add(articleCode)
    '                    Catch ex As Exception

    '                    End Try
    '                Next
    '                '
    '                If cG.DicSubGroups.ContainsKey(ssG.subgroupCode) = False Then cG.DicSubGroups.Add(ssG.subgroupCode, ssG)
    '            Next
    '        End If
    '        ' Rellenar weightAll, weightLAll y formAreaAll con las unidades correspondientes.
    '        If uf.colGroups.ContainsKey(cG.groupCode) = False Then uf.colGroups.Add(cG.groupCode, cG)
    '    Next
    'End Sub

    Public Sub RellenaGruposXMLParaFree_Linq(queFichero As String)
        queFichero = PonFicheroXMLCorrecto(queFichero)
        ''
        uf.colGroups = New Dictionary(Of String, ULMALGFree.clsGroup)
        ''
        Dim GruposXML As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
        ''
        For Each grupo As System.Xml.Linq.XElement In GruposXML.Descendants("group")
            System.Threading.Thread.Sleep(1)
            Dim cG As New ULMALGFree.clsGroup
            '
            ' <groupCode>E1851</groupCode>
            Try
                cG.groupCode = grupo.Element("groupCode").Value
            Catch ex As Exception
                cG.groupCode = ""
            End Try
            ''
            If uf.colGroups.ContainsKey(cG.groupCode) = True OrElse cG.groupCode = "" Then
                Continue For
            End If
            '
            ' description language="local">CILINDRO ESCUADRA CRI-35/80</description> * Hay varios
            For Each des As XElement In grupo.Elements("description")
                Dim queDes As String = des.Value
                Try
                    Dim language As String = des.Attribute("language").Value
                    If language.ToUpper.Trim = "LOCAL" Then
                        cG.DefaultDescription = queDes
                    End If
                    If cG.DicDescritions.ContainsKey(language) = False Then cG.DicDescritions.Add(language, queDes)
                Catch ex As Exception
                    'Console.WriteLine(ex.Message & vbCrLf & queDes)
                End Try
            Next
            '
            ' <productType>EV</productType> ' SuperGroup
            Try
                cG.productType = grupo.Element("productType").Value
            Catch ex As Exception
                cG.productType = ""
            End Try
            '
            ' <shortName>COMAIN</shortName>
            Try
                cG.shortName = grupo.Element("shortName").Value
            Catch ex As Exception
                cG.shortName = ""
            End Try
            '
            ' <active>True</active>
            cG.active = False
            Try
                Dim activeTemp As String = grupo.Element("active").Value
                cG.active = (activeTemp.ToUpper.Trim = "TRUE")
            Catch ex As Exception
                cG.active = False
            End Try
            '
            ' ** SUBGROUPS ** ' Cogemos de cada subgroup los articulos.
            ' las familias terminar por _[articleCode].rfa
            ' elements\article
            cG.LstarticleCodes = New List(Of String)
            For Each subgroup As XElement In grupo.Elements("subgroups").Descendants("subgroup")
                System.Threading.Thread.Sleep(1)
                For Each article As XElement In subgroup.Elements("elements").Descendants("article")
                    Try
                        Dim articleCode As String = article.Element("articleCode").Value
                        If articleCode <> "" AndAlso cG.LstarticleCodes.Contains(articleCode) = False Then
                            cG.LstarticleCodes.Add(articleCode)
                        End If
                    Catch ex As Exception

                    End Try
                Next
            Next
            ' Rellenar los nombres de los ficheros en clsGroup
            cG.RellenaFilenameOnly()
            ' Rellenar weightAll, weightLAll y formAreaAll con las unidades correspondientes.
            If uf.colGroups.ContainsKey(cG.groupCode) = False Then
                uf.colGroups.Add(cG.groupCode, cG)
            End If
        Next
    End Sub
    'Public Sub RellenaFamCodeXML_Linq(queFicheros() As String)
    '    '' Borrar todos los fichero viejos "XXX_revit_dynamicfamilies.xml"
    '    'BorraViejosDynamicfamilies()
    '    '
    '    For x As Integer = 0 To UBound(queFicheros)
    '        queFicheros(x) = PonFicheroXMLCorrecto(queFicheros(x))
    '    Next
    '    '<revitDynamicFamilies>
    '    '    <family>
    '    '           <FAMILY_CODE>3_LAYER_PLYWOOD</FAMILY_CODE>
    '    '           <ITEM_CODE>7251130</ITEM_CODE>
    '    '           <ITEM_LENGTH unit="mm">1000</ITEM_LENGTH>
    '    '           <ITEM_WIDTH unit="mm">503</ITEM_WIDTH>
    '    '           <ITEM_HEIGHT unit="mm">27</ITEM_HEIGHT>
    '    '           <ITEM_GENERIC>0</ITEM_GENERIC>
    '    '           <ITEM_DESCRIPTION language="en">3 LAYER PLYWOOD 1000x500x27</ITEM_DESCRIPTION>
    '    '           <ITEM_DESCRIPTION language="es">TRICAPA 1000x503x27</ITEM_DESCRIPTION>
    '    '   </family>
    '    '</revitDynamicFamilies>
    '    ''
    '    '' Borrar los xml viejos *.xml (Ahora trabajaremos con los codificados .xmlc)
    '    'If encrypt = True Then
    '    '    Dim viejos() As String = IO.Directory.GetFiles(IO.Path.Combine(_dirApp, ficherosXML), "*.xml", IO.SearchOption.AllDirectories)
    '    '    For Each viejo As String In viejos
    '    '        If IO.Path.GetExtension(viejo).ToUpper = ".XMLC" Then
    '    '            Continue For
    '    '        End If
    '    '        Try
    '    '            IO.File.Delete(viejo)
    '    '        Catch ex As Exception
    '    '            Console.WriteLine(ex)
    '    '        End Try
    '    '    Next
    '    'End If
    '    ''
    '    colFamCode = New Hashtable
    '    ''
    '    For Each queFichero As String In queFicheros
    '        '' Recibimos array con los 3 ficheros a procesar. Si no existe el fichero, continuar.
    '        If IO.File.Exists(queFichero) = False Then Continue For
    '        ''
    '        Dim mercado As String = Split(IO.Path.GetFileNameWithoutExtension(queFichero), "_"c)(0)
    '        '' Cargar dynamicFamilies
    '        Dim dynamicFamilies As System.Xml.Linq.XDocument = XDocument.Load(queFichero)
    '        ''
    '        For Each articulo As System.Xml.Linq.XElement In dynamicFamilies.Descendants("family")
    '            uf.cFC = New ULMALGFree.clsFamCode
    '            uf.cFC.W_MARKET = ""
    '            uf.cFC.IMARKET = mercado
    '            ''
    '            '' FAMILY_CODE  (<FAMILY_CODE>3_LAYER_PLYWOOD</FAMILY_CODE>)
    '            Try
    '                uf.cFC.FCODE = articulo.Element("FAMILY_CODE").Value
    '            Catch ex As Exception
    '                uf.cFC.FCODE = ""
    '            End Try
    '            ''
    '            '' ITEM_CODE  (<ITEM_CODE>7251G00</ITEM_CODE>)
    '            Try
    '                uf.cFC.ICODE = articulo.Element("ITEM_CODE").Value
    '            Catch ex As Exception
    '                uf.cFC.ICODE = ""
    '            End Try
    '            '
    '            Dim unidTemp As String = ""
    '            ''
    '            '' ITEM_LENGTH  (<ITEM_LENGTH unit="">0</ITEM_LENGTH>)
    '            Try
    '                uf.cFC.ILENGTH = articulo.Element("ITEM_LENGTH").Value
    '                If uf.cFC.ILENGTH = "0" Then uf.cFC.ILENGTH = ""
    '            Catch ex As Exception
    '                uf.cFC.ILENGTH = ""
    '            End Try
    '            '' ITEM_LENGTH (Atributo unit)
    '            Try
    '                uf.cFC.IUnid = articulo.Element("ITEM_LENGTH").Attribute("unit").Value
    '            Catch ex As Exception
    '                uf.cFC.IUnid = ""
    '            End Try
    '            ''
    '            '' ITEM_WIDTH  (<ITEM_WIDTH unit="">0</ITEM_WIDTH>)
    '            Try
    '                uf.cFC.IWIDTH = articulo.Element("ITEM_WIDTH").Value
    '                If uf.cFC.IWIDTH = "0" Then uf.cFC.IWIDTH = ""
    '            Catch ex As Exception
    '                uf.cFC.IWIDTH = ""
    '            End Try
    '            '' ITEM_WIDTH (Atributo unit)
    '            Try
    '                unidTemp = articulo.Element("ITEM_WIDTH").Attribute("unit").Value
    '                If uf.cFC.IUnid = "" AndAlso unidTemp <> "" Then uf.cFC.IUnid = unidTemp
    '                unidTemp = ""
    '            Catch ex As Exception
    '                ''
    '            End Try
    '            ''
    '            '' ITEM_HEIGHT  (<ITEM_HEIGHT unit="">0</ITEM_HEIGHT>)
    '            Try
    '                uf.cFC.IHEIGHT = articulo.Element("ITEM_HEIGHT").Value
    '                If uf.cFC.IHEIGHT = "0" Then uf.cFC.IHEIGHT = ""
    '            Catch ex As Exception
    '                uf.cFC.IHEIGHT = ""
    '            End Try
    '            '' ITEM_LENGTH (Atributo unit)
    '            Try
    '                unidTemp = articulo.Element("ITEM_HEIGHT").Attribute("unit").Value
    '                If uf.cFC.IUnid = "" AndAlso unidTemp <> "" Then uf.cFC.IUnid = unidTemp
    '            Catch ex As Exception
    '                ''
    '            End Try
    '            '' Si finalmente no tiene unidades, ponemos milimetros
    '            If uf.cFC.IUnid = "" Then
    '                uf.cFC.IUnid = "mm"
    '            End If
    '            If uf.cFC.IUnid.ToLower <> "mm" Then
    '                unidLo = uf.cFC.IUnid.ToLower
    '            End If
    '            ''
    '            '' ITEM_GENERIC  (<ITEM_GENERIC>2</ITEM_GENERIC>)
    '            Try
    '                uf.cFC.IGENERIC = articulo.Element("ITEM_GENERIC").Value
    '            Catch ex As Exception
    '                uf.cFC.IGENERIC = ""
    '            End Try
    '            ''
    '            '' revitfamilyCode  (<revitfamilyCode>RT000012</revitfamilyCode>)
    '            Try
    '                uf.cFC.RFCODE = articulo.Element("revitfamilyCode").Value
    '            Catch ex As Exception
    '                uf.cFC.RFCODE = ""
    '            End Try
    '            ''
    '            '' fileName  (<fileName>3 LAYER PLYWOOD_DYN.rfa</fileName>)
    '            Try
    '                uf.cFC.fileName = articulo.Element("fileName").Value
    '            Catch ex As Exception
    '                uf.cFC.fileName = ""
    '            End Try
    '            ''
    '            ''<ITEM_DESCRIPTION language="en">3 LAYER PLYWOOD 2500x500x27</ITEM_DESCRIPTION>
    '            ''<ITEM_DESCRIPTION language="es">TRICAPA 2500x500x27</ITEM_DESCRIPTION>
    '            For Each des As XElement In articulo.Elements("ITEM_DESCRIPTION")
    '                If des.Attribute("language").Value.ToUpper = "EN" Then
    '                    uf.cFC.IDESCRIPTIONEN = des.Value
    '                ElseIf des.Attribute("language").Value.ToUpper = "ES" Then
    '                    uf.cFC.IDESCRIPTIONES = des.Value
    '                End If
    '            Next
    '            '
    '            '' Fundamental ejecutar esta función. Pone las medidas en mm con 2 decimales y clave.
    '            uf.cFC.PonMedidas()
    '            '' 
    '            '' Finalmente añadimos la clase uf.cFC a la colección colFamCode (key=clave, value=uf.cFC)
    '            If colFamCode.ContainsKey(uf.cFC.clave) = False Then
    '                colFamCode.Add(uf.cFC.clave, uf.cFC)
    '            End If
    '        Next
    '        '
    '    Next
    'End Sub
    '
    ' En este caso leeremos los campos 1 (ES) y 3 (Spanish)...  El 2(Español)
    Public Sub RellenaIdiomas_CSV(queFiCSV As String,
                              ByRef quecolIdiomaCod As Dictionary(Of String, String),
                              ByRef quecolCodIdioma As Dictionary(Of String, String))
        'Dim sb As New System.Text.StringBuilder()
        ''
        quecolIdiomaCod = New Dictionary(Of String, String)
        quecolCodIdioma = New Dictionary(Of String, String)
        ''
        Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(queFiCSV)

            reader.TextFieldType = FieldType.Delimited
            reader.SetDelimiters(";") ' Los campos del archivo están separados por comas.

            ' Recorremos el archivo hasta el final del mismo.
            While Not reader.EndOfData
                Try
                    ' Obtenemos una matriz con los datos existentes en la línea actual.
                    Dim lineaActual As String() = reader.ReadFields()
                    quecolIdiomaCod.Add(lineaActual(2), lineaActual(0))
                    quecolCodIdioma.Add(lineaActual(0), lineaActual(2))
                    'For Each campo As String In lineaActual
                    'sb.AppendFormat("{0}|", campo)
                    'Next
                    'sb.AppendLine()
                Catch ex As MalformedLineException
                    MsgBox("La línea actual " & ex.Message & " no es válida.", MsgBoxStyle.Critical)
                End Try
            End While
        End Using
        ' Mostramos el contenido en la ventana de Salida de Visual Studio.
        'Console.WriteLine(sb.ToString())
    End Sub
    ''
    '' En este caso leeremos los campos 1 (ES) y 3 (Spanish)...  El 2(Español)
    'Public Sub Formatea_CSV(queFiCSV As String,
    '                        Optional unidPeso As String = "",
    '                        Optional unidArea As String = "",
    '                        Optional unidLength As String = "",
    '                        Optional unidLengthM As String = "")
    '    Dim sb As New System.Text.StringBuilder()
    '    ''
    '    Using reader As New Microsoft.VisualBasic.FileIO.TextFieldParser(queFiCSV)

    '        reader.TextFieldType = FieldType.Delimited
    '        reader.SetDelimiters(";") ' Los campos del archivo están separados por comas.

    '        ' Recorremos el archivo hasta el final del mismo.
    '        Dim contador As Integer = 0
    '        Dim cabeceras As String() = Nothing                 '' Array con los nombres de las cabeceras.
    '        Dim colNoFormatear As List(Of Integer) = Nothing    '' Indice de las columnas que no hay que formatear
    '        Dim simPeso As String = IIf(unidPeso = "", "kg", unidPeso).ToString     'kg
    '        Dim simArea As String = IIf(unidArea = "", "m2", unidArea).ToString     'm2
    '        Dim simLength As String = IIf(unidLength = "", "mm", unidLength).ToString     '"mm"
    '        Dim simLengthM As String = IIf(unidLengthM = "", "m", unidLengthM).ToString     '"m"
    '        Try
    '            Dim fvOpt As New Autodesk.Revit.DB.FormatValueOptions
    '            fvOpt.AppendUnitSymbol = True
    '            If unidPeso = "" Then
    '                simPeso = UnitFormatUtils.Format(evRevit.evAppUI.ActiveUIDocument.Document.GetUnits, UnitType.UT_Mass, 100, False, False, fvOpt).Split(" "c)(1)
    '            End If
    '            If unidArea = "" Then
    '                simArea = UnitFormatUtils.Format(evRevit.evAppUI.ActiveUIDocument.Document.GetUnits, UnitType.UT_Area, 100, False, False, fvOpt).Split(" "c)(1)
    '            End If
    '            'simLength = UnitFormatUtils.Format(oDoc.GetUnits, UnitType.UT_Length, 100, False, False, fvOpt).Split(" "c)(1)
    '            Select Case evRevit.evAppUI.ActiveUIDocument.Document.GetUnits.GetFormatOptions(UnitType.UT_Length).DisplayUnits
    '                Case DisplayUnitType.DUT_MILLIMETERS
    '                    simLength = "mm"
    '                Case DisplayUnitType.DUT_CENTIMETERS
    '                    simLength = "cm"
    '                Case DisplayUnitType.DUT_DECIMETERS
    '                    simLength = "dm"
    '                Case DisplayUnitType.DUT_METERS
    '                    simLength = "m"
    '                Case DisplayUnitType.DUT_DECIMAL_FEET
    '                    simLength = "ft"
    '                    simLengthM = "ft"
    '                Case DisplayUnitType.DUT_DECIMAL_INCHES
    '                    simLength = "in"
    '                    simLengthM = "in"
    '                Case DisplayUnitType.DUT_FEET_FRACTIONAL_INCHES
    '                    simLength = "ft-in" & comi
    '                    simLengthM = "ft-in"
    '                Case DisplayUnitType.DUT_FRACTIONAL_INCHES
    '                    simLength = "in-in" & comi
    '                    simLengthM = "'"
    '            End Select
    '            fvOpt = Nothing
    '        Catch ex As Exception
    '            ''
    '        End Try
    '        Dim encontradascabeceras As Boolean = False
    '        ''
    '        While Not reader.EndOfData
    '            Try
    '                ' Obtenemos una matriz con los datos existentes en la línea actual.
    '                Dim lineaActual As String() = reader.ReadFields()
    '                '' Si contiene la palabra CODE, será la fila inicial de cabeceras.
    '                '' El resto de filas hacia abajo, serán datos.
    '                If lineaActual(0).ToUpper.Contains("CODE") AndAlso encontradascabeceras = False Then
    '                    encontradascabeceras = True
    '                    ReDim cabeceras(lineaActual.GetUpperBound(0))
    '                    lineaActual.CopyTo(cabeceras, 0)
    '                    For x = 0 To lineaActual.Count - 1
    '                        Dim dato As String = lineaActual(x)
    '                        If dato.ToUpper = "QUANTITY" AndAlso queFiCSV.ToUpper.EndsWith("PARTS.CSV") = True Then
    '                            '' No ponemos las unidades en esta columna. Ya que puede llevar varias.
    '                        ElseIf dato.ToUpper = "AREA" AndAlso queFiCSV.ToUpper.EndsWith("AREA.CSV") = True Then
    '                            dato &= " (" & simArea & ")"
    '                            'ElseIf dato.ToUpper = "QUANTITY" AndAlso queFiCSV.ToUpper.EndsWith("LINEAL.CSV") = True Then
    '                            '    dato &= " (" & simLengthM & ")"
    '                        ElseIf dato.ToUpper.Contains("WEIGHT") = True Then
    '                            dato &= " (" & simPeso & ")"
    '                        ElseIf dato.ToUpper.EndsWith("LENGTH") OrElse dato.ToUpper.EndsWith("WIDTH") OrElse dato.ToUpper.EndsWith("HEIGHT") Then
    '                            dato &= " (" & simLength & ")"

    '                        End If
    '                        '
    '                        If x < lineaActual.Count - 1 Then
    '                            sb.AppendFormat("{0};", dato)
    '                        Else
    '                            ' No poner ; en la última columna
    '                            sb.AppendFormat("{0}", dato)
    '                        End If
    '                    Next
    '                    sb.AppendLine()
    '                    Continue While
    '                End If
    '                '
    '                ' Esto ya serán datos (Guardar QUANTITY y WEIGHT para multiplicar en TOTAL_WEIGHT)
    '                ' La columna 2, QUANTITY(1) llevará la cantidad (m2 o m en _PART, _AREA y _LINEAL respectivamente)
    '                ' La columna 4, WEIGHT (3) llevará el peso unitario en _PART, _AREA y _LINEAL.
    '                ' La columna 5, TOTAL_WEIGHT (4) llevará el peso total _PART, _AREA y _LINEAL.
    '                Dim quantity As Double = 0
    '                Dim weight As Double = 0
    '                Dim total_weight As Double = 0
    '                '
    '                For x = 0 To lineaActual.Count - 1
    '                    Dim dato As String = lineaActual(x)
    '                    '
    '                    Dim valores As String() = dato.Split(" "c)
    '                    '' Esto será un número con unidades
    '                    If valores.Count = 2 AndAlso (IsNumeric(valores(0).Chars(0)) AndAlso IsNumeric(valores(1)) = False) Then
    '                        dato = valores(0)
    '                        If dato.Contains(".") AndAlso uf._sepdecimal <> "." Then
    '                            dato = dato.Replace(".", uf._sepdecimal)
    '                        ElseIf dato.Contains(",") AndAlso uf._sepdecimal <> "," Then
    '                            dato = dato.Replace(",", uf._sepdecimal)
    '                        End If
    '                    ElseIf IsNumeric(dato) Then ' valores.Count = 1 AndAlso IsNumeric(valores(0).Chars(0)) Then
    '                        dato = valores(0)
    '                        If dato.Contains(".") AndAlso uf._sepdecimal <> "." Then
    '                            dato = dato.Replace(".", uf._sepdecimal)
    '                        ElseIf dato.Contains(",") AndAlso uf._sepdecimal <> "," Then
    '                            dato = dato.Replace(",", uf._sepdecimal)
    '                        End If
    '                    End If
    '                    ' ** Quitaremos esto cuando ya se calculen en REVIT
    '                    ' Calcularemos el total. Si no va ya calculado en la Tabla de Planificación.
    '                    ' Ahora no va calculado al no poder multiplicar M2 o M con WEIGHT (Que va en kilos)
    '                    ' Solo en PARTS, AREA y LINEAL (Aunque PARTS ya lo lleva calculado
    '                    If queFiCSV.ToUpper.EndsWith("PARTS.CSV") = True OrElse queFiCSV.ToUpper.EndsWith("AREA.CSV") = True OrElse queFiCSV.ToUpper.EndsWith("LINEAL.CSV") = True Then
    '                        If x = 1 AndAlso IsNumeric(dato) Then
    '                            quantity = CDbl(dato.Replace(",", "."))     ' Poner con . para hacer los cálculos
    '                        ElseIf x = 3 AndAlso IsNumeric(dato) Then
    '                            weight = CDbl(dato.Replace(",", "."))       ' Poner con . para hacer los cálculos
    '                        ElseIf x = 4 AndAlso quantity > 0 And weight > 0 Then
    '                            total_weight = Math.Round(quantity * weight, 3)    ' Con redondeo a 3 decimales
    '                            'total_weight = quantity * weight                    ' Sin redondear
    '                            '
    '                            ' Calcular y poner total_weight en el fichero CSV con el separador correcto.
    '                            If total_weight.ToString.Contains(".") AndAlso uf._sepdecimal <> "." Then
    '                                dato = total_weight.ToString.Replace(".", uf._sepdecimal)
    '                            ElseIf dato.Contains(",") AndAlso uf._sepdecimal <> "," Then
    '                                dato = total_weight.ToString.Replace(",", uf._sepdecimal)
    '                            End If
    '                        End If
    '                    End If
    '                    '
    '                    If x < lineaActual.Count - 1 Then
    '                        ' Poner ; como separador del siguiente valor
    '                        sb.AppendFormat("{0};", dato)
    '                    Else
    '                        ' No poner ; en la última columna
    '                        sb.AppendFormat("{0}", dato)
    '                    End If
    '                Next
    '                sb.AppendLine()
    '            Catch ex As MalformedLineException
    '                MsgBox("La línea actual " & ex.Message & " no es válida.", MsgBoxStyle.Critical)
    '            End Try
    '        End While
    '    End Using
    '    ''
    '    '' Borrar el fichero actual y volver a crearlo con el texto del StringBuilder.
    '    Try
    '        IO.File.Delete(queFiCSV)
    '        IO.File.WriteAllText(queFiCSV, sb.ToString, System.Text.Encoding.UTF8)
    '    Catch ex As Exception
    '        '' Ha dado error al borrarlo. No hacemos nada.
    '    End Try
    'End Sub
    Public Sub CargaDatosMercadosYO()
        Dim queFichero As String = ""
        'If encrypt = False Then
        queFichero = IO.Path.Combine(uf.path_offlineBDIdata, "\" & fijoComp & ".xml")
        'ElseIf encrypt = True Then
        '    queFichero = IO.Path.Combine(_dirApp, ficherosXML & "\" & fijoComp & extXMLc)
        'End If
        ''
        RellenaCompaniesXML_Linq(queFichero)
        '
    End Sub
End Module
'