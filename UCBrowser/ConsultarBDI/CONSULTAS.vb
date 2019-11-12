Imports System.Configuration
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports System.Xml
Imports Magnitudes



Public Class CONSULTAS

    Public Const COMPANIAMATRIZ As Integer = 120
    Public Const COMPANIA_ES As Integer = 120
    Public Const COMPANIA_PE As Integer = 316
    Public Const COMPANIA_CL As Integer = 318
    Public Const COMPANIA_MX As Integer = 112
    Public Const COMPANIA_FR As Integer = 102
    Public Const COMPANIA_PT As Integer = 100
    Public Const COMPANIA_US As Integer = 107
    Public Const COMPANIA_DE As Integer = 110
    Public Const COMPANIA_PL As Integer = 310
    Public Const COMPANIA_IT As Integer = 312
    Public Const COMPANIA_CZ As Integer = 300
    Public Const COMPANIA_SK As Integer = 302
    Public Const COMPANIA_RO As Integer = 304

    Public Const IDIOMA_CASTELLANO As Char = "5"c
    Public Const IDIOMA_ENGLISH As Char = "2"c
    Public Const IDIOMA_FRANÇAIS As Char = "4"c
    Public Const IDIOMA_DEUTSCH As Char = "3"c
    Public Const IDIOMA_PORTUGUESE As Char = "P"c



    Private m_CadenaConexionSQLServer As String
    Private m_CarpetaArchivosXML As IO.DirectoryInfo



    Public Enum ModoDeObtenerLosDatos
        offline
        online
    End Enum

    Public Sub New(modoDeObtenerLosDatos As ModoDeObtenerLosDatos)
        If modoDeObtenerLosDatos = CONSULTAS.ModoDeObtenerLosDatos.online Then
            m_CadenaConexionSQLServer = getCadenaConexionSQLServerPorDefecto()
            m_CarpetaArchivosXML = Nothing
        Else
            m_CadenaConexionSQLServer = Nothing
            m_CarpetaArchivosXML = getCarpetaArchivosXMLPorDefecto()
        End If
    End Sub

    Public Sub New(ByVal CarpetaAlternativaDeDondeLeerOfflineLosDatos As IO.DirectoryInfo)
        m_CadenaConexionSQLServer = Nothing
        m_CarpetaArchivosXML = CarpetaAlternativaDeDondeLeerOfflineLosDatos
    End Sub

    Public Sub New(ByVal CadenaDeConexionAlternativa_encriptada As String)
        Dim maquinaCriptografica As New CriptoCriptex("fdosbgpdfubgfousdbpgfb59ht345nowr8grvnsocfnsv809n4twt04n8tgnsdfivndfo89vf9784")
        m_CadenaConexionSQLServer = maquinaCriptografica.Descifrar(CadenaDeConexionAlternativa_encriptada)
        m_CarpetaArchivosXML = Nothing
    End Sub

    Public Function getNombreDeLaFuenteActualmenteEnUso() As String
        If IsNothing(m_CadenaConexionSQLServer) Then
            If IsNothing(m_CarpetaArchivosXML) Then
                Return ""
            Else
                Return m_CarpetaArchivosXML.FullName
            End If
        Else
            Dim patronABuscar As Regex
            patronABuscar = New Regex("Data Source=.+?;")
            Dim encontrado As Match
            encontrado = patronABuscar.Match(m_CadenaConexionSQLServer)
            If encontrado.Success Then
                Return encontrado.Value
            Else
                Return ""
            End If
        End If
    End Function


    Private Shared Function getCarpetaArchivosXMLPorDefecto() As IO.DirectoryInfo
        Dim carpeta As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
        If carpeta.EndsWith("UCBrowser") = False Then
            carpeta = IO.Path.Combine(carpeta, "UCBrowser")
        End If
        Return New IO.DirectoryInfo(carpeta _
                                    & IO.Path.DirectorySeparatorChar & "offlineBDIdata")
    End Function

    Private Shared Function getCadenaConexionSQLServerPorDefecto() As String
        'Dim CadenaConexionBDEncriptada As String
        'Using archivo As New IO.StreamReader(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) _
        '                                     & System.IO.Path.DirectorySeparatorChar + "CadenaDeConexionSQLServer.txt")
        '    CadenaConexionBDEncriptada = archivo.ReadLine
        'End Using
        Dim maquinaCriptografica As New CriptoCriptex("fdosbgpdfubgfousdbpgfb59ht345nowr8grvnsocfnsv809n4twt04n8tgnsdfivndfo89vf9784")
        'Return maquinaCriptografica.Descifrar(CadenaConexionBDEncriptada)
        Return maquinaCriptografica.Descifrar("532E16486855B1E6F51CDEC2F6D8386097E272F8994BA5F1695E1FD77597AB24B6D27BB007D0B4130C2A9D7B4EB949DC2983611178C1971642E63085A31465AF1085A1989BEE504C64122AA024FBCD9F1F5A4CE2F7B58CB974552BB7CAC609BDB7B882E2BCE2791EDFCE5F53077F6E221CD1436B3334F20E44109B557857F85DA53B1378B5E1B2EF6AB6B08191CE9A04D2C879F5FC19E752671A037B01C236D3")
    End Function





    Public Function getCompanias() As List(Of Compania)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getCompanias_online()
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getCompanias_offline()
        End If
        Return New List(Of Compania)
    End Function

    Private Function getCompanias_online() As List(Of Compania)

        Dim lista As List(Of Compania)
        lista = New List(Of Compania)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT CodCompania, Descripcion, CodIdioma, CodPais" & _
                            "  FROM Companias" & _
                            " WHERE Vigente=1" & _
                            " ORDER by CodCompania"

        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim compania As Compania
                    compania = New Compania
                    If Not IsDBNull(reader(reader.GetOrdinal("CodCompania"))) Then
                        compania.codigo = reader.GetInt16(reader.GetOrdinal("CodCompania"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("Descripcion"))) Then
                        compania.descripcion = reader.GetString(reader.GetOrdinal("Descripcion")).Trim
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodIdioma"))) Then
                        compania.idioma = reader.GetString(reader.GetOrdinal("CodIdioma")).Trim
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodPais"))) Then
                        compania.pais = reader.GetString(reader.GetOrdinal("CodPais")).Trim
                    End If
                    lista.Add(compania)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return lista

    End Function

    Private Function getCompanias_offline() As List(Of Compania)
        Dim companias As New List(Of Compania)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & Compania.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim ncompania As XmlNode
            For Each ncompania In docXML.SelectNodes(compania.zzXML_EtiquetaRaiz & "/" & compania.zzXML_EtiquetaNodo)
                Dim compania As New Compania
                compania.fromXML(LeercompaniaDesde:=ncompania)
                companias.Add(compania)
            Next
        End If
        Return companias
    End Function




    Public Function getIdiomas() As List(Of Idioma)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getIdiomas_online()
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getIdiomas_offline()
        End If
        Return New List(Of Idioma)
    End Function

    Private Function getIdiomas_online() As List(Of Idioma)

        Dim lista As List(Of Idioma)
        lista = New List(Of Idioma)

        lista.Add(New Idioma(codigo:="es", descripcion:="Español"))
        lista.Add(New Idioma(codigo:="en", descripcion:="English"))

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT DISTINCT (Companias.CodIdioma + '-'+ Companias.CodPais ) as CodLanguage, " _
                        & "                (ISNULL(Idiomas.Descripcion,i120.Descripcion) + ' ('+ ISNULL(Paises.Descripcion,p120.Descripcion) +')' ) as Description " _
                        & "  FROM Companias" _
                        & "  LEFT JOIN Idiomas ON Idiomas.CodCompania=Companias.CodCompania and Idiomas.CodIdioma=Companias.CodIdioma " _
                        & "  LEFT JOIN Idiomas i120 ON i120.CodCompania=120 and i120.CodIdioma=Companias.CodIdioma" _
                        & "  LEFT JOIN Paises ON Paises.CodCompania=Companias.CodCompania and Paises.CodPais=Companias.CodPais " _
                        & "  LEFT JOIN Paises p120 ON p120.CodCompania=120 and p120.CodPais=Companias.CodPais " _
                        & " WHERE Companias.Vigente = 1" _
                        & " ORDER BY CodLanguage"

        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim codigo As String = ""
                    Dim descripcion As String = ""
                    If Not IsDBNull(reader(reader.GetOrdinal("CodLanguage"))) Then
                        codigo = reader.GetString(reader.GetOrdinal("CodLanguage"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("Description"))) Then
                        descripcion = reader.GetString(reader.GetOrdinal("Description")).Trim
                    End If
                    lista.Add(New Idioma(codigo:=codigo, descripcion:=descripcion))
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return lista

    End Function

    Private Function getIdiomas_offline() As List(Of Idioma)
        Dim idiomas As New List(Of Idioma)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & Idioma.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nIdioma As XmlNode
            For Each nIdioma In docXML.SelectNodes(idioma.zzXML_EtiquetaRaiz & "/" & idioma.zzXML_EtiquetaNodo)
                Dim idioma As New Idioma(codigo:="", descripcion:="")
                idioma.fromXML(nIdioma)
                idiomas.Add(idioma)
            Next
        End If
        Return idiomas
    End Function




    Public Function getNormas() As List(Of Norma)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getNormas_online()
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getNormas_offline()
        End If
        Return New List(Of Norma)
    End Function

    Private Function getNormas_online() As List(Of Norma)

        Dim lista As List(Of Norma)
        lista = New List(Of Norma)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT CodNorma, Descripcion as DescripcionLocal, t1.Texto as Descripcion_en " & _
                            "  FROM Normas n" & _
                            "  LEFT JOIN dbo.Textos t1 ON t1.NumText = n.NumText" & _
                            "                         AND t1.CodCompania = " & COMPANIAMATRIZ & _
                            "                         AND t1.CodIdiomaSistema = '" & IDIOMA_ENGLISH & "'" & _
                            " WHERE Vigente=1" & _
                            "   AND n.CodCompania = '" & COMPANIAMATRIZ & "'" & _
                            " ORDER by CodNorma"

        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim norma As Norma
                    norma = New Norma
                    If Not IsDBNull(reader(reader.GetOrdinal("Codnorma"))) Then
                        norma.Codigo = reader.GetString(reader.GetOrdinal("Codnorma")).Trim
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("DescripcionLocal"))) Then
                        norma.NombreLocal = reader.GetString(reader.GetOrdinal("DescripcionLocal")).Trim
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_en"))) Then
                        norma.NombreOficialIngles = reader.GetString(reader.GetOrdinal("Descripcion_en")).Trim
                    End If
                    lista.Add(norma)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return lista

    End Function

    Private Function getNormas_offline() As List(Of Norma)
        Dim normas As New List(Of Norma)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & Norma.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nNorma As XmlNode
            For Each nNorma In docXML.SelectNodes(norma.zzXML_EtiquetaRaiz & "/" & norma.zzXML_EtiquetaNodo)
                Dim norma As New Norma
                norma.fromXML(LeerNormaDesde:=nNorma)
                normas.Add(norma)
            Next
        End If
        Return normas
    End Function





    Public Function getGrupos(ByVal CodCompania As Integer) As Dictionary(Of String, Grupo)
        If IsNothing(m_CarpetaArchivosXML) Then
            Dim grupos As New Dictionary(Of String, Grupo)
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                                 TipoElemento:=Elemento.TiposDeElemento.grupo, _
                                                 limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=False, _
                                                 limitarASoloGruposActivos:=False)
                grupos.Add(item.Key, CType(item.Value, Grupo))
            Next
            Return grupos
        ElseIf IsNothing(m_CadenaConexionSQLServer) Then
            Dim resultados As resultadosGrupoSsubgrupoElemento
            resultados = getEstructuraGruposSubgruposElementos_offline(CodCompania:=CodCompania)
            Return resultados.grupos
        Else
            Return New Dictionary(Of String, Grupo)
        End If
    End Function

    Public Function getSubgrupos(ByVal CodCompania As Integer) As Dictionary(Of String, Subgrupo)
        If IsNothing(m_CarpetaArchivosXML) Then
            Dim subgrupos As New Dictionary(Of String, Subgrupo)
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                                 TipoElemento:=Elemento.TiposDeElemento.subgrupo, _
                                                 limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=False, _
                                                 limitarASoloGruposActivos:=False)
                subgrupos.Add(item.Key, CType(item.Value, Subgrupo))
            Next
            Return subgrupos
        ElseIf IsNothing(m_CadenaConexionSQLServer) Then
            Dim resultado As resultadosGrupoSsubgrupoElemento
            resultado = getEstructuraGruposSubgruposElementos_offline(CodCompania:=CodCompania)
            Return resultado.subgrupos
        Else
            Return New Dictionary(Of String, Subgrupo)
        End If
    End Function

    Public Function getEstructuraGruposSubgruposElementos(ByVal CodCompania As Integer) As List(Of EstructuraGrupoSubgrupoElemento)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getEstructuraGruposSubgruposElementos_online(CodCompania:=CodCompania)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Dim resultado As resultadosGrupoSsubgrupoElemento
            resultado = getEstructuraGruposSubgruposElementos_offline(CodCompania:=CodCompania)
            Return resultado.estructura_GrupoSubgrupoElemento
        End If
        Return New List(Of EstructuraGrupoSubgrupoElemento)
    End Function

    Private Function getEstructuraGruposSubgruposElementos_online(ByVal CodCompania As Integer) As List(Of EstructuraGrupoSubgrupoElemento)
        Dim lista As New List(Of EstructuraGrupoSubgrupoElemento)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT e.CodGrupo,  " & _
                          "       e.CodSubgrupo," & _
                          "       e.Orden, " & _
                          "       e.CodElemento, " & _
                          "       e.TipoElemento, " & _
                          "       e.CodModoSuministro " & _
                          "  FROM dbo.ElementosGruposSubgrupos e" & _
                          "  JOIN dbo.Grupos g ON g.CodCompania = e.CodCompania" & _
                          "                    AND g.CodGrupo = e.CodGrupo" & _
                          "  JOIN dbo.Subgrupos sg ON sg.CodCompania = e.CodCompania" & _
                          "                        AND sg.CodSubgrupo = e.CodSubgrupo" & _
                          " WHERE e.CodCompania = " & CodCompania & _
                          " ORDER BY e.CodGrupo, sg.OrdenSubgrupo, e.CodSubgrupo, e.Orden, e.CodElemento"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim unElemento As New EstructuraGrupoSubgrupoElemento
                    If Not IsDBNull(reader(reader.GetOrdinal("CodElemento"))) Then
                        unElemento.CodElemento = reader.GetString(reader.GetOrdinal("CodElemento")).Trim

                        Try
                            unElemento.TipoElemento = CType([Enum].Parse(enumType:=GetType(Elemento.TiposDeElemento), _
                                                                         value:=reader.GetInt16(reader.GetOrdinal("TipoElemento")).ToString),  _
                                                            Elemento.TiposDeElemento)
                        Catch ex As ArgumentException
                            unElemento.TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
                        End Try

                        If Not IsDBNull(reader(reader.GetOrdinal("Orden"))) Then
                            unElemento.NumOrdenParaListar = reader.GetInt16(reader.GetOrdinal("Orden"))
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("CodModoSuministro"))) Then
                            unElemento.CodModoSuministro = reader.GetString(reader.GetOrdinal("CodModoSuministro")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("CodSubgrupo"))) Then
                            unElemento.CodSubgrupoAlQuePertenece = reader.GetString(reader.GetOrdinal("CodSubgrupo")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("CodGrupo"))) Then
                            unElemento.CodGrupoAlQuePertenece = reader.GetString(reader.GetOrdinal("CodGrupo")).Trim
                        End If

                        lista.Add(unElemento)
                    End If
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return lista
    End Function

    Private Class resultadosGrupoSsubgrupoElemento
        Friend grupos As Dictionary(Of String, Grupo)
        Friend subgrupos As Dictionary(Of String, Subgrupo)
        Friend estructura_GrupoSubgrupoElemento As List(Of EstructuraGrupoSubgrupoElemento)
    End Class
    Private Function getEstructuraGruposSubgruposElementos_offline(ByVal CodCompania As Integer) As resultadosGrupoSsubgrupoElemento
        Dim listaGrupos As New Dictionary(Of String, Grupo)
        Dim listaSubgrupos As New Dictionary(Of String, Subgrupo)
        Dim listaEstructura As New List(Of EstructuraGrupoSubgrupoElemento)


        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & EstructuraGrupoSubgrupoElemento.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)

            Dim nGrupo As XmlNode
            For Each nGrupo In docXML.SelectNodes(EstructuraGrupoSubgrupoElemento.zzXML_EtiquetaRaiz & "/" & Elemento.zzXML_EtiquetaNodoGrupo)

                Dim grupo As New Grupo()
                grupo.fromXML(LeerElementoDesde:=nGrupo)
                listaGrupos.Add(grupo.CodElemento, grupo)

                Dim nSubgrupo As XmlNode
                For Each nSubgrupo In nGrupo.SelectNodes(EstructuraGrupoSubgrupoElemento.zzXML_EtiquetaNodoSubgrupos & "/" & Elemento.zzXML_EtiquetaNodoSubgrupo)

                    Dim subgrupo As New Subgrupo()
                    subgrupo.fromXML(LeerElementoDesde:=nSubgrupo)
                    subgrupo.CodGrupoProducto = grupo.CodElemento
                    listaSubgrupos.Add(key:=subgrupo.getCodigoUnificadoConGrupo, value:=subgrupo)

                    Dim nElemento As XmlNode
                    For Each nElemento In nSubgrupo.SelectNodes(EstructuraGrupoSubgrupoElemento.zzXML_EtiquetaNodoElementos & "/*")

                        Dim estructuraGSE As New EstructuraGrupoSubgrupoElemento
                        estructuraGSE.fromXML(LeerEstructuraDesde:=nElemento, _
                                              CodGrupoAAsignarle:=grupo.CodElemento, _
                                              CodSubgrupoAAsignarle:=subgrupo.CodElemento)
                        listaEstructura.Add(estructuraGSE)
                    Next
                Next
            Next
        End If

        Dim resultados As New resultadosGrupoSsubgrupoElemento
        resultados.grupos = listaGrupos
        resultados.subgrupos = listaSubgrupos
        resultados.estructura_GrupoSubgrupoElemento = listaEstructura
        Return resultados
    End Function




    Public Function getArticulos(ByVal CodCompania As Integer, _
                                 ByVal limitarASoloLosPresentesEnGrupoSubgrupo As Boolean, _
                                 ByVal limitarASoloGruposActivos As Boolean) _
                                As Dictionary(Of String, Articulo)
        Dim articulos As New Dictionary(Of String, Articulo)
        If IsNothing(m_CarpetaArchivosXML) Then
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                                 TipoElemento:=Elemento.TiposDeElemento.articulo, _
                                                 limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=limitarASoloLosPresentesEnGrupoSubgrupo, _
                                                 limitarASoloGruposActivos:=limitarASoloGruposActivos)
                articulos.Add(item.Key, CType(item.Value, Articulo))
            Next
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            For Each item In getElementosConDatosAdicionales_offline(CodCompania, Elemento.TiposDeElemento.articulo).elementos
                articulos.Add(item.Key, CType(item.Value, Articulo))
            Next
        End If
        Return articulos
    End Function

    Public Function getDatosAdicionalesDeArticulos(ByVal CodCompania As Integer) As Dictionary(Of String, ElementoDatosadicionales)
        Return getDatosAdicionalesDeElementos(CodCompania, Elemento.TiposDeElemento.articulo)
    End Function



    Public Function getConfiguraciones(ByVal CodCompania As Integer) As Dictionary(Of String, Configuracion)
        Dim configuraciones As New Dictionary(Of String, Configuracion)
        If IsNothing(m_CarpetaArchivosXML) Then
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                       TipoElemento:=Elemento.TiposDeElemento.configuracion, _
                                       limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=False, _
                                       limitarASoloGruposActivos:=False)
                configuraciones.Add(item.Key, CType(item.Value, Configuracion))
            Next
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            For Each item In getEstructuraConfiguracionesArticulos_offline(CodCompania).configuraciones
                configuraciones.Add(item.Key, CType(item.Value, Configuracion))
            Next
        End If
        Return configuraciones
    End Function

    Public Function getEstructuraConfiguracionesArticulos(ByVal CodCompania As Integer) As List(Of EstructuraConfiguracionArticulo)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getEstructuraConfiguracionesArticulos_online(CodCompania:=CodCompania)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getEstructuraConfiguracionesArticulos_offline(CodCompania).configuracionesarticulos
        End If
        Return New List(Of EstructuraConfiguracionArticulo)
    End Function

    Public Function getDatosAdicionalesDeConfiguraciones(ByVal CodCompania As Integer) As Dictionary(Of String, ElementoDatosadicionales)
        Return getDatosAdicionalesDeElementos(CodCompania, Elemento.TiposDeElemento.configuracion)
    End Function



    Public Function getArticulosGenericos(ByVal CodCompania As Integer, _
                                          ByVal limitarASoloLosPresentesEnGrupoSubgrupo As Boolean, _
                                          ByVal limitarASoloGruposActivos As Boolean) _
                                         As Dictionary(Of String, ArticuloGenerico)
        Dim articulos As New Dictionary(Of String, ArticuloGenerico)
        If IsNothing(m_CarpetaArchivosXML) Then
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                       TipoElemento:=Elemento.TiposDeElemento.articuloGenerico, _
                                       limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=limitarASoloLosPresentesEnGrupoSubgrupo, _
                                       limitarASoloGruposActivos:=limitarASoloGruposActivos)
                articulos.Add(item.Key, CType(item.Value, ArticuloGenerico))
            Next
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            For Each item In getElementosConDatosAdicionales_offline(CodCompania, Elemento.TiposDeElemento.articuloGenerico).elementos
                articulos.Add(item.Key, CType(item.Value, ArticuloGenerico))
            Next
        End If
        Return articulos
    End Function



    Public Function getBloquesAuxiliares(ByVal CodCompania As Integer) As Dictionary(Of String, BloqueAuxiliarAutocad)
        Dim bloques As New Dictionary(Of String, BloqueAuxiliarAutocad)
        If IsNothing(m_CarpetaArchivosXML) Then
            For Each item In getElementos_online(CodCompania:=CodCompania, _
                                       TipoElemento:=Elemento.TiposDeElemento.bloqueAuxiliarAutocad, _
                                       limitarASoloElementosPresentesEnAlgunGrupoSubgrupo:=False, _
                                       limitarASoloGruposActivos:=False)
                bloques.Add(item.Key, CType(item.Value, BloqueAuxiliarAutocad))
            Next
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            For Each item In getElementosConDatosAdicionales_offline(CodCompania, Elemento.TiposDeElemento.bloqueAuxiliarAutocad).elementos
                bloques.Add(item.Key, CType(item.Value, BloqueAuxiliarAutocad))
            Next
        End If
        Return bloques
    End Function





    Private Const zzzCriteriosDeConsultaSQLParaCogerSoloArticulosUtilizables As String = _
                                "   AND a.CodFamiliaArticulo <> 'BORRAR' " & _
                                "   AND a.CodTipoProducto NOT IN ('00R','0FN') " & _
                                "   AND (a.CodigoAviso NOT IN ('OBG','OBL','LOG','LOC','UNI','OBS','ABR') OR a.CodigoAviso IS NULL) " & _
                                "   AND (a.CodigoSeleccion NOT IN ('  G',' PE',' T9') OR a.CodigoSeleccion IS NULL)"
    ''//Explicacion de los campos filtro:
    ''// CodTipoProducto: 00R es materia prima, 0FN es semielaborado no vendible
    ''// CodigoAviso (viejos, no se usan desde octubre2018): OBS es obsoleto, ABR es a borrar, 
    ''// CodigoAviso: OBG es obsoleto global, OBL es obsoleto local, , 
    ''//              LOG es bloqueo logistico global, LOC es bloqueo logistico local,
    ''//              UNI es unificado (este código se ha retirado, hay otro codigo válido)
    ''// CodigoSeleccion: G corresponde a Agricola, PE a Ulmape y T9 a articulos especiales

    Private Function generar_Trozo03_Joins_para(nombreTabla As String, nombreColumna As String) As String
        Return "  FROM dbo." & nombreTabla & " a" & _
               "  LEFT JOIN dbo." & nombreTabla & " b ON b.CodCompania = " & COMPANIAMATRIZ & _
               "                           AND b." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b1 ON b1.CodCompania = " & COMPANIA_PL & _
               "                            AND b1." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b2 ON b2.CodCompania = " & COMPANIA_DE & _
               "                            AND b2." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b3 ON b3.CodCompania = " & COMPANIA_ES & _
               "                            AND b3." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b4 ON b4.CodCompania = " & COMPANIA_PT & _
               "                            AND b4." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b5 ON b5.CodCompania = " & COMPANIA_FR & _
               "                            AND b5." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b6 ON b6.CodCompania = " & COMPANIA_US & _
               "                            AND b6." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b7 ON b7.CodCompania = " & COMPANIA_MX & _
               "                            AND b7." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b8 ON b8.CodCompania = " & COMPANIA_PE & _
               "                            AND b8." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b9 ON b9.CodCompania = " & COMPANIA_CL & _
               "                            AND b9." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b10 ON b10.CodCompania = " & COMPANIA_CZ & _
               "                            AND b10." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b11 ON b11.CodCompania = " & COMPANIA_SK & _
               "                            AND b11." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b12 ON b12.CodCompania = " & COMPANIA_RO & _
               "                            AND b12." & nombreColumna & " = a." & nombreColumna & _
               "  LEFT JOIN dbo." & nombreTabla & " b13 ON b13.CodCompania = " & COMPANIA_IT & _
               "                            AND b13." & nombreColumna & " = a." & nombreColumna
    End Function

    Private Function getElementos_online(ByVal CodCompania As Integer, _
                                         ByVal TipoElemento As Elemento.TiposDeElemento, _
                                         ByVal limitarASoloElementosPresentesEnAlgunGrupoSubgrupo As Boolean, _
                                         ByVal limitarASoloGruposActivos As Boolean) _
                                      As Dictionary(Of String, Elemento)
        Dim ListaElementos As New Dictionary(Of String, Elemento)

        Dim trozo02_Selects_ComunesATodosLosTiposDeElemento As String
        trozo02_Selects_ComunesATodosLosTiposDeElemento =
                       "                a.Descripcion as DescripcionLocal, " & _
                       "                b.Descripcion as Descripcion_es, t1.Texto as Descripcion_en, t2.Texto as Descripcion_fr," & _
                       "                b1.Descripcion as Descripcion_plPL," & _
                       "                b2.Descripcion as Descripcion_deDE," & _
                       "                b3.Descripcion as Descripcion_esES," & _
                       "                b4.Descripcion as Descripcion_ptPT," & _
                       "                b5.Descripcion as Descripcion_frFR," & _
                       "                b6.Descripcion as Descripcion_enUS," & _
                       "                b7.Descripcion as Descripcion_esMX," & _
                       "                b8.Descripcion as Descripcion_esPE," & _
                       "                b9.Descripcion as Descripcion_esCL," & _
                       "                b10.Descripcion as Descripcion_czCZ," & _
                       "                b11.Descripcion as Descripcion_skSK," & _
                       "                b12.Descripcion as Descripcion_roRO," & _
                       "                b13.Descripcion as Descripcion_itIT"

        Dim trozo02bis_Selects_AbreviaturasPorIdiomasParaLosGrupos As String
        If TipoElemento = Elemento.TiposDeElemento.grupo Then
            trozo02bis_Selects_AbreviaturasPorIdiomasParaLosGrupos =
                           "                a.Abreviatura as AbreviaturaLocal, " & _
                           "                b.Abreviatura as Abreviatura_es, tr.TextoTraducido as Abreviatura_en, " & _
                           "                b1.Abreviatura as Abreviatura_plPL," & _
                           "                b2.Abreviatura as Abreviatura_deDE," & _
                           "                b3.Abreviatura as Abreviatura_esES," & _
                           "                b4.Abreviatura as Abreviatura_ptPT," & _
                           "                b5.Abreviatura as Abreviatura_frFR," & _
                           "                b6.Abreviatura as Abreviatura_enUS," & _
                           "                b7.Abreviatura as Abreviatura_esMX," & _
                           "                b8.Abreviatura as Abreviatura_esPE," & _
                           "                b9.Abreviatura as Abreviatura_esCL," & _
                           "                b10.Abreviatura as Abreviatura_czCZ," & _
                           "                b11.Abreviatura as Abreviatura_skSK," & _
                           "                b12.Abreviatura as Abreviatura_roRO," & _
                           "                b13.Abreviatura as Abreviatura_itIT,"
        Else
            trozo02bis_Selects_AbreviaturasPorIdiomasParaLosGrupos = ""
        End If

        Dim trozo01_InicioDeSelect_SegunTipoDeElemento As String
        Dim trozo03_Joins_SegunTipoDeElemento As String

        Select Case TipoElemento

            Case Elemento.TiposDeElemento.articulo
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT  a.CodArticulo as Codigo, " & _
                                                             "'' as CodLineaProducto, '' as CodGrupoProducto,  '' as PosicionDeVisualizacion, null as Activo, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="Articulos", nombreColumna:="CodArticulo")

            Case Elemento.TiposDeElemento.configuracion
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT  a.CodConfiguracion as Codigo, " & _
                                                             "'' as CodLineaProducto, '' as CodGrupoProducto, '' as PosicionDeVisualizacion, null as Activo, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="Configuraciones", nombreColumna:="CodConfiguracion")

            Case Elemento.TiposDeElemento.articuloGenerico
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT  a.CodArticuloGenerico as Codigo, " & _
                                                             "'' as CodLineaProducto, '' as CodGrupoProducto, '' as PosicionDeVisualizacion, null as Activo, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="ArticulosGenericos", nombreColumna:="CodArticuloGenerico")

            Case Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT  a.CodBloqueAuxiliar as Codigo, " & _
                                                             "'' as CodLineaProducto, '' as CodGrupoProducto,  '' as PosicionDeVisualizacion, null as Activo, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="BloquesAuxiliares", nombreColumna:="CodBloqueAuxiliar")

            Case Elemento.TiposDeElemento.grupo
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT a.CodGrupo as Codigo, " & _
                                                             "a.CodLineaProducto as CodLineaProducto, a.CodGrupo as CodGrupoProducto, '' as PosicionDeVisualizacion, a.Activo as Activo, " & _
                                                             "a.Abreviatura, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="Grupos", nombreColumna:="CodGrupo") & _
                                                    "  LEFT JOIN dbo.Traducciones tr ON tr.NumTraduccion = b.NumTraduccion" & _
                                                    "                               AND tr.CodCompania = " & COMPANIAMATRIZ & _
                                                    "                               AND tr.CodIdiomaSistema = '" & IDIOMA_ENGLISH & "'"

            Case Elemento.TiposDeElemento.subgrupo
                trozo01_InicioDeSelect_SegunTipoDeElemento = "SELECT a.CodSubgrupo as Codigo, " & _
                                                             "'' as CodLineaProducto, a.CodGrupo as CodGrupoProducto, a.OrdenSubgrupo as PosicionDeVisualizacion, null as Activo, "
                trozo03_Joins_SegunTipoDeElemento = generar_Trozo03_Joins_para(nombreTabla:="Subgrupos", nombreColumna:="CodSubgrupo")

            Case Else
                trozo01_InicioDeSelect_SegunTipoDeElemento = ""
                trozo03_Joins_SegunTipoDeElemento = ""
        End Select


        Dim trozo04_JoinsYWheres_ComunesATodosLosTiposDeElemento As String
        trozo04_JoinsYWheres_ComunesATodosLosTiposDeElemento = _
                      "  LEFT JOIN dbo.Textos t1 ON t1.NumText = b.NumText" & _
                      "                         AND t1.CodCompania = " & COMPANIAMATRIZ & _
                      "                         AND t1.CodIdiomaSistema = '" & IDIOMA_ENGLISH & "'" & _
                      "  LEFT JOIN dbo.Textos t2 ON t2.NumText = b.NumText" & _
                      "                         AND t2.CodCompania = " & COMPANIAMATRIZ & _
                      "                         AND t2.CodIdiomaSistema = '" & IDIOMA_FRANÇAIS & "'" & _
                      " WHERE a.CodCompania = " & CodCompania & " "

        Dim trozo05_Wheres_QueLimitanArticulosASoloArticulosUtilizables As String
        If TipoElemento = Elemento.TiposDeElemento.articulo Then
            trozo05_Wheres_QueLimitanArticulosASoloArticulosUtilizables = zzzCriteriosDeConsultaSQLParaCogerSoloArticulosUtilizables & _
                                                                          "   AND (b.CodigoAviso NOT IN ('OBG','LOG','UNI') OR b.CodigoAviso IS NULL) "
        Else
            trozo05_Wheres_QueLimitanArticulosASoloArticulosUtilizables = ""
        End If


        Dim trozo06_WheresQueLimitanArticulosASoloLosPresentesEnAlgunGrupoSubgrupo As String
        If TipoElemento = Elemento.TiposDeElemento.articulo And limitarASoloElementosPresentesEnAlgunGrupoSubgrupo Then
            If limitarASoloGruposActivos Then
                trozo06_WheresQueLimitanArticulosASoloLosPresentesEnAlgunGrupoSubgrupo = _
                             "   AND a.CodArticulo IN (SELECT distinct CodElemento " _
                           & "                           FROM dbo.ElementosGruposSubgrupos e " _
                           & "	                         LEFT JOIN dbo.Grupos g " _
                           & "	                                ON g.CodCompania = e.CodCompania " _
                           & "                                 AND g.CodGrupo = e.CodGrupo " _
                           & "                          WHERE e.CodCompania = " & CodCompania _
                           & "                            AND g.Activo = 'True'" _
                           & "                            AND e.TipoElemento = " & Elemento.TiposDeElemento.articulo & ") "
            Else
                trozo06_WheresQueLimitanArticulosASoloLosPresentesEnAlgunGrupoSubgrupo = _
                             "   AND a.CodArticulo IN (SELECT distinct CodElemento " _
                           & "                           FROM dbo.ElementosGruposSubgrupos e " _
                           & "                          WHERE e.CodCompania = " & CodCompania _
                           & "                            AND e.TipoElemento = " & Elemento.TiposDeElemento.articulo & ") "
            End If
        Else
            trozo06_WheresQueLimitanArticulosASoloLosPresentesEnAlgunGrupoSubgrupo = ""
        End If

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 180 '//por defecto 30 segundos; nosotros ponemos 180 segundos; si hace falta mas segundos se pueden aumentar aquí.

        cmd.CommandText = trozo01_InicioDeSelect_SegunTipoDeElemento & _
                          trozo02bis_Selects_AbreviaturasPorIdiomasParaLosGrupos & _
                          trozo02_Selects_ComunesATodosLosTiposDeElemento & _
                          trozo03_Joins_SegunTipoDeElemento & _
                          trozo04_JoinsYWheres_ComunesATodosLosTiposDeElemento & _
                          trozo05_Wheres_QueLimitanArticulosASoloArticulosUtilizables & _
                          trozo06_WheresQueLimitanArticulosASoloLosPresentesEnAlgunGrupoSubgrupo


        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    If Not IsDBNull(reader(reader.GetOrdinal("Codigo"))) Then
                        Dim unElemento As Elemento
                        Try
                            Select Case TipoElemento

                                Case Elemento.TiposDeElemento.articulo
                                    unElemento = New Articulo()
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=unElemento.CodElemento, value:=unElemento)

                                Case Elemento.TiposDeElemento.articuloGenerico
                                    unElemento = New ArticuloGenerico()
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=unElemento.CodElemento, value:=unElemento)

                                Case Elemento.TiposDeElemento.configuracion
                                    unElemento = New Configuracion()
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=unElemento.CodElemento, value:=unElemento)

                                Case Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                                    unElemento = New BloqueAuxiliarAutocad()
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=unElemento.CodElemento, value:=unElemento)

                                Case Elemento.TiposDeElemento.grupo
                                    unElemento = New Grupo()
                                    If Not IsDBNull(reader(reader.GetOrdinal("CodLineaProducto"))) Then
                                        CType(unElemento, Grupo).CodLineaProducto = reader.GetString(reader.GetOrdinal("CodLineaProducto")).Trim
                                    End If
                                    If Not IsDBNull(reader(reader.GetOrdinal("Activo"))) Then
                                        ''Boolean.TryParse(reader.GetString(reader.GetOrdinal("Activo")), unElemento.estaActivo)
                                        CType(unElemento, Grupo).estaActivo = reader.GetBoolean(reader.GetOrdinal("Activo"))
                                    End If
                                    If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura"))) Then
                                        CType(unElemento, Grupo).Abreviatura = reader.GetString(reader.GetOrdinal("Abreviatura")).Trim
                                    End If
                                    ProcesarCamposComunesDeGrupo(CType(unElemento, Grupo), reader)
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=unElemento.CodElemento, value:=unElemento)

                                Case Elemento.TiposDeElemento.subgrupo
                                    unElemento = New Subgrupo()
                                    If Not IsDBNull(reader(reader.GetOrdinal("CodGrupoProducto"))) Then
                                        CType(unElemento, Subgrupo).CodGrupoProducto = reader.GetString(reader.GetOrdinal("CodGrupoProducto")).Trim
                                    End If
                                    If Not IsDBNull(reader(reader.GetOrdinal("PosicionDeVisualizacion"))) Then
                                        If TypeOf (reader(reader.GetOrdinal("PosicionDeVisualizacion"))) Is System.Int32 Then
                                            CType(unElemento, Subgrupo).posicionDeVisualizacion = reader.GetInt32(reader.GetOrdinal("PosicionDeVisualizacion"))
                                        Else
                                            CType(unElemento, Subgrupo).posicionDeVisualizacion = 0
                                        End If
                                    End If
                                    ProcesarCamposComunesATodosLosElementos(unElemento, reader)
                                    ListaElementos.Add(key:=CType(unElemento, Subgrupo).getCodigoUnificadoConGrupo, value:=unElemento)

                            End Select
                        Catch ex As ArgumentException
                            RegistrosyMensajes.getMensajero.DarAviso(Nivel:=RegistrosyMensajes.nivel.n4_WARNING, _
                                                                     Mensaje:=ex.Message, _
                                                                     Procedencia:=ex.StackTrace)
                            Continue While
                        End Try

                    End If
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return ListaElementos
    End Function

    Private Sub ProcesarCamposComunesATodosLosElementos(ByRef unElemento As Elemento, ByRef reader As SqlDataReader)

        unElemento.CodElemento = reader.GetString(reader.GetOrdinal("Codigo")).Trim

        If Not IsDBNull(reader(reader.GetOrdinal("DescripcionLocal"))) Then
            unElemento.NombreLocal = reader.GetString(reader.GetOrdinal("DescripcionLocal")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_es"))) Then
            unElemento.NAME_es = reader.GetString(reader.GetOrdinal("Descripcion_es")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_en"))) Then
            unElemento.NAME_en = reader.GetString(reader.GetOrdinal("Descripcion_en")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_fr"))) Then
            unElemento.NAME_fr = reader.GetString(reader.GetOrdinal("Descripcion_fr")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_plPL"))) Then
            unElemento.NAME_plPL = reader.GetString(reader.GetOrdinal("Descripcion_plPL")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_deDE"))) Then
            unElemento.NAME_deDE = reader.GetString(reader.GetOrdinal("Descripcion_deDE")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_esES"))) Then
            unElemento.NAME_esES = reader.GetString(reader.GetOrdinal("Descripcion_esES")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_ptPT"))) Then
            unElemento.NAME_ptPT = reader.GetString(reader.GetOrdinal("Descripcion_ptPT")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_itIT"))) Then
            unElemento.NAME_itIT = reader.GetString(reader.GetOrdinal("Descripcion_itIT")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_frFR"))) Then
            unElemento.NAME_frFR = reader.GetString(reader.GetOrdinal("Descripcion_frFR")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_enUS"))) Then
            unElemento.NAME_enUS = reader.GetString(reader.GetOrdinal("Descripcion_enUS")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_esMX"))) Then
            unElemento.NAME_esMX = reader.GetString(reader.GetOrdinal("Descripcion_esMX")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_esPE"))) Then
            unElemento.NAME_esPE = reader.GetString(reader.GetOrdinal("Descripcion_esPE")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_esCL"))) Then
            unElemento.NAME_esCL = reader.GetString(reader.GetOrdinal("Descripcion_esCL")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_czCZ"))) Then
            unElemento.NAME_czCZ = reader.GetString(reader.GetOrdinal("Descripcion_czCZ")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_skSK"))) Then
            unElemento.NAME_skSK = reader.GetString(reader.GetOrdinal("Descripcion_skSK")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion_roRO"))) Then
            unElemento.NAME_roRO = reader.GetString(reader.GetOrdinal("Descripcion_roRO")).Trim
        End If

        ''Nota: Lo de poner las descripciones localizadas solo en caso de ser distintas de NAME_es 
        ''      es debido a que los articulos nacen todos en la compañia 120 y se copian tal cual a las filiales.
        ''      Es decir, todos los artículos llegan en castellano y los van traduciendo a medida que los van necesitando.
        ''Nota: Este filtro tiene el PELIGRO de dejar en blanco aquellas descripciones que sean iguales en todos los idiomas.
        ''      (por ejemplo, nombres de grupo que sean solo el nombre del producto y no tengan ninguna otra palabra que dependa del idioma)
        If unElemento.NAME_frFR = unElemento.NAME_es Then unElemento.NAME_frFR = ""
        If unElemento.NAME_ptPT = unElemento.NAME_es Then unElemento.NAME_ptPT = ""
        If unElemento.NAME_deDE = unElemento.NAME_es Then unElemento.NAME_deDE = ""
        If unElemento.NAME_plPL = unElemento.NAME_es Then unElemento.NAME_plPL = ""
        If unElemento.NAME_enUS = unElemento.NAME_es Then unElemento.NAME_enUS = ""
        If unElemento.NAME_itIT = unElemento.NAME_es Then unElemento.NAME_itIT = ""
        If unElemento.NAME_czCZ = unElemento.NAME_es Then unElemento.NAME_czCZ = ""
        If unElemento.NAME_skSK = unElemento.NAME_es Then unElemento.NAME_skSK = ""
        If unElemento.NAME_roRO = unElemento.NAME_es Then unElemento.NAME_roRO = ""

    End Sub

    Private Sub ProcesarCamposComunesDeGrupo(ByRef unGrupo As Grupo, ByRef reader As SqlDataReader)

        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_es"))) Then
            unGrupo.SHORT_es = reader.GetString(reader.GetOrdinal("Abreviatura_es")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_en"))) Then
            unGrupo.SHORT_en = reader.GetString(reader.GetOrdinal("Abreviatura_en")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_plPL"))) Then
            unGrupo.SHORT_plPL = reader.GetString(reader.GetOrdinal("Abreviatura_plPL")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_deDE"))) Then
            unGrupo.SHORT_deDE = reader.GetString(reader.GetOrdinal("Abreviatura_deDE")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_esES"))) Then
            unGrupo.SHORT_esES = reader.GetString(reader.GetOrdinal("Abreviatura_esES")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_ptPT"))) Then
            unGrupo.SHORT_ptPT = reader.GetString(reader.GetOrdinal("Abreviatura_ptPT")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_itIT"))) Then
            unGrupo.SHORT_itIT = reader.GetString(reader.GetOrdinal("Abreviatura_itIT")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_frFR"))) Then
            unGrupo.SHORT_frFR = reader.GetString(reader.GetOrdinal("Abreviatura_frFR")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_enUS"))) Then
            unGrupo.SHORT_enUS = reader.GetString(reader.GetOrdinal("Abreviatura_enUS")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_esMX"))) Then
            unGrupo.SHORT_esMX = reader.GetString(reader.GetOrdinal("Abreviatura_esMX")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_esPE"))) Then
            unGrupo.SHORT_esPE = reader.GetString(reader.GetOrdinal("Abreviatura_esPE")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_esCL"))) Then
            unGrupo.SHORT_esCL = reader.GetString(reader.GetOrdinal("Abreviatura_esCL")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_czCZ"))) Then
            unGrupo.SHORT_czCZ = reader.GetString(reader.GetOrdinal("Abreviatura_czCZ")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_skSK"))) Then
            unGrupo.SHORT_skSK = reader.GetString(reader.GetOrdinal("Abreviatura_skSK")).Trim
        End If
        If Not IsDBNull(reader(reader.GetOrdinal("Abreviatura_roRO"))) Then
            unGrupo.SHORT_roRO = reader.GetString(reader.GetOrdinal("Abreviatura_roRO")).Trim
        End If

    End Sub

    Private Function getDatosAdicionalesDeElementos(ByVal CodCompania As Integer, ByVal TipoElemento As Elemento.TiposDeElemento) As Dictionary(Of String, ElementoDatosadicionales)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getDatosAdicionalesDeElementos_online(CodCompania:=CodCompania, TipoElemento:=TipoElemento)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getElementosConDatosAdicionales_offline(CodCompania:=CodCompania, TipoElemento:=TipoElemento).datosadicionales
        End If
        Return New Dictionary(Of String, ElementoDatosadicionales)
    End Function

    Private Function getDatosAdicionalesDeElementos_online(ByVal CodCompania As Integer, ByVal TipoElemento As Elemento.TiposDeElemento) As Dictionary(Of String, ElementoDatosadicionales)
        Dim lista As Dictionary(Of String, ElementoDatosadicionales)
        lista = New Dictionary(Of String, ElementoDatosadicionales)

        Dim parteInicial As String
        If TipoElemento = Elemento.TiposDeElemento.articulo Then
            parteInicial = "SELECT a.CodArticulo, " & _
                           "       a.Peso, a.CodUnidadPeso, " & _
                           "       a.AreaEncofrado, a.UnidadAreaEncofrado, " & _
                           "       a.Material, " & _
                           "       a.UnidadStock, "
        Else
            parteInicial = "SELECT d.CodElemento, "
        End If
        ''//Nota: parece que el material bueno es a.Material; pero esta tambien d.CodMaterial, que no se que es.

        Dim parteFinal As String
        If TipoElemento = Elemento.TiposDeElemento.articulo Then
            parteFinal = "  FROM dbo.Articulos a  " & _
                         "  LEFT JOIN dbo.ElementosDatosAdicionales d ON d.CodElemento = a.CodArticulo " & _
                         "                                           AND d.CodCompania = a.CodCompania " & _
                         " WHERE a.CodCompania = " & CodCompania & _
                         zzzCriteriosDeConsultaSQLParaCogerSoloArticulosUtilizables & _
                         " ORDER by a.CodArticulo"
        Else
            parteFinal = "  FROM dbo.ElementosDatosAdicionales d  " & _
                         " WHERE d.CodCompania = " & CodCompania & _
                         "   AND d.TipoElemento = " & TipoElemento & _
                         " ORDER by d.CodElemento"
        End If


        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = parteInicial & _
                            "       d.CodMaterial, d.Generico, " & _
                            "       d.Longitud, d.CodUnidadLongitud, " & _
                            "       d.Anchura, d.CodUnidadAnchura, " & _
                            "       d.Area, d.CodUnidadArea, " & _
                            "       d.Altura, d.CodUnidadAltura, " & _
                            "       d.AlturaMin, d.CodUnidadAlturaMin, " & _
                            "       d.AlturaMax, d.CodUnidadAlturaMax, " & _
                            "       d.Espesor, d.CodUnidadEspesor, " & _
                            "       d.Cortante, d.CodUnidadCortante, " & _
                            "       d.Inercia, d.CodUnidadInercia, " & _
                            "       d.Momento, d.CodUnidadMomento, " & _
                            "       d.Densidad, d.CodUnidadDensidad, " & _
                            "       d.ResisFlexionPerpen, d.CodUnidadResisFlexionPerpen, " & _
                            "       d.ResisFlexionParal, d.CodUnidadResisFlexionParal, " & _
                            "       d.ModuloElasticoMedioPerpen, d.CodUnidadModuloElasticoMedioPerpen, " & _
                            "       d.ModuloElasticoMedioParal, d.CodUnidadModuloElasticoMedioParal, " & _
                            "       d.DiametroTuboInterior, d.CodUnidadDiametroTuboInt, " & _
                            "       d.ModuloYoung, d.CodUnidadModuloYoung, " & _
                            "       d.AreaCortante, d.CodUnidadAreaCortante, " & _
                            "       d.ModuloResistente, d.CodUnidadModuloResistente, " & _
                            "       d.ResisCortantePerpend, d.CodUnidadResisCortantePerpend, " & _
                            "       d.ResisCortanteParal, d.CodUnidadResisCortanteParal, " & _
                            "       d.ModuloRigidezMedioPerpend, d.CodUnidadModuloRigidezMedioPerpend, " & _
                            "       d.ModuloRigidezMedioParalelo, d.CodUnidadModuloRigidezMedioParalelo " & _
                            parteFinal

        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim datos As ElementoDatosadicionales
                    datos = New ElementoDatosadicionales

                    If TipoElemento = Elemento.TiposDeElemento.articulo Then

                        If Not IsDBNull(reader(reader.GetOrdinal("CodArticulo"))) Then
                            datos.CodElemento = reader.GetString(reader.GetOrdinal("CodArticulo")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("Peso"))) Then
                            datos.Peso.valor = reader.GetDouble(reader.GetOrdinal("Peso"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadPeso"))) Then
                            datos.Peso.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadPeso")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("AreaEncofrado"))) Then
                            datos.AreaEncofrado.valor = reader.GetDouble(reader.GetOrdinal("AreaEncofrado"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("UnidadAreaEncofrado"))) Then
                            datos.AreaEncofrado.unidaddemedida = reader.GetString(reader.GetOrdinal("UnidadAreaEncofrado")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("Material"))) Then
                            datos.Material = reader.GetString(reader.GetOrdinal("Material")).Trim
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("UnidadStock"))) Then
                            datos.UnidadStock = reader.GetString(reader.GetOrdinal("UnidadStock")).Trim
                        End If

                    Else

                        If Not IsDBNull(reader(reader.GetOrdinal("CodElemento"))) Then
                            datos.CodElemento = reader.GetString(reader.GetOrdinal("CodElemento")).Trim
                        End If

                    End If

                    'If Not IsDBNull(reader(reader.GetOrdinal("CodMaterial"))) Then
                    '    datos.CodMaterial = reader(reader.GetOrdinal("CodMaterial")).Trim
                    'End If


                    If Not IsDBNull(reader(reader.GetOrdinal("Generico"))) Then
                        datos.Generico = reader.GetString(reader.GetOrdinal("Generico")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Longitud"))) Then
                        datos.Longitud.valor = reader.GetDecimal(reader.GetOrdinal("Longitud"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadLongitud"))) Then
                        datos.Longitud.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadLongitud")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Anchura"))) Then
                        datos.Anchura.valor = reader.GetDecimal(reader.GetOrdinal("Anchura"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAnchura"))) Then
                        datos.Anchura.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAnchura")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Area"))) Then
                        datos.Area.valor = reader.GetDecimal(reader.GetOrdinal("Area"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadArea"))) Then
                        datos.Area.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadArea")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Altura"))) Then
                        datos.Altura.valor = reader.GetDecimal(reader.GetOrdinal("Altura"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAltura"))) Then
                        datos.Altura.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAltura")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("AlturaMin"))) Then
                        datos.AlturaMin.valor = reader.GetDecimal(reader.GetOrdinal("AlturaMin"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAlturaMin"))) Then
                        datos.AlturaMin.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAlturaMin")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("AlturaMax"))) Then
                        datos.AlturaMax.valor = reader.GetDecimal(reader.GetOrdinal("AlturaMax"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAlturaMax"))) Then
                        datos.AlturaMax.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAlturaMax")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Espesor"))) Then
                        datos.Espesor.valor = reader.GetDecimal(reader.GetOrdinal("Espesor"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadEspesor"))) Then
                        datos.Espesor.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadEspesor")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Cortante"))) Then
                        datos.Cortante.valor = reader.GetDecimal(reader.GetOrdinal("Cortante"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadCortante"))) Then
                        datos.Cortante.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadCortante")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Inercia"))) Then
                        datos.Inercia.valor = reader.GetDecimal(reader.GetOrdinal("Inercia"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadInercia"))) Then
                        datos.Inercia.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadInercia")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Momento"))) Then
                        datos.Momento.valor = reader.GetDecimal(reader.GetOrdinal("Momento"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadMomento"))) Then
                        datos.Momento.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadMomento")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Densidad"))) Then
                        datos.Densidad.valor = reader.GetDecimal(reader.GetOrdinal("Densidad"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadDensidad"))) Then
                        datos.Densidad.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadDensidad")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ResisFlexionPerpen"))) Then
                        datos.ResisFlexionPerpen.valor = reader.GetDecimal(reader.GetOrdinal("ResisFlexionPerpen"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadResisFlexionPerpen"))) Then
                        datos.ResisFlexionPerpen.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadResisFlexionPerpen")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ResisFlexionParal"))) Then
                        datos.ResisFlexionParal.valor = reader.GetDecimal(reader.GetOrdinal("ResisFlexionParal"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadResisFlexionParal"))) Then
                        datos.ResisFlexionParal.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadResisFlexionParal")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloElasticoMedioPerpen"))) Then
                        datos.ModuloElasticoMedioPerpen.valor = reader.GetDecimal(reader.GetOrdinal("ModuloElasticoMedioPerpen"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloElasticoMedioPerpen"))) Then
                        datos.ModuloElasticoMedioPerpen.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloElasticoMedioPerpen")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloElasticoMedioParal"))) Then
                        datos.ModuloElasticoMedioParal.valor = reader.GetDecimal(reader.GetOrdinal("ModuloElasticoMedioParal"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloElasticoMedioParal"))) Then
                        datos.ModuloElasticoMedioParal.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloElasticoMedioParal")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("DiametroTuboInterior"))) Then
                        datos.DiametroTuboInterior.valor = reader.GetDecimal(reader.GetOrdinal("DiametroTuboInterior"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadDiametroTuboInt"))) Then
                        datos.DiametroTuboInterior.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadDiametroTuboInt")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloYoung"))) Then
                        datos.ModuloYoung.valor = reader.GetDecimal(reader.GetOrdinal("ModuloYoung"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloYoung"))) Then
                        datos.ModuloYoung.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloYoung")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("AreaCortante"))) Then
                        datos.AreaCortante.valor = reader.GetDecimal(reader.GetOrdinal("AreaCortante"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAreaCortante"))) Then
                        datos.AreaCortante.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAreaCortante")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloResistente"))) Then
                        datos.ModuloResistente.valor = reader.GetDecimal(reader.GetOrdinal("ModuloResistente"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloResistente"))) Then
                        datos.ModuloResistente.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloResistente")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ResisCortantePerpend"))) Then
                        datos.ResisCortantePerpend.valor = reader.GetDecimal(reader.GetOrdinal("ResisCortantePerpend"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadResisCortantePerpend"))) Then
                        datos.ResisCortantePerpend.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadResisCortantePerpend")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ResisCortanteParal"))) Then
                        datos.ResisCortanteParal.valor = reader.GetDecimal(reader.GetOrdinal("ResisCortanteParal"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadResisCortanteParal"))) Then
                        datos.ResisCortanteParal.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadResisCortanteParal")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloRigidezMedioPerpend"))) Then
                        datos.ModuloRigidezMedioPerpend.valor = reader.GetDecimal(reader.GetOrdinal("ModuloRigidezMedioPerpend"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloRigidezMedioPerpend"))) Then
                        datos.ModuloRigidezMedioPerpend.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloRigidezMedioPerpend")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModuloRigidezMedioParalelo"))) Then
                        datos.ModuloRigidezMedioParalelo.valor = reader.GetDecimal(reader.GetOrdinal("ModuloRigidezMedioParalelo"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadModuloRigidezMedioParalelo"))) Then
                        datos.ModuloRigidezMedioParalelo.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadModuloRigidezMedioParalelo")).Trim
                    End If

                    lista.Add(key:=datos.CodElemento, value:=datos)
                End While

                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return lista

    End Function

    Private Class resultadosElementosConDatosAdicionales
        Friend elementos As Dictionary(Of String, Elemento)
        Friend datosadicionales As Dictionary(Of String, ElementoDatosadicionales)
    End Class
    Private Function getElementosConDatosAdicionales_offline(ByVal CodCompania As Integer, ByVal TipoElemento As Elemento.TiposDeElemento) As resultadosElementosConDatosAdicionales
        Dim resultados As New resultadosElementosConDatosAdicionales
        resultados.elementos = New Dictionary(Of String, Elemento)
        resultados.datosadicionales = New Dictionary(Of String, ElementoDatosadicionales)

        Dim nombreArchivo As String
        Dim etiquetaRaiz As String
        Dim etiquetaNodo As String
        Select Case TipoElemento
            Case Elemento.TiposDeElemento.articulo
                nombreArchivo = Elemento.zzXML_NombreArchivoArticulos
                etiquetaRaiz = Elemento.zzXML_EtiquetaRaizArticulos
                etiquetaNodo = Elemento.zzXML_EtiquetaNodoArticulo
            Case Elemento.TiposDeElemento.articuloGenerico
                nombreArchivo = Elemento.zzXML_NombreArchivoArticulosGenericos
                etiquetaRaiz = Elemento.zzXML_EtiquetaRaizArticulosGenericos
                etiquetaNodo = Elemento.zzXML_EtiquetaNodoArticuloGenerico
            Case Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                nombreArchivo = Elemento.zzXML_NombreArchivoBloquesAuxiliares
                etiquetaRaiz = Elemento.zzXML_EtiquetaRaizBloquesAuxiliares
                etiquetaNodo = Elemento.zzXML_EtiquetaNodoBloqueAuxiliar
            Case Else
                nombreArchivo = ""
                etiquetaRaiz = ""
                etiquetaNodo = ""
        End Select

        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & nombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nElemento As XmlNode
            For Each nElemento In docXML.SelectNodes(etiquetaRaiz & "/" & etiquetaNodo)
                Dim elemento As Elemento
                Select Case TipoElemento
                    Case ConsultarBDI.Elemento.TiposDeElemento.articulo
                        elemento = New Articulo()
                    Case ConsultarBDI.Elemento.TiposDeElemento.articuloGenerico
                        elemento = New ArticuloGenerico()
                    Case ConsultarBDI.Elemento.TiposDeElemento.bloqueAuxiliarAutocad
                        elemento = New BloqueAuxiliarAutocad()
                    Case Else
                        Elemento = New DummyElementoDesconocido()
                End Select
                elemento.fromXML(LeerElementoDesde:=nElemento)
                resultados.elementos.Add(elemento.CodElemento, elemento)
                Dim nDatosAdicionales As XmlNode
                nDatosAdicionales = nElemento.SelectSingleNode(ElementoDatosadicionales.zzXMLEtiquetaNodoDatosadicionales)
                If Not IsNothing(nDatosAdicionales) Then
                    Dim datos As New ElementoDatosadicionales
                    datos.fromXML(LeerDatosDesde:=nDatosAdicionales, CodElementoAlQueAsignarlos:=elemento.CodElemento, TipoDeElemento:=elemento.TipoElemento)
                    resultados.datosadicionales.Add(datos.CodElemento, datos)
                End If
            Next
        End If

        Return resultados
    End Function



    Private Function getEstructuraConfiguracionesArticulos_online(ByVal CodCompania As Integer) As List(Of EstructuraConfiguracionArticulo)
        Dim lista As List(Of EstructuraConfiguracionArticulo)
        lista = New List(Of EstructuraConfiguracionArticulo)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT CodArticulo, Orden, Cantidad, CodUnidadCantidad, ModoSuministro, CodConfiguracion" & _
                            "  FROM ArticulosConfiguracion" & _
                            " WHERE CodCompania = " & CodCompania & _
                            " ORDER by CodConfiguracion, Orden"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim articulo As EstructuraConfiguracionArticulo
                    articulo = New EstructuraConfiguracionArticulo

                    If Not IsDBNull(reader(reader.GetOrdinal("CodArticulo"))) Then
                        articulo.CodArticulo = reader.GetString(reader.GetOrdinal("CodArticulo")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Orden"))) Then
                        articulo.NumOrdenParaListar = reader.GetInt32(reader.GetOrdinal("Orden"))
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("Cantidad"))) Then
                        articulo.Cantidad.valor = reader.GetDecimal(reader.GetOrdinal("Cantidad"))
                    End If
                    If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadCantidad"))) Then
                        articulo.Cantidad.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadCantidad")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("ModoSuministro"))) Then
                        articulo.CodModoSuministro = reader.GetString(reader.GetOrdinal("ModoSuministro")).Trim
                    End If

                    If Not IsDBNull(reader(reader.GetOrdinal("CodConfiguracion"))) Then
                        articulo.CodConfiguracionALaQuePertenece = reader.GetString(reader.GetOrdinal("CodConfiguracion")).Trim
                    End If

                    lista.Add(articulo)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try


        Return lista
    End Function

    Private Class resultadosConfiguracionArticulo
        Friend configuraciones As Dictionary(Of String, Elemento)
        Friend configuracionesarticulos As List(Of EstructuraConfiguracionArticulo)
    End Class
    Private Function getEstructuraConfiguracionesArticulos_offline(ByVal CodCompania As Integer) As resultadosConfiguracionArticulo
        Dim resultados As New resultadosConfiguracionArticulo
        resultados.configuraciones = New Dictionary(Of String, Elemento)
        resultados.configuracionesarticulos = New List(Of EstructuraConfiguracionArticulo)

        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & EstructuraConfiguracionArticulo.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nConfiguracion As XmlNode
            For Each nConfiguracion In docXML.SelectNodes(EstructuraConfiguracionArticulo.zzXML_EtiquetaRaizConfiguraciones & _
                                                     "/" & Elemento.zzXML_EtiquetaNodoConfiguracion)
                Dim configuracion As New Configuracion()
                configuracion.fromXML(LeerElementoDesde:=nConfiguracion)
                resultados.configuraciones.Add(configuracion.CodElemento, configuracion)
                Dim nArticulos As XmlNode
                nArticulos = nConfiguracion.SelectSingleNode(EstructuraConfiguracionArticulo.zzXML_EtiquetaNodoConfiguracionArticulos)
                If Not IsNothing(nArticulos) Then
                    Dim nArticulo As XmlNode
                    For Each nArticulo In nArticulos.ChildNodes
                        Dim ca As New EstructuraConfiguracionArticulo
                        ca.fromXML(LeerArticuloDesde:=nArticulo, CodConfiguracionALaQueAsignarlo:=configuracion.CodElemento)
                        resultados.configuracionesarticulos.Add(ca)
                    Next
                End If
            Next
        End If

        Return resultados
    End Function






    Public Function getTarifas(ByVal CodCompania As Integer) As List(Of Tarifa)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getTarifas_online(CodCompania:=CodCompania, listaDeCodigosDeTarifaARecuperar:="")
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getTarifas_offline(CodCompania:=CodCompania)
        End If
        Return New List(Of Tarifa)
    End Function
    Public Function getTarifas(ByVal CodCompania As Integer, ByVal listaDeCodigosDeTarifaARecuperar As String) As List(Of Tarifa)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getTarifas_online(CodCompania:=CodCompania, listaDeCodigosDeTarifaARecuperar:=listaDeCodigosDeTarifaARecuperar)
        Else
            Return New List(Of Tarifa)
        End If
    End Function

    Public Function getPrecios(ByVal CodCompania As Integer, ByVal CodTarifa As String, ByVal limitarASoloArticulosPresentesEnAlgunGrupoSubgrupo As Boolean) As Dictionary(Of String, Double)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getPrecios_online(CodCompania:=CodCompania, CodTarifa:=CodTarifa, limitarASoloArticulosPresentesEnAlgunGrupoSubgrupo:=limitarASoloArticulosPresentesEnAlgunGrupoSubgrupo)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getPrecios_offline(CodCompania:=CodCompania, CodTarifa:=CodTarifa)
        End If
        Return New Dictionary(Of String, Double)
    End Function


    Private Function getTarifas_online(ByVal CodCompania As Integer, ByVal listaDeCodigosDeTarifaARecuperar As String) As List(Of Tarifa)
        Dim ListaTarifas As New List(Of Tarifa)

        Dim condicionExtra As String
        If listaDeCodigosDeTarifaARecuperar = "" Then
            condicionExtra = ""
        Else
            condicionExtra = "   AND CodTarifa IN (" & listaDeCodigosDeTarifaARecuperar & ")"
        End If

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT CodTarifa, Descripcion, CodDivisa" _
                        & "  FROM dbo.Tarifas" _
                        & " WHERE CodCompania = " & CodCompania _
                        & condicionExtra _
                        & " ORDER BY CodTarifa"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim tarifa As New Tarifa
                    If Not IsDBNull(reader(reader.GetOrdinal("CodTarifa"))) Then
                        ''// ¡ojo!, no hacer un .Trim del codigo porque en la BD es un char(3) 
                        ''//        y  luego necesitamos los espacios en blanco a la hora de buscar los precios
                        ''//        correspondientes a cada tarifa.
                        tarifa.CodTarifa = reader.GetString(reader.GetOrdinal("CodTarifa"))
                        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion"))) Then
                            tarifa.NombreTarifa = reader.GetString(reader.GetOrdinal("Descripcion")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodDivisa"))) Then
                            tarifa.CodMoneda = reader.GetString(reader.GetOrdinal("CodDivisa")).Trim
                        End If
                    End If
                    ListaTarifas.Add(tarifa)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return ListaTarifas
    End Function

    Private Function getTarifas_offline(ByVal CodCompania As Integer) As List(Of Tarifa)
        Dim tarifas As New List(Of Tarifa)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & Tarifa.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nTarifa As XmlNode
            For Each nTarifa In docXML.SelectNodes(tarifa.zzXML_EtiquetaRaiz & "/" & tarifa.zzXML_EtiquetaNodo)
                Dim tarifa As New Tarifa
                tarifa.fromXML(LeerTarifaDesde:=nTarifa)
                tarifas.Add(tarifa)
            Next
        End If
        Return tarifas
    End Function

    Private Function getPrecios_online(ByVal CodCompania As Integer, ByVal CodTarifa As String, ByVal limitarASoloArticulosPresentesEnAlgunGrupoSubgrupo As Boolean) As Dictionary(Of String, Double)
        Dim ListaPrecios As New Dictionary(Of String, Double)

        Dim condicionExtra As String = ""
        If limitarASoloArticulosPresentesEnAlgunGrupoSubgrupo Then
            condicionExtra = "   AND CodArticulo IN (SELECT distinct CodElemento " & _
                             "                         FROM dbo.ElementosGruposSubgrupos e " & _
                             "                        WHERE e.CodCompania = " & CodCompania & _
                             "                          AND e.TipoElemento = " & Elemento.TiposDeElemento.articulo & ") "
        End If

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT CodArticulo, Precio" & _
                            "  FROM dbo.ArticuloTarifas" & _
                            " WHERE CodCompania = " & CodCompania & " " & _
                            "   AND CodTarifa = '" & CodTarifa & "' " & _
                            condicionExtra & _
                            " ORDER BY CodArticulo"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim codigo As String
                    Dim precio As Double
                    If Not IsDBNull(reader(reader.GetOrdinal("CodArticulo"))) Then
                        codigo = reader.GetString(reader.GetOrdinal("CodArticulo")).Trim
                        If Not IsDBNull(reader(reader.GetOrdinal("Precio"))) Then
                            precio = reader.GetDecimal(reader.GetOrdinal("Precio"))
                        End If
                        Try
                            ListaPrecios.Add(key:=codigo, value:=precio)
                        Catch ex As ArgumentException
                            RegistrosyMensajes.getMensajero.DarAviso(Nivel:=RegistrosyMensajes.nivel.n4_WARNING, _
                                                                     Mensaje:="- CodCompania " & CodCompania & " - CodArticulo " & codigo & " - " & Environment.NewLine & ex.Message, _
                                                                     Procedencia:=ex.StackTrace)
                            Continue While
                        End Try
                    End If
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return ListaPrecios
    End Function

    Private Function getPrecios_offline(ByVal CodCompania As Integer, ByVal CodTarifa As String) As Dictionary(Of String, Double)
        Dim precios As New Dictionary(Of String, Double)

        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & Tarifa.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then

            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nTarifa As XmlNode
            For Each nTarifa In docXML.SelectNodes(tarifa.zzXML_EtiquetaRaiz & "/" & tarifa.zzXML_EtiquetaNodo)
                Dim tarifa As New Tarifa
                tarifa.fromXML(LeerTarifaDesde:=nTarifa)

                If tarifa.CodTarifa = CodTarifa Then
                    Dim nPrecios As XmlNodeList
                    nPrecios = nTarifa.SelectNodes(tarifa.zzXML_EtiquetaRaizArticulos & "/" & tarifa.zzXML_EtiquetaNodoArticulo)
                    If Not IsNothing(nPrecios) Then
                        Dim nArticulo As XmlNode
                        For Each nArticulo In nPrecios
                            Dim nCodigo As XmlNode
                            nCodigo = nArticulo.SelectSingleNode(tarifa.zzEtiquetaCodigoArticulo)
                            Dim nPrecio As XmlNode
                            nPrecio = nArticulo.SelectSingleNode(tarifa.zzEtiquetaPrecio)
                            If Not IsNothing(nCodigo) And Not IsNothing(nPrecio) Then
                                precios.Add(key:=nCodigo.InnerText, value:=Double.Parse(nPrecio.InnerText, Globalization.CultureInfo.InvariantCulture))
                            End If
                        Next
                    End If
                End If

            Next

        End If

        Return precios
    End Function





    Public Function getListasDeCargasDeUso(ByVal CodCompania As Integer) As List(Of ListaDeCargasDeUso)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getListasDeCargasDeUso_online(CodCompania:=CodCompania)
        ElseIf IsNothing(m_CadenaConexionSQLServer) Then
            Return getListasDeCargasDeUso_offline(CodCompania:=CodCompania)
        Else
            Return New List(Of ListaDeCargasDeUso)
        End If
    End Function

    Private Class DatoDeCargas : Inherits DatoDeCargasDeUso
        Public CodNorma As String
        Public Function getClaveUnica() As String
            Return getClaveUnica(CodNorma, TipoElemento, CodElemento, TuboInterior.ToString(), Altura)
        End Function
        Public Shared Function getClaveUnica(ByVal CodNorma As String, _
                                             ByVal TipoElemento As Elemento.TiposDeElemento, _
                                             ByVal CodElemento As String, _
                                             ByVal TuboInterior As String, _
                                             ByVal Altura As Magnitud) _
                                          As String
            Return CodNorma & "-" & TipoElemento & "-" & CodElemento & "-" & TuboInterior.ToString & "-" & Altura.ToStringConCulturaNeutra()
        End Function
    End Class
    Private Function getListasDeCargasDeUso_online(ByVal CodCompania As Integer) As List(Of ListaDeCargasDeUso)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120

        cmd.CommandText = "SELECT CodNorma, TipoElemento, CodElemento," & _
                            "       Altura, CodUnidadAltura," & _
                            "       TuboInterior," & _
                            "       Carga, CodUnidadCarga," & _
                            "       CargaPlantaHormigonada, CodUnidadCargaPlantaHormigonada," & _
                            "       CargaReapuntalado, CodUnidadCargaReapuntalado," & _
                            "       Viento," & _
                            "       Rigidez, CodUnidadRigidez" & _
                            "  FROM dbo.CargasDeUso" & _
                            " WHERE CodCompania = " & CodCompania & " " & _
                            " ORDER BY CodNorma, TipoElemento, CodElemento, Altura, TuboInterior"

        Dim cargasLeidas As New SortedDictionary(Of String, DatoDeCargas)
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()

                While reader.Read()
                    Dim cargaDeuso As DatoDeCargas
                    cargaDeuso = New DatoDeCargas
                    If Not IsDBNull(reader(reader.GetOrdinal("CodNorma"))) _
                    And Not IsDBNull(reader(reader.GetOrdinal("TipoElemento"))) _
                    And Not IsDBNull(reader(reader.GetOrdinal("CodElemento"))) Then

                        cargaDeuso.CodNorma = reader.GetString(reader.GetOrdinal("CodNorma")).Trim

                        Try
                            cargaDeuso.TipoElemento = CType([Enum].Parse(enumType:=GetType(Elemento.TiposDeElemento), _
                                                                         value:=reader.GetInt16(reader.GetOrdinal("TipoElemento")).ToString.Trim),  _
                                                            Elemento.TiposDeElemento)
                        Catch ex As ArgumentException
                            cargaDeuso.TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
                        End Try

                        cargaDeuso.CodElemento = reader.GetString(reader.GetOrdinal("CodElemento")).Trim

                        If Not IsDBNull(reader(reader.GetOrdinal("Altura"))) Then
                            cargaDeuso.Altura.valor = reader.GetDecimal(reader.GetOrdinal("Altura"))
                            If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAltura"))) Then
                                cargaDeuso.Altura.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAltura")).Trim
                            End If
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("Viento"))) Then
                            cargaDeuso.Viento = reader.GetDecimal(reader.GetOrdinal("Viento"))
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("TuboInterior"))) Then
                            cargaDeuso.TuboInterior = CType([Enum].Parse(GetType(DatoDeCargasDeUso.OrientacionDelTuboInterior), reader.GetString(reader.GetOrdinal("TuboInterior"))), DatoDeCargasDeUso.OrientacionDelTuboInterior)
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("Carga"))) Then
                            cargaDeuso.Carga.valor = reader.GetDecimal(reader.GetOrdinal("Carga"))
                            If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadCarga"))) Then
                                cargaDeuso.Carga.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadCarga")).Trim
                            End If
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("CargaPlantaHormigonada"))) Then
                            cargaDeuso.CargaPlantaHormigonada.valor = reader.GetDecimal(reader.GetOrdinal("CargaPlantaHormigonada"))
                            If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadCargaPlantaHormigonada"))) Then
                                cargaDeuso.CargaPlantaHormigonada.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadCargaPlantaHormigonada")).Trim
                            End If
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("CargaReapuntalado"))) Then
                            cargaDeuso.CargaReapuntalado.valor = reader.GetDecimal(reader.GetOrdinal("CargaReapuntalado"))
                            If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadCargaReapuntalado"))) Then
                                cargaDeuso.CargaReapuntalado.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadCargaReapuntalado")).Trim
                            End If
                        End If

                        If Not IsDBNull(reader(reader.GetOrdinal("Rigidez"))) Then
                            cargaDeuso.Rigidez.valor = reader.GetDecimal(reader.GetOrdinal("Rigidez"))
                            If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadRigidez"))) Then
                                cargaDeuso.Rigidez.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadRigidez")).Trim
                            End If
                        End If

                        Try
                            cargasLeidas.Add(key:=cargaDeuso.getClaveUnica, value:=cargaDeuso)
                        Catch ex As ArgumentException
                            Throw New ArgumentException(cargaDeuso.getClaveUnica)
                        End Try

                    End If
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Dim listas As New List(Of ListaDeCargasDeUso)
        If cargasLeidas.Count > 0 Then
            Dim cargas As New List(Of DatoDeCargasDeUso)
            Dim listaAnterior As String = cargasLeidas.Values(0).CodNorma
            Dim carga As DatoDeCargas
            For Each carga In cargasLeidas.Values
                If carga.CodNorma <> listaAnterior Then
                    listas.Add(New ListaDeCargasDeUso(codNorma:=listaAnterior, cargas:=cargas))
                    cargas = New List(Of DatoDeCargasDeUso)
                End If
                cargas.Add(carga)
                listaAnterior = carga.CodNorma
            Next
            listas.Add(New ListaDeCargasDeUso(codNorma:=listaAnterior, cargas:=cargas))
        End If
        Return listas

    End Function

    Private Function getListasDeCargasDeUso_offline(ByVal CodCompania As Integer) As List(Of ListaDeCargasDeUso)
        Dim listas As New List(Of ListaDeCargasDeUso)

        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & ListaDeCargasDeUso.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nLista As XmlNode
            For Each nLista In docXML.ChildNodes(1).ChildNodes
                listas.Add(New ListaDeCargasDeUso(nLista))
            Next
        End If

        Return listas
     End Function





    Public Function getFamiliasRevit(ByVal CodCompania As Integer) As List(Of FamiliaRevit)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getFamiliasRevit_online(CodCompania:=CodCompania)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getFamiliasRevit_offline(CodCompania:=CodCompania)
        End If
        Return New List(Of FamiliaRevit)
    End Function

    Private Function getFamiliasRevit_online(ByVal CodCompania As Integer) As List(Of FamiliaRevit)
        Dim ListaFamiliasRevit As New List(Of FamiliaRevit)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos; nosotros ponemos 120 segundos; si se necesita mas, se puede aumentar

        cmd.CommandText = "SELECT CodFamiliaRevit, " _
                        & "       Descripcion, " _
                        & "       EsDinamica, " _
                        & "       EsSET, " _
                        & "       EsAnnotationSymbol, " _
                        & "       NombreFichero " _
                        & "  FROM dbo.FamiliasRevit " _
                        & " WHERE CodCompania = " & CodCompania _
                        & " ORDER BY CodFamiliaRevit"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim familiaRevit As New FamiliaRevit
                    If Not IsDBNull(reader(reader.GetOrdinal("CodFamiliaRevit"))) Then
                        familiaRevit.CodElemento = reader.GetString(reader.GetOrdinal("CodFamiliaRevit"))
                        If Not IsDBNull(reader(reader.GetOrdinal("Descripcion"))) Then
                            familiaRevit.NombreLocal = reader.GetString(reader.GetOrdinal("Descripcion")).Trim
                            familiaRevit.NAME_en = reader.GetString(reader.GetOrdinal("Descripcion")).Trim
                            familiaRevit.NAME_es = reader.GetString(reader.GetOrdinal("Descripcion")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("EsDinamica"))) Then
                            familiaRevit.esDinamica = reader.GetBoolean(reader.GetOrdinal("EsDinamica"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("EsSET"))) Then
                            familiaRevit.esConjunto = reader.GetBoolean(reader.GetOrdinal("EsSET"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("EsAnnotationSymbol"))) Then
                            familiaRevit.EsAnnotationSymbol = reader.GetBoolean(reader.GetOrdinal("EsAnnotationSymbol"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("NombreFichero"))) Then
                            familiaRevit.nombreFichero = reader.GetString(reader.GetOrdinal("NombreFichero")).Trim
                        End If
                    End If
                    ListaFamiliasRevit.Add(familiaRevit)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return ListaFamiliasRevit
    End Function

    Private Function getFamiliasRevit_offline(ByVal CodCompania As Integer) As List(Of FamiliaRevit)
        Dim FamiliasRevit As New List(Of FamiliaRevit)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & FamiliaRevit.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nFamiliaRevit As XmlNode
            For Each nFamiliaRevit In docXML.SelectNodes(FamiliaRevit.zzXML_EtiquetaRaiz & "/" & FamiliaRevit.zzXML_EtiquetaNodo)
                Dim FamiliaRevit As New FamiliaRevit
                FamiliaRevit.fromXML(nFamiliaRevit)
                FamiliasRevit.Add(FamiliaRevit)
            Next
        End If
        Return FamiliasRevit
    End Function



    Public Function getFamiliasDinamicasRevit(ByVal CodCompania As Integer) As List(Of FamiliaDinamicaRevit)
        If IsNothing(m_CarpetaArchivosXML) Then
            Return getFamiliasDinamicasRevit_online(CodCompania:=CodCompania)
        End If
        If IsNothing(m_CadenaConexionSQLServer) Then
            Return getFamiliasDinamicasRevit_offline(CodCompania:=CodCompania)
        End If
        Return New List(Of FamiliaDinamicaRevit)
    End Function

    Private Function getFamiliasDinamicasRevit_online(ByVal CodCompania As Integer) As List(Of FamiliaDinamicaRevit)
        Dim ListaFamiliaDinamicas As New List(Of FamiliaDinamicaRevit)

        Dim cmd As New SqlCommand()
        cmd.CommandTimeout = 120 '//por defecto son 30 segundos

        cmd.CommandText = "SELECT FAMILY_CODE, " _
                        & "       Longitud, CodUnidadLongitud, " _
                        & "       Anchura, CodUnidadAnchura, " _
                        & "       Altura, CodUnidadAltura, " _
                        & "       TipoGenerico, CodElemento, " _
                        & "       d.CodFamiliaRevit, f.NombreFichero" _
                        & "  FROM dbo.FamiliasRevitDinamicas d" _
                        & "  LEFT JOIN dbo.FamiliasRevit f ON f.CodCompania = d.CodCompania" _
                        & "                               AND f.CodFamiliaRevit = d.CodFamiliaRevit" _
                        & " WHERE d.CodCompania = " & CodCompania _
                        & " ORDER BY FAMILY_CODE, TipoGenerico, Longitud, Anchura, Altura"
        Try
            Using conexion As New SqlConnection(m_CadenaConexionSQLServer)
                conexion.Open()
                cmd.Connection = conexion
                Dim reader As SqlDataReader
                reader = cmd.ExecuteReader()
                While reader.Read()
                    Dim familiaDinamica As New FamiliaDinamicaRevit
                    If Not IsDBNull(reader(reader.GetOrdinal("FAMILY_CODE"))) Then
                        familiaDinamica.FAMILY_CODE = reader.GetString(reader.GetOrdinal("FAMILY_CODE"))
                        If Not IsDBNull(reader(reader.GetOrdinal("Longitud"))) Then
                            familiaDinamica.longitud.valor = reader.GetDecimal(reader.GetOrdinal("Longitud"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadLongitud"))) Then
                            familiaDinamica.longitud.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadLongitud")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("Anchura"))) Then
                            familiaDinamica.anchura.valor = reader.GetDecimal(reader.GetOrdinal("Anchura"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAnchura"))) Then
                            familiaDinamica.anchura.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAnchura")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("Altura"))) Then
                            familiaDinamica.altura.valor = reader.GetDecimal(reader.GetOrdinal("Altura"))
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodUnidadAltura"))) Then
                            familiaDinamica.altura.unidaddemedida = reader.GetString(reader.GetOrdinal("CodUnidadAltura")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("TipoGenerico"))) Then
                            familiaDinamica.tipoGenerico = CType([Enum].Parse(GetType(FamiliaDinamicaRevit.TiposDeGenerico), reader.GetString(reader.GetOrdinal("TipoGenerico"))), FamiliaDinamicaRevit.TiposDeGenerico)
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodElemento"))) Then
                            familiaDinamica.CodElementoAlQueRepresenta = reader.GetString(reader.GetOrdinal("CodElemento")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("CodFamiliaRevit"))) Then
                            familiaDinamica.CodElemento = reader.GetString(reader.GetOrdinal("CodFamiliaRevit")).Trim
                        End If
                        If Not IsDBNull(reader(reader.GetOrdinal("NombreFichero"))) Then
                            familiaDinamica.nombreFichero = reader.GetString(reader.GetOrdinal("NombreFichero")).Trim
                        End If
                    End If
                    ListaFamiliaDinamicas.Add(familiaDinamica)
                End While
                reader.Close()
                conexion.Close()
            End Using
        Catch
        End Try

        Return ListaFamiliaDinamicas
    End Function

    Private Function getFamiliasDinamicasRevit_offline(ByVal CodCompania As Integer) As List(Of FamiliaDinamicaRevit)
        Dim FamiliasDinamicas As New List(Of FamiliaDinamicaRevit)
        Dim archivoXML As String
        archivoXML = m_CarpetaArchivosXML.FullName & IO.Path.DirectorySeparatorChar & CodCompania & "_" & FamiliaDinamicaRevit.zzXML_NombreArchivo & ".xml"
        If IO.File.Exists(archivoXML) Then
            Dim docXML As Xml.XmlDocument
            docXML = New XmlDocument()
            docXML.Load(archivoXML)
            Dim nFamiliaDinamica As XmlNode
            For Each nFamiliaDinamica In docXML.SelectNodes(FamiliaDinamicaRevit.zzXML_EtiquetaRaiz & "/" & FamiliaDinamicaRevit.zzXML_EtiquetaNodo)
                Dim FamiliaDinamica As New FamiliaDinamicaRevit
                FamiliaDinamica.fromXML(nFamiliaDinamica)
                FamiliasDinamicas.Add(FamiliaDinamica)
            Next
        End If
        Return FamiliasDinamicas
    End Function




End Class
