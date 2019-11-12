Imports System.Xml



Public MustInherit Class Elemento : Implements IEquatable(Of Object) : Implements IComparable(Of Elemento)


    Public Enum TiposDeElemento As Integer
        DESCONOCIDO = 0
        articulo = 1
        configuracion = 2
        articuloGenerico = 3
        bloqueAuxiliarAutocad = 4
        ''Nota: los grupos y subgrupos no son elementos, pero comparten muchos campos con ellos.
        grupo = 901
        subgrupo = 902
        ''Nota: Las familias Revit son muy distintas de los elementos (de hecho, se definen en una clase totalmente independiente).
        ''      Pero necesitamos un tipo de elemento para ellas, para incluirlas en la estructura grupo-subgrupo-elemento.
        familiaRevit = 5
    End Enum


    Public TipoElemento As TiposDeElemento
    Public CodElemento As String
    Public NombreLocal As String
    Public NAME_es As String
    Public NAME_en As String
    Public NAME_fr As String
    Public NAME_esES As String
    Public NAME_esPE As String
    Public NAME_esCL As String
    Public NAME_esMX As String
    Public NAME_frFR As String
    Public NAME_ptPT As String
    Public NAME_deDE As String
    Public NAME_plPL As String
    Public NAME_enUS As String
    Public NAME_itIT As String
    Public NAME_czCZ As String
    Public NAME_skSK As String
    Public NAME_roRO As String

    Public Sub New()
        Reset()
    End Sub


    Private Sub Reset()
        CodElemento = ""
        NombreLocal = ""
        NAME_es = ""
        NAME_en = ""
        NAME_fr = ""
        NAME_es = ""
        NAME_esES = ""
        NAME_esPE = ""
        NAME_esCL = ""
        NAME_esMX = ""
        NAME_frFR = ""
        NAME_ptPT = ""
        NAME_deDE = ""
        NAME_plPL = ""
        NAME_enUS = ""
        NAME_itIT = ""
        NAME_czCZ = ""
        NAME_skSK = ""
        NAME_roRO = ""
    End Sub

    Public Sub New(ByVal otroElemento As Elemento)
        TipoElemento = otroElemento.TipoElemento
        CodElemento = otroElemento.CodElemento
        NombreLocal = otroElemento.NombreLocal
        NAME_es = otroElemento.NAME_es
        NAME_en = otroElemento.NAME_en
        NAME_fr = otroElemento.NAME_fr
        NAME_esES = otroElemento.NAME_esES
        NAME_esPE = otroElemento.NAME_esPE
        NAME_esCL = otroElemento.NAME_esCL
        NAME_esMX = otroElemento.NAME_esMX
        NAME_frFR = otroElemento.NAME_frFR
        NAME_ptPT = otroElemento.NAME_ptPT
        NAME_deDE = otroElemento.NAME_deDE
        NAME_plPL = otroElemento.NAME_plPL
        NAME_enUS = otroElemento.NAME_enUS
        NAME_itIT = otroElemento.NAME_itIT
        NAME_czCZ = otroElemento.NAME_czCZ
        NAME_skSK = otroElemento.NAME_skSK
        NAME_roRO = otroElemento.NAME_roRO
    End Sub


    Public Function getNombre(idioma As String) As String
        Dim nombre As String = ""

        Select Case idioma
            Case "es"
                nombre = NAME_es
            Case "en"
                nombre = NAME_en
            Case "fr"
                nombre = NAME_fr
            Case "es-ES"
                If NAME_esES = "" Then
                    nombre = NAME_es
                Else
                    nombre = NAME_esES
                End If
            Case "es-PE"
                If NAME_esPE = "" Then
                    nombre = NAME_es
                Else
                    nombre = NAME_esPE
                End If
            Case "es-CL"
                If NAME_esCL = "" Then
                    nombre = NAME_es
                Else
                    nombre = NAME_esCL
                End If
            Case "es-MX"
                If NAME_esMX = "" Then
                    nombre = NAME_es
                Else
                    nombre = NAME_esMX
                End If
            Case "fr-FR"
                If NAME_frFR = "" Then
                    nombre = NAME_fr
                Else
                    nombre = NAME_frFR
                End If
            Case "pt-PT"
                nombre = NAME_ptPT
            Case "de-DE"
                nombre = NAME_deDE
            Case "pl-PL"
                nombre = NAME_plPL
            Case "en-US"
                If NAME_enUS = "" Then
                    nombre = NAME_en
                Else
                    nombre = NAME_enUS
                End If
            Case "it-IT"
                nombre = NAME_itIT
            Case "cz-CZ"
                nombre = NAME_czCZ
            Case "sk-SK"
                nombre = NAME_skSK
            Case "ro-RO"
                nombre = NAME_roRO
            Case Else
                nombre = NAME_en
        End Select

        If nombre = "" Then nombre = NAME_en
        Return nombre
    End Function


    Public Const zzXML_NombreArchivoArticulos As String = "articles"
    Public Const zzXML_EtiquetaRaizArticulos As String = "articles"
    Public Const zzXML_EtiquetaNodoArticulo As String = "article"
    Public Const zzEtiquetaCodigoArticulo As String = "articleCode"

    Public Const zzXML_NombreArchivoConfiguraciones As String = "configurations"
    Public Const zzXML_EtiquetaRaizConfiguraciones As String = "configurations"
    Public Const zzXML_EtiquetaNodoConfiguracion As String = "configuration"
    Public Const zzEtiquetaCodigoConfiguracion As String = "configurationCode"

    Public Const zzXML_NombreArchivoArticulosGenericos As String = "generics"
    Public Const zzXML_EtiquetaRaizArticulosGenericos As String = "generics"
    Public Const zzXML_EtiquetaNodoArticuloGenerico As String = "generic"
    Public Const zzEtiquetaCodigoArticuloGenerico As String = "genericCode"

    Public Const zzXML_NombreArchivoBloquesAuxiliares As String = "autocad_auxiliaryblocks"
    Public Const zzXML_EtiquetaRaizBloquesAuxiliares As String = "autocadAuxiliaryBlocks"
    Public Const zzXML_EtiquetaNodoBloqueAuxiliar As String = "block"
    Public Const zzEtiquetaCodigoBloqueAuxiliar As String = "blockCode"

    Public Const zzXML_NombreArchivoGrupos As String = "groups"
    Public Const zzXML_EtiquetaRaizGrupos As String = "groups"
    Public Const zzXML_EtiquetaNodoGrupo As String = "group"
    Public Const zzEtiquetaCodigoGrupo As String = "groupCode"

    Public Const zzXML_NombreArchivoSubgrupos As String = "subgroups"
    Public Const zzXML_EtiquetaRaizSubgrupos As String = "subgroups"
    Public Const zzXML_EtiquetaNodoSubgrupo As String = "subgroup"
    Public Const zzEtiquetaCodigoSubgrupo As String = "subgroupCode"

    Public Const zzEtiquetaDescripcion As String = "description"
    Public Const zzEtiquetaPeso As String = "weight"


    Public Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, ByVal incluirDescripciones As Boolean, descripcionesEstiloRepcon As Boolean) As XmlNode

        Dim nElemento As XmlNode
        Dim nCodigo As XmlNode
        Select Case TipoElemento
            Case TiposDeElemento.articulo
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoArticulo)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoArticulo)
            Case TiposDeElemento.configuracion
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoConfiguracion)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoConfiguracion)
            Case Elemento.TiposDeElemento.articuloGenerico
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoArticuloGenerico)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoArticuloGenerico)
            Case TiposDeElemento.bloqueAuxiliarAutocad
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoBloqueAuxiliar)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoBloqueAuxiliar)
            Case TiposDeElemento.familiaRevit
                nElemento = documentoDondeSeVaAInsertar.CreateElement(FamiliaRevit.zzXML_EtiquetaNodo)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(FamiliaRevit.zzXML_EtiquetaCodigoFamiliaRevit)
            Case TiposDeElemento.grupo
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoGrupo)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoGrupo)
            Case TiposDeElemento.subgrupo
                nElemento = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodoSubgrupo)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoSubgrupo)
            Case Else
                nElemento = documentoDondeSeVaAInsertar.CreateElement("ElementType_" & TipoElemento)
                nCodigo = documentoDondeSeVaAInsertar.CreateElement("code")
        End Select
        nCodigo.InnerText = CodElemento
        nElemento.AppendChild(nCodigo)

        If incluirDescripciones Then

            Dim nDescripcionLocal As XmlNode
            nDescripcionLocal = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
            nDescripcionLocal.InnerText = NombreLocal
            Dim aIdiomaLocal As XmlAttribute
            aIdiomaLocal = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdiomaLocal.InnerText = "local"
            nDescripcionLocal.Attributes.Append(aIdiomaLocal)
            nElemento.AppendChild(nDescripcionLocal)

            If descripcionesEstiloRepcon Then

                Dim nDescripcion_es As XmlNode
                nDescripcion_es = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_es.InnerText = NAME_es
                Dim aIdioma_es As XmlAttribute
                aIdioma_es = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_es.InnerText = "EScorp"
                nDescripcion_es.Attributes.Append(aIdioma_es)
                nElemento.AppendChild(nDescripcion_es)

                Dim nDescripcion_en As XmlNode
                nDescripcion_en = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_en.InnerText = NAME_en
                Dim aIdioma_en As XmlAttribute
                aIdioma_en = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_en.InnerText = "ENcorp"
                nDescripcion_en.Attributes.Append(aIdioma_en)
                nElemento.AppendChild(nDescripcion_en)

                If TipoElemento = TiposDeElemento.articulo _
                Or TipoElemento = TiposDeElemento.articuloGenerico _
                Or TipoElemento = TiposDeElemento.bloqueAuxiliarAutocad _
                Or TipoElemento = TiposDeElemento.configuracion _
                Then
                    Dim nDescripcion_fr As XmlNode
                    nDescripcion_fr = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                    nDescripcion_fr.InnerText = NAME_fr
                    Dim aIdioma_fr As XmlAttribute
                    aIdioma_fr = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                    aIdioma_fr.InnerText = "FRcorp"
                    nDescripcion_fr.Attributes.Append(aIdioma_fr)
                    nElemento.AppendChild(nDescripcion_fr)
                End If

            Else

                Dim nDescripcion_es As XmlNode
                nDescripcion_es = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_es.InnerText = NAME_es
                Dim aIdioma_es As XmlAttribute
                aIdioma_es = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_es.InnerText = "es"
                nDescripcion_es.Attributes.Append(aIdioma_es)
                nElemento.AppendChild(nDescripcion_es)

                Dim nDescripcion_en As XmlNode
                nDescripcion_en = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_en.InnerText = NAME_en
                Dim aIdioma_en As XmlAttribute
                aIdioma_en = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_en.InnerText = "en"
                nDescripcion_en.Attributes.Append(aIdioma_en)
                nElemento.AppendChild(nDescripcion_en)

                Dim nDescripcion_fr As XmlNode
                nDescripcion_fr = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_fr.InnerText = NAME_fr
                Dim aIdioma_fr As XmlAttribute
                aIdioma_fr = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_fr.InnerText = "fr"
                nDescripcion_fr.Attributes.Append(aIdioma_fr)
                nElemento.AppendChild(nDescripcion_fr)

                Dim nDescripcion_esES As XmlNode
                nDescripcion_esES = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_esES.InnerText = NAME_esES
                Dim aIdioma_esES As XmlAttribute
                aIdioma_esES = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_esES.InnerText = "es-ES"
                nDescripcion_esES.Attributes.Append(aIdioma_esES)
                nElemento.AppendChild(nDescripcion_esES)

                Dim nDescripcion_esPE As XmlNode
                nDescripcion_esPE = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_esPE.InnerText = NAME_esPE
                Dim aIdioma_esPE As XmlAttribute
                aIdioma_esPE = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_esPE.InnerText = "es-PE"
                nDescripcion_esPE.Attributes.Append(aIdioma_esPE)
                nElemento.AppendChild(nDescripcion_esPE)

                Dim nDescripcion_esCL As XmlNode
                nDescripcion_esCL = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_esCL.InnerText = NAME_esCL
                Dim aIdioma_esCL As XmlAttribute
                aIdioma_esCL = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_esCL.InnerText = "es-CL"
                nDescripcion_esCL.Attributes.Append(aIdioma_esCL)
                nElemento.AppendChild(nDescripcion_esCL)

                Dim nDescripcion_esMX As XmlNode
                nDescripcion_esMX = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_esMX.InnerText = NAME_esMX
                Dim aIdioma_esMX As XmlAttribute
                aIdioma_esMX = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_esMX.InnerText = "es-MX"
                nDescripcion_esMX.Attributes.Append(aIdioma_esMX)
                nElemento.AppendChild(nDescripcion_esMX)

                Dim nDescripcion_enUS As XmlNode
                nDescripcion_enUS = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_enUS.InnerText = NAME_enUS
                Dim aIdioma_enUS As XmlAttribute
                aIdioma_enUS = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_enUS.InnerText = "en-US"
                nDescripcion_enUS.Attributes.Append(aIdioma_enUS)
                nElemento.AppendChild(nDescripcion_enUS)

                Dim nDescripcion_deDE As XmlNode
                nDescripcion_deDE = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_deDE.InnerText = NAME_deDE
                Dim aIdioma_deDE As XmlAttribute
                aIdioma_deDE = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_deDE.InnerText = "de-DE"
                nDescripcion_deDE.Attributes.Append(aIdioma_deDE)
                nElemento.AppendChild(nDescripcion_deDE)

                Dim nDescripcion_frFR As XmlNode
                nDescripcion_frFR = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_frFR.InnerText = NAME_frFR
                Dim aIdioma_frFR As XmlAttribute
                aIdioma_frFR = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_frFR.InnerText = "fr-FR"
                nDescripcion_frFR.Attributes.Append(aIdioma_frFR)
                nElemento.AppendChild(nDescripcion_frFR)

                Dim nDescripcion_itIT As XmlNode
                nDescripcion_itIT = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_itIT.InnerText = NAME_itIT
                Dim aIdioma_itIT As XmlAttribute
                aIdioma_itIT = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_itIT.InnerText = "it-IT"
                nDescripcion_itIT.Attributes.Append(aIdioma_itIT)
                nElemento.AppendChild(nDescripcion_itIT)

                Dim nDescripcion_plPL As XmlNode
                nDescripcion_plPL = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_plPL.InnerText = NAME_plPL
                Dim aIdioma_plPL As XmlAttribute
                aIdioma_plPL = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_plPL.InnerText = "pl-PL"
                nDescripcion_plPL.Attributes.Append(aIdioma_plPL)
                nElemento.AppendChild(nDescripcion_plPL)

                Dim nDescripcion_ptPT As XmlNode
                nDescripcion_ptPT = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_ptPT.InnerText = NAME_ptPT
                Dim aIdioma_ptPT As XmlAttribute
                aIdioma_ptPT = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_ptPT.InnerText = "pt-PT"
                nDescripcion_ptPT.Attributes.Append(aIdioma_ptPT)
                nElemento.AppendChild(nDescripcion_ptPT)

                Dim nDescripcion_czCZ As XmlNode
                nDescripcion_czCZ = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_czCZ.InnerText = NAME_czCZ
                Dim aIdioma_czCZ As XmlAttribute
                aIdioma_czCZ = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_czCZ.InnerText = "cz-CZ"
                nDescripcion_czCZ.Attributes.Append(aIdioma_czCZ)
                nElemento.AppendChild(nDescripcion_czCZ)

                Dim nDescripcion_skSK As XmlNode
                nDescripcion_skSK = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_skSK.InnerText = NAME_skSK
                Dim aIdioma_skSK As XmlAttribute
                aIdioma_skSK = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_skSK.InnerText = "sk-SK"
                nDescripcion_skSK.Attributes.Append(aIdioma_skSK)
                nElemento.AppendChild(nDescripcion_skSK)

                Dim nDescripcion_roRO As XmlNode
                nDescripcion_roRO = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaDescripcion)
                nDescripcion_roRO.InnerText = NAME_roRO
                Dim aIdioma_roRO As XmlAttribute
                aIdioma_roRO = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
                aIdioma_roRO.InnerText = "ro-RO"
                nDescripcion_roRO.Attributes.Append(aIdioma_roRO)
                nElemento.AppendChild(nDescripcion_roRO)

            End If

        End If


        Return nElemento

    End Function


    Public Sub fromXML(ByVal LeerElementoDesde As XmlNode)
        If LeerElementoDesde.Name = zzXML_EtiquetaNodoArticulo _
        Or LeerElementoDesde.Name = zzXML_EtiquetaNodoArticuloGenerico _
        Or LeerElementoDesde.Name = zzXML_EtiquetaNodoConfiguracion _
        Or LeerElementoDesde.Name = zzXML_EtiquetaNodoBloqueAuxiliar _
        Or LeerElementoDesde.Name = FamiliaRevit.zzXML_EtiquetaNodo _
        Or LeerElementoDesde.Name = zzXML_EtiquetaNodoGrupo _
        Or LeerElementoDesde.Name = zzXML_EtiquetaNodoSubgrupo _
        Then

            Me.Reset()

            Dim nCodElemento As XmlNode = Nothing
            Select Case LeerElementoDesde.Name
                Case zzXML_EtiquetaNodoArticulo
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoArticulo)
                    TipoElemento = TiposDeElemento.articulo
                Case zzXML_EtiquetaNodoArticuloGenerico
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoArticuloGenerico)
                    TipoElemento = TiposDeElemento.articuloGenerico
                Case zzXML_EtiquetaNodoConfiguracion
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoConfiguracion)
                    TipoElemento = TiposDeElemento.configuracion
                Case zzXML_EtiquetaNodoBloqueAuxiliar
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoBloqueAuxiliar)
                    TipoElemento = TiposDeElemento.bloqueAuxiliarAutocad
                Case FamiliaRevit.zzXML_EtiquetaNodo
                    nCodElemento = LeerElementoDesde.SelectSingleNode(FamiliaRevit.zzXML_EtiquetaCodigoFamiliaRevit)
                    TipoElemento = TiposDeElemento.familiaRevit
                Case zzXML_EtiquetaNodoGrupo
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoGrupo)
                    TipoElemento = TiposDeElemento.grupo
                Case zzXML_EtiquetaNodoSubgrupo
                    nCodElemento = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoSubgrupo)
                    TipoElemento = TiposDeElemento.subgrupo
                Case Else
                    TipoElemento = TiposDeElemento.DESCONOCIDO
            End Select

            If Not IsNothing(nCodElemento) Then
                CodElemento = nCodElemento.InnerText
            End If


            Dim nNombreLocal As XmlNode
            nNombreLocal = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & _
                                                    "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                    "='local']")
            If Not IsNothing(nNombreLocal) Then
                NombreLocal = nNombreLocal.InnerText
            End If

            Dim nNombre_es As XmlNode
            nNombre_es = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='es']")
            If Not IsNothing(nNombre_es) Then
                NAME_es = nNombre_es.InnerText
            End If

            Dim nNombre_esES As XmlNode
            nNombre_esES = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='es-ES']")
            If Not IsNothing(nNombre_esES) Then
                NAME_esES = nNombre_esES.InnerText
            End If

            Dim nNombre_esPE As XmlNode
            nNombre_esPE = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='es-PE']")
            If Not IsNothing(nNombre_esPE) Then
                NAME_esPE = nNombre_esPE.InnerText
            End If

            Dim nNombre_esCL As XmlNode
            nNombre_esCL = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='es-CL']")
            If Not IsNothing(nNombre_esCL) Then
                NAME_esCL = nNombre_esCL.InnerText
            End If

            Dim nNombre_esMX As XmlNode
            nNombre_esMX = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='es-MX']")
            If Not IsNothing(nNombre_esMX) Then
                NAME_esMX = nNombre_esMX.InnerText
            End If

            Dim nNombre_en As XmlNode
            nNombre_en = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='en']")
            If Not IsNothing(nNombre_en) Then
                NAME_en = nNombre_en.InnerText
            End If

            Dim nNombre_enUS As XmlNode
            nNombre_enUS = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='en-US']")
            If Not IsNothing(nNombre_enUS) Then
                NAME_enUS = nNombre_enUS.InnerText
            End If

            Dim nNombre_deDE As XmlNode
            nNombre_deDE = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='de-DE']")
            If Not IsNothing(nNombre_deDE) Then
                NAME_deDE = nNombre_deDE.InnerText
            End If

            Dim nNombre_fr As XmlNode
            nNombre_fr = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='fr']")
            If Not IsNothing(nNombre_fr) Then
                NAME_fr = nNombre_fr.InnerText
            End If

            Dim nNombre_frFR As XmlNode
            nNombre_frFR = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='fr-FR']")
            If Not IsNothing(nNombre_frFR) Then
                NAME_frFR = nNombre_frFR.InnerText
            End If

            Dim nNombre_itIT As XmlNode
            nNombre_itIT = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='it-IT']")
            If Not IsNothing(nNombre_itIT) Then
                NAME_itIT = nNombre_itIT.InnerText
            End If

            Dim nNombre_plPL As XmlNode
            nNombre_plPL = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='pl-PL']")
            If Not IsNothing(nNombre_plPL) Then
                NAME_plPL = nNombre_plPL.InnerText
            End If


            Dim nNombre_ptPT As XmlNode
            nNombre_ptPT = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='pt-PT']")
            If Not IsNothing(nNombre_ptPT) Then
                NAME_ptPT = nNombre_ptPT.InnerText
            End If

            Dim nNombre_czCZ As XmlNode
            nNombre_czCZ = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='cz-CZ']")
            If Not IsNothing(nNombre_czCZ) Then
                NAME_czCZ = nNombre_czCZ.InnerText
            End If

            Dim nNombre_skSK As XmlNode
            nNombre_skSK = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='sk-SK']")
            If Not IsNothing(nNombre_skSK) Then
                NAME_skSK = nNombre_skSK.InnerText
            End If

            Dim nNombre_roRO As XmlNode
            nNombre_roRO = LeerElementoDesde.SelectSingleNode(ConstantesComunesParaXML.zzEtiquetaDescripcion & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                          "='ro-RO']")
            If Not IsNothing(nNombre_roRO) Then
                NAME_roRO = nNombre_roRO.InnerText
            End If


        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerElementoDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodoArticulo _
                                                                                                                                 & " o " & zzXML_EtiquetaNodoArticuloGenerico _
                                                                                                                                 & " o " & zzXML_EtiquetaNodoConfiguracion _
                                                                                                                                 & " o " & zzXML_EtiquetaNodoBloqueAuxiliar _
                                                                                                                                 & " o " & zzXML_EtiquetaNodoGrupo _
                                                                                                                                 & " o " & zzXML_EtiquetaNodoSubgrupo & ".")
        End If
    End Sub






    Public Overrides Function Equals(ByVal otroObjeto As Object) As Boolean Implements IEquatable(Of Object).Equals
        If otroObjeto.GetType.Equals(Me.GetType) Then
            Dim otroElemento = CType(otroObjeto, Elemento)
            If otroElemento.CodElemento.Equals(Me.CodElemento) _
            And otroElemento.NombreLocal.Equals(Me.NombreLocal) _
            And otroElemento.NAME_es.Equals(Me.NAME_es) _
            And otroElemento.NAME_esES.Equals(Me.NAME_esES) _
            And otroElemento.NAME_esPE.Equals(Me.NAME_esPE) _
            And otroElemento.NAME_esCL.Equals(Me.NAME_esCL) _
            And otroElemento.NAME_esMX.Equals(Me.NAME_esMX) _
            And otroElemento.NAME_en.Equals(Me.NAME_en) _
            And otroElemento.NAME_enUS.Equals(Me.NAME_enUS) _
            And otroElemento.NAME_fr.Equals(Me.NAME_fr) _
            And otroElemento.NAME_frFR.Equals(Me.NAME_frFR) _
            And otroElemento.NAME_deDE.Equals(Me.NAME_deDE) _
            And otroElemento.NAME_itIT.Equals(Me.NAME_itIT) _
            And otroElemento.NAME_ptPT.Equals(Me.NAME_ptPT) _
            And otroElemento.NAME_plPL.Equals(Me.NAME_plPL) _
            And otroElemento.NAME_czCZ.Equals(Me.NAME_czCZ) _
            And otroElemento.NAME_skSK.Equals(Me.NAME_skSK) _
            And otroElemento.NAME_roRO.Equals(Me.NAME_roRO) _
            Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Public Function CompareTo(ByVal otroElemento As Elemento) As Integer Implements IComparable(Of Elemento).CompareTo
        If otroElemento.CodElemento.Equals(Me.CodElemento) Then
            Return 0
        ElseIf otroElemento.CodElemento < Me.CodElemento Then
            Return 1
        Else
            Return -1
        End If
    End Function



End Class
