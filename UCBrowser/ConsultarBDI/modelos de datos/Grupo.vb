Imports System.Xml


Public Class Grupo : Inherits Elemento

    Public CodLineaProducto As String
    Public Abreviatura As String
    Public estaActivo As Boolean = True
    Public SHORT_es As String
    Public SHORT_en As String
    Public SHORT_esES As String
    Public SHORT_esPE As String
    Public SHORT_esCL As String
    Public SHORT_esMX As String
    Public SHORT_frFR As String
    Public SHORT_ptPT As String
    Public SHORT_deDE As String
    Public SHORT_plPL As String
    Public SHORT_enUS As String
    Public SHORT_itIT As String
    Public SHORT_czCZ As String
    Public SHORT_skSK As String
    Public SHORT_roRO As String


    Public Sub New()
        TipoElemento = TiposDeElemento.grupo
    End Sub

    Public Function getAbreviatura(idioma As String) As String
        Dim abreviatura As String = ""
        Select Case idioma
            Case "es"
                abreviatura = SHORT_es
            Case "en"
                abreviatura = SHORT_en
            Case "es-ES"
                abreviatura = SHORT_esES
            Case "es-PE"
                abreviatura = SHORT_esPE
            Case "es-CL"
                abreviatura = SHORT_esCL
            Case "es-MX"
                abreviatura = SHORT_esMX
            Case "fr-FR"
                abreviatura = SHORT_frFR
            Case "pt-PT"
                abreviatura = SHORT_ptPT
            Case "de-DE"
                abreviatura = SHORT_deDE
            Case "pl-PL"
                abreviatura = SHORT_plPL
            Case "en-US"
                abreviatura = SHORT_enUS
            Case "it-IT"
                abreviatura = SHORT_itIT
            Case "cz-CZ"
                abreviatura = SHORT_czCZ
            Case "sk-SK"
                abreviatura = SHORT_skSK
            Case "ro-RO"
                abreviatura = SHORT_roRO
        End Select

        If abreviatura = "" Then abreviatura = SHORT_en
        Return abreviatura
    End Function

    Public Const zzEtiquetaCodigoLineaProducto As String = "productType"
    Public Const zzEtiquetaActivo As String = "active"
    Public Const zzEtiquetaAbreviatura As String = "shortName"

    Public Overloads Function ToXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, ByVal incluirDescripciones As Boolean, descripcionesEstiloRepcon As Boolean) As XmlNode

        Dim nElemento As XmlNode
        nElemento = MyBase.toXML(documentoDondeSeVaAInsertar, incluirDescripciones, descripcionesEstiloRepcon)

        Dim nCodLineaDeProducto As XmlNode = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaCodigoLineaProducto)
        nCodLineaDeProducto.InnerText = CodLineaProducto
        nElemento.AppendChild(nCodLineaDeProducto)

        Dim nAbreviatura As XmlNode = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
        nAbreviatura.InnerText = Abreviatura
        nElemento.AppendChild(nAbreviatura)

        If Not descripcionesEstiloRepcon Then

            Dim nActivo As XmlNode
            nActivo = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaActivo)
            nActivo.InnerText = estaActivo.ToString()
            nElemento.AppendChild(nActivo)

            Dim nAbreviatura_es As XmlNode
            nAbreviatura_es = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_es.InnerText = SHORT_es
            Dim aIdioma_es As XmlAttribute
            aIdioma_es = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_es.InnerText = "es"
            nAbreviatura_es.Attributes.Append(aIdioma_es)
            nElemento.AppendChild(nAbreviatura_es)

            Dim nAbreviatura_en As XmlNode
            nAbreviatura_en = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_en.InnerText = SHORT_en
            Dim aIdioma_en As XmlAttribute
            aIdioma_en = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_en.InnerText = "en"
            nAbreviatura_en.Attributes.Append(aIdioma_en)
            nElemento.AppendChild(nAbreviatura_en)

            Dim nAbreviatura_esES As XmlNode
            nAbreviatura_esES = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_esES.InnerText = SHORT_esES
            Dim aIdioma_esES As XmlAttribute
            aIdioma_esES = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_esES.InnerText = "es-ES"
            nAbreviatura_esES.Attributes.Append(aIdioma_esES)
            nElemento.AppendChild(nAbreviatura_esES)

            Dim nAbreviatura_esPE As XmlNode
            nAbreviatura_esPE = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_esPE.InnerText = SHORT_esPE
            Dim aIdioma_esPE As XmlAttribute
            aIdioma_esPE = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_esPE.InnerText = "es-PE"
            nAbreviatura_esPE.Attributes.Append(aIdioma_esPE)
            nElemento.AppendChild(nAbreviatura_esPE)

            Dim nAbreviatura_esCL As XmlNode
            nAbreviatura_esCL = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_esCL.InnerText = SHORT_esCL
            Dim aIdioma_esCL As XmlAttribute
            aIdioma_esCL = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_esCL.InnerText = "es-CL"
            nAbreviatura_esCL.Attributes.Append(aIdioma_esCL)
            nElemento.AppendChild(nAbreviatura_esCL)

            Dim nAbreviatura_esMX As XmlNode
            nAbreviatura_esMX = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_esMX.InnerText = SHORT_esMX
            Dim aIdioma_esMX As XmlAttribute
            aIdioma_esMX = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_esMX.InnerText = "es-MX"
            nAbreviatura_esMX.Attributes.Append(aIdioma_esMX)
            nElemento.AppendChild(nAbreviatura_esMX)

            Dim nAbreviatura_enUS As XmlNode
            nAbreviatura_enUS = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_enUS.InnerText = SHORT_enUS
            Dim aIdioma_enUS As XmlAttribute
            aIdioma_enUS = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_enUS.InnerText = "en-US"
            nAbreviatura_enUS.Attributes.Append(aIdioma_enUS)
            nElemento.AppendChild(nAbreviatura_enUS)

            Dim nAbreviatura_deDE As XmlNode
            nAbreviatura_deDE = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_deDE.InnerText = SHORT_deDE
            Dim aIdioma_deDE As XmlAttribute
            aIdioma_deDE = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_deDE.InnerText = "de-DE"
            nAbreviatura_deDE.Attributes.Append(aIdioma_deDE)
            nElemento.AppendChild(nAbreviatura_deDE)

            Dim nAbreviatura_frFR As XmlNode
            nAbreviatura_frFR = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_frFR.InnerText = SHORT_frFR
            Dim aIdioma_frFR As XmlAttribute
            aIdioma_frFR = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_frFR.InnerText = "fr-FR"
            nAbreviatura_frFR.Attributes.Append(aIdioma_frFR)
            nElemento.AppendChild(nAbreviatura_frFR)

            Dim nAbreviatura_itIT As XmlNode
            nAbreviatura_itIT = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_itIT.InnerText = SHORT_itIT
            Dim aIdioma_itIT As XmlAttribute
            aIdioma_itIT = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_itIT.InnerText = "it-IT"
            nAbreviatura_itIT.Attributes.Append(aIdioma_itIT)
            nElemento.AppendChild(nAbreviatura_itIT)

            Dim nAbreviatura_plPL As XmlNode
            nAbreviatura_plPL = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_plPL.InnerText = SHORT_plPL
            Dim aIdioma_plPL As XmlAttribute
            aIdioma_plPL = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_plPL.InnerText = "pl-PL"
            nAbreviatura_plPL.Attributes.Append(aIdioma_plPL)
            nElemento.AppendChild(nAbreviatura_plPL)

            Dim nAbreviatura_ptPT As XmlNode
            nAbreviatura_ptPT = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_ptPT.InnerText = SHORT_ptPT
            Dim aIdioma_ptPT As XmlAttribute
            aIdioma_ptPT = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_ptPT.InnerText = "pt-PT"
            nAbreviatura_ptPT.Attributes.Append(aIdioma_ptPT)
            nElemento.AppendChild(nAbreviatura_ptPT)

            Dim nAbreviatura_czCZ As XmlNode
            nAbreviatura_czCZ = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_czCZ.InnerText = SHORT_czCZ
            Dim aIdioma_czCZ As XmlAttribute
            aIdioma_czCZ = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_czCZ.InnerText = "cz-CZ"
            nAbreviatura_czCZ.Attributes.Append(aIdioma_czCZ)
            nElemento.AppendChild(nAbreviatura_czCZ)

            Dim nAbreviatura_skSK As XmlNode
            nAbreviatura_skSK = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_skSK.InnerText = SHORT_skSK
            Dim aIdioma_skSK As XmlAttribute
            aIdioma_skSK = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_skSK.InnerText = "sk-SK"
            nAbreviatura_skSK.Attributes.Append(aIdioma_skSK)
            nElemento.AppendChild(nAbreviatura_skSK)

            Dim nAbreviatura_roRO As XmlNode
            nAbreviatura_roRO = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaAbreviatura)
            nAbreviatura_roRO.InnerText = SHORT_roRO
            Dim aIdioma_roRO As XmlAttribute
            aIdioma_roRO = documentoDondeSeVaAInsertar.CreateAttribute(ConstantesComunesParaXML.zzEtiquetaIdioma)
            aIdioma_roRO.InnerText = "ro-RO"
            nAbreviatura_roRO.Attributes.Append(aIdioma_roRO)
            nElemento.AppendChild(nAbreviatura_roRO)

        End If

        Return nElemento

    End Function

    Public Overloads Sub fromXML(ByVal LeerElementoDesde As XmlNode)

        MyBase.fromXML(LeerElementoDesde)

        Dim nCodLineaProducto As XmlNode
        nCodLineaProducto = LeerElementoDesde.SelectSingleNode(zzEtiquetaCodigoLineaProducto)
        If Not IsNothing(nCodLineaProducto) Then
            CodLineaProducto = nCodLineaProducto.InnerText
        End If

        Dim nAbreviatura As XmlNode
        nAbreviatura = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura)
        If Not IsNothing(nAbreviatura) Then
            Abreviatura = nAbreviatura.InnerText
        End If

        Dim nActivo As XmlNode
        nActivo = LeerElementoDesde.SelectSingleNode(zzEtiquetaActivo)
        If Not IsNothing(nActivo) Then
            Boolean.TryParse(nActivo.InnerText, estaActivo)
        End If

        Dim nAbreviatura_es As XmlNode
        nAbreviatura_es = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='es']")
        If Not IsNothing(nAbreviatura_es) Then
            SHORT_es = nAbreviatura_es.InnerText
        End If

        Dim nAbreviatura_esES As XmlNode
        nAbreviatura_esES = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='es-ES']")
        If Not IsNothing(nAbreviatura_esES) Then
            SHORT_esES = nAbreviatura_esES.InnerText
        End If

        Dim nAbreviatura_esPE As XmlNode
        nAbreviatura_esPE = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='es-PE']")
        If Not IsNothing(nAbreviatura_esPE) Then
            SHORT_esPE = nAbreviatura_esPE.InnerText
        End If

        Dim nAbreviatura_esCL As XmlNode
        nAbreviatura_esCL = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='es-CL']")
        If Not IsNothing(nAbreviatura_esCL) Then
            SHORT_esCL = nAbreviatura_esCL.InnerText
        End If

        Dim nAbreviatura_esMX As XmlNode
        nAbreviatura_esMX = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='es-MX']")
        If Not IsNothing(nAbreviatura_esMX) Then
            SHORT_esMX = nAbreviatura_esMX.InnerText
        End If

        Dim nAbreviatura_en As XmlNode
        nAbreviatura_en = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='en']")
        If Not IsNothing(nAbreviatura_en) Then
            SHORT_en = nAbreviatura_en.InnerText
        End If

        Dim nAbreviatura_enUS As XmlNode
        nAbreviatura_enUS = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='en-US']")
        If Not IsNothing(nAbreviatura_enUS) Then
            SHORT_enUS = nAbreviatura_enUS.InnerText
        End If

        Dim nAbreviatura_deDE As XmlNode
        nAbreviatura_deDE = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='de-DE']")
        If Not IsNothing(nAbreviatura_deDE) Then
            SHORT_deDE = nAbreviatura_deDE.InnerText
        End If

        Dim nAbreviatura_fr As XmlNode
        nAbreviatura_fr = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='fr']")
        Dim nAbreviatura_frFR As XmlNode
        nAbreviatura_frFR = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='fr-FR']")
        If Not IsNothing(nAbreviatura_frFR) Then
            SHORT_frFR = nAbreviatura_frFR.InnerText
        End If

        Dim nAbreviatura_itIT As XmlNode
        nAbreviatura_itIT = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='it-IT']")
        If Not IsNothing(nAbreviatura_itIT) Then
            SHORT_itIT = nAbreviatura_itIT.InnerText
        End If

        Dim nAbreviatura_plPL As XmlNode
        nAbreviatura_plPL = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='pl-PL']")
        If Not IsNothing(nAbreviatura_plPL) Then
            SHORT_plPL = nAbreviatura_plPL.InnerText
        End If


        Dim nAbreviatura_ptPT As XmlNode
        nAbreviatura_ptPT = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='pt-PT']")
        If Not IsNothing(nAbreviatura_ptPT) Then
            SHORT_ptPT = nAbreviatura_ptPT.InnerText
        End If

        Dim nAbreviatura_czCZ As XmlNode
        nAbreviatura_czCZ = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='cz-CZ']")
        If Not IsNothing(nAbreviatura_czCZ) Then
            SHORT_czCZ = nAbreviatura_czCZ.InnerText
        End If

        Dim nAbreviatura_skSK As XmlNode
        nAbreviatura_skSK = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='sk-SK']")
        If Not IsNothing(nAbreviatura_skSK) Then
            SHORT_skSK = nAbreviatura_skSK.InnerText
        End If

        Dim nAbreviatura_roRO As XmlNode
        nAbreviatura_roRO = LeerElementoDesde.SelectSingleNode(zzEtiquetaAbreviatura & "[@" & ConstantesComunesParaXML.zzEtiquetaIdioma & _
                                                      "='ro-RO']")
        If Not IsNothing(nAbreviatura_roRO) Then
            SHORT_roRO = nAbreviatura_roRO.InnerText
        End If

    End Sub


End Class
