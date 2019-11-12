Imports System.Xml
Imports Magnitudes


Public Class ElementoDatosadicionales : Implements IEquatable(Of Object) : Implements IComparable(Of ElementoDatosadicionales)

    Public TipoElemento As Elemento.TiposDeElemento
    Public CodElemento As String
    Public Peso As Magnitud
    Public AreaEncofrado As Magnitud
    Public Material As String
    Public UnidadStock As String
    Public Generico As String
    Public Longitud As Magnitud
    Public Anchura As Magnitud
    Public Area As Magnitud
    Public Altura As Magnitud
    Public AlturaMin As Magnitud
    Public AlturaMax As Magnitud
    Public Espesor As Magnitud
    Public Cortante As Magnitud
    Public Inercia As Magnitud
    Public Momento As Magnitud
    Public Densidad As Magnitud
    Public ResisFlexionPerpen As Magnitud
    Public ResisFlexionParal As Magnitud
    Public ModuloElasticoMedioPerpen As Magnitud
    Public ModuloElasticoMedioParal As Magnitud
    Public DiametroTuboInterior As Magnitud
    Public ModuloYoung As Magnitud
    Public AreaCortante As Magnitud
    Public ModuloResistente As Magnitud
    Public ResisCortantePerpend As Magnitud
    Public ResisCortanteParal As Magnitud
    Public ModuloRigidezMedioPerpend As Magnitud
    Public ModuloRigidezMedioParalelo As Magnitud


    Public Sub New()
        Reset()
    End Sub
    Public Sub Reset()
        TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
        CodElemento = ""
        Peso = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        AreaEncofrado = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Material = ""
        Generico = ""
        UnidadStock = ""
        Longitud = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Anchura = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Area = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Altura = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        AlturaMin = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        AlturaMax = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Espesor = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Cortante = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Inercia = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Momento = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Densidad = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ResisFlexionPerpen = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ResisFlexionParal = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloElasticoMedioPerpen = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloElasticoMedioParal = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        DiametroTuboInterior = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloYoung = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        AreaCortante = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloResistente = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ResisCortantePerpend = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ResisCortanteParal = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloRigidezMedioPerpend = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        ModuloRigidezMedioParalelo = New Magnitud(valor:=Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
    End Sub



    Public Const zzXMLEtiquetaNodoDatosadicionales As String = "aditionalData"

    Public Const zzEtiquetaPeso As String = "weight"
    Public Const zzEtiquetaAreaEncofrado As String = "formArea"
    Public Const zzEtiquetaMaterial As String = "material"
    Public Const zzEtiquetaUnidadStock As String = "stockUnit"
    Public Const zzEtiquetaLongitud As String = "length"
    Public Const zzEtiquetaAnchura As String = "width"
    Public Const zzEtiquetaArea As String = "area"
    Public Const zzEtiquetaAltura As String = "height"
    Public Const zzEtiquetaAlturaMin As String = "minHeight"
    Public Const zzEtiquetaAlturaMax As String = "maxHeight"
    Public Const zzEtiquetaEspesor As String = "thickness"
    Public Const zzEtiquetaCortante As String = "shear"
    Public Const zzEtiquetaInercia As String = "inertia"
    Public Const zzEtiquetaMomento As String = "moment"
    Public Const zzEtiquetaDensidad As String = "density"
    Public Const zzEtiquetaResisFlexionPerpen As String = "perpendicularStrenghBending"
    Public Const zzEtiquetaResisFlexionParal As String = "parallelStrenghBending"
    Public Const zzEtiquetaModuloElasticoMedioPerpen As String = "perpendicularElasticModulus"
    Public Const zzEtiquetaModuloElasticoMedioParal As String = "parallelElasticModulus"
    Public Const zzEtiquetaDiametroTuboInterior As String = "internalTubeDiameterUnit"
    Public Const zzEtiquetaModuloYoung As String = "youngModulus"
    Public Const zzEtiquetaGenerico As String = "generic"
    Public Const zzEtiquetaAreaCortante As String = "shearArea"
    Public Const zzEtiquetaModuloResistente As String = "sectionModulus"
    Public Const zzEtiquetaResisCortantePerpend As String = "perpendicularShear"
    Public Const zzEtiquetaResisCortanteParal As String = "parallelShear"
    Public Const zzEtiquetaModuloRigidezMedioPerpend As String = "perpendicularModulusRigidity"
    Public Const zzEtiquetaModuloRigidezMedioParalelo As String = "parallelModulusRigidity"


    Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, ByVal omitirDatosVaciosONoValidos As Boolean) As XmlNode
        Dim nDatosadicionales As XmlNode
        nDatosadicionales = documentoDondeSeVaAInsertar.CreateElement(zzXMLEtiquetaNodoDatosadicionales)

        If Not (omitirDatosVaciosONoValidos And Material.Equals("")) Then
            Dim nMaterial As XmlNode
            nMaterial = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaMaterial)
            nMaterial.InnerText = Material
            nDatosadicionales.AppendChild(nMaterial)
        End If

        If Not (omitirDatosVaciosONoValidos And Generico.Equals("")) Then
            Dim nGenerico As XmlNode
            nGenerico = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaGenerico)
            nGenerico.InnerText = Generico
            nDatosadicionales.AppendChild(nGenerico)
        End If

        If Not (omitirDatosVaciosONoValidos And Peso.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Peso.toXML(documentoDondeSeVaAInsertar, zzEtiquetaPeso))
        End If

        If Not (omitirDatosVaciosONoValidos And UnidadStock.Equals("")) Then
            Dim nUnidadStock As XmlNode
            nUnidadStock = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaUnidadStock)
            nUnidadStock.InnerText = UnidadStock
            nDatosadicionales.AppendChild(nUnidadStock)
        End If

        If Not (omitirDatosVaciosONoValidos And AreaEncofrado.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(AreaEncofrado.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaAreaEncofrado))
        End If

        If Not (omitirDatosVaciosONoValidos And Longitud.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Longitud.toXML(documentoDondeSeVaAInsertar, zzEtiquetaLongitud))
        End If

        If Not (omitirDatosVaciosONoValidos And Anchura.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Anchura.toXML(documentoDondeSeVaAInsertar, zzEtiquetaAnchura))
        End If

        If Not (omitirDatosVaciosONoValidos And Area.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Area.toXML(documentoDondeSeVaAInsertar, zzEtiquetaArea))
        End If

        If Not (omitirDatosVaciosONoValidos And Altura.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Altura.toXML(documentoDondeSeVaAInsertar, zzEtiquetaAltura))
        End If

        If Not (omitirDatosVaciosONoValidos And AlturaMin.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(AlturaMin.toXML(documentoDondeSeVaAInsertar, zzEtiquetaAlturaMin))
        End If

        If Not (omitirDatosVaciosONoValidos And AlturaMax.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(AlturaMax.toXML(documentoDondeSeVaAInsertar, zzEtiquetaAlturaMax))
        End If

        If Not (omitirDatosVaciosONoValidos And Espesor.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(Espesor.toXML(documentoDondeSeVaAInsertar, zzEtiquetaEspesor))
        End If

        If Not (omitirDatosVaciosONoValidos And Cortante.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(Cortante.toXML(documentoDondeSeVaAInsertar, zzEtiquetaCortante))
        End If

        If Not (omitirDatosVaciosONoValidos And Inercia.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(Inercia.toXML(documentoDondeSeVaAInsertar, zzEtiquetaInercia))
        End If

        If Not (omitirDatosVaciosONoValidos And Momento.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(Momento.toXML(documentoDondeSeVaAInsertar, zzEtiquetaMomento))
        End If

        If Not (omitirDatosVaciosONoValidos And Densidad.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(Densidad.toXML(documentoDondeSeVaAInsertar, zzEtiquetaDensidad))
        End If

        If Not (omitirDatosVaciosONoValidos And ResisFlexionPerpen.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(ResisFlexionPerpen.toXML(documentoDondeSeVaAInsertar, zzEtiquetaResisFlexionPerpen))
        End If

        If Not (omitirDatosVaciosONoValidos And ResisFlexionParal.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(ResisFlexionParal.toXML(documentoDondeSeVaAInsertar, zzEtiquetaResisFlexionParal))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloElasticoMedioPerpen.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(ModuloElasticoMedioPerpen.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloElasticoMedioPerpen))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloElasticoMedioParal.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(ModuloElasticoMedioParal.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloElasticoMedioParal))
        End If

        If Not (omitirDatosVaciosONoValidos And DiametroTuboInterior.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(DiametroTuboInterior.toXML(documentoDondeSeVaAInsertar, zzEtiquetaDiametroTuboInterior))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloYoung.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(ModuloYoung.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloYoung))
        End If

        If Not (omitirDatosVaciosONoValidos And AreaCortante.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(AreaCortante.toXML(documentoDondeSeVaAInsertar, zzEtiquetaAreaCortante))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloResistente.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(ModuloResistente.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloResistente))
        End If

        If Not (omitirDatosVaciosONoValidos And ResisCortantePerpend.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(ResisCortantePerpend.toXML(documentoDondeSeVaAInsertar, zzEtiquetaResisCortantePerpend))
        End If

        If Not (omitirDatosVaciosONoValidos And ResisCortanteParal.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
             nDatosadicionales.AppendChild(ResisCortanteParal.toXML(documentoDondeSeVaAInsertar, zzEtiquetaResisCortanteParal))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloRigidezMedioPerpend.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(ModuloRigidezMedioPerpend.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloRigidezMedioPerpend))
        End If

        If Not (omitirDatosVaciosONoValidos And ModuloRigidezMedioParalelo.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nDatosadicionales.AppendChild(ModuloRigidezMedioParalelo.toXML(documentoDondeSeVaAInsertar, zzEtiquetaModuloRigidezMedioParalelo))
        End If

        Return nDatosadicionales
    End Function


    Sub fromXML(ByVal LeerDatosDesde As XmlNode, ByVal CodElementoAlQueAsignarlos As String, ByVal TipoDeElemento As Elemento.TiposDeElemento)
        If LeerDatosDesde.Name = zzXMLEtiquetaNodoDatosadicionales Then

            Reset()

            CodElemento = CodElementoAlQueAsignarlos
            TipoDeElemento = TipoDeElemento

            Dim nodo As XmlNode
            For Each nodo In LeerDatosDesde.ChildNodes
                Select Case nodo.Name

                  Case zzEtiquetaPeso
                      Peso = New Magnitud(nodo)

                  Case zzEtiquetaAreaEncofrado
                      AreaEncofrado = New Magnitud(nodo)

                  Case zzEtiquetaMaterial
                        Material = nodo.InnerText

                    Case zzEtiquetaUnidadStock
                        UnidadStock = nodo.InnerText

                    Case zzEtiquetaLongitud
                        Longitud = New Magnitud(nodo)

                  Case zzEtiquetaAnchura
                      Anchura.valor = Double.Parse(nodo.InnerText, Globalization.CultureInfo.InvariantCulture)

                  Case zzEtiquetaArea
                      Area = New Magnitud(nodo)

                  Case zzEtiquetaAltura
                      Altura = New Magnitud(nodo)

                  Case zzEtiquetaAlturaMin
                      AlturaMin = New Magnitud(nodo)

                  Case zzEtiquetaAlturaMax
                      AlturaMax = New Magnitud(nodo)

                  Case zzEtiquetaEspesor
                      Espesor = New Magnitud(nodo)

                  Case zzEtiquetaCortante
                      Cortante = New Magnitud(nodo)

                  Case zzEtiquetaInercia
                      Inercia = New Magnitud(nodo)

                  Case zzEtiquetaMomento
                      Momento = New Magnitud(nodo)

                  Case zzEtiquetaDensidad
                      Densidad = New Magnitud(nodo)

                  Case zzEtiquetaResisFlexionPerpen
                      ResisFlexionPerpen = New Magnitud(nodo)

                  Case zzEtiquetaResisFlexionParal
                      ResisFlexionParal = New Magnitud(nodo)

                  Case zzEtiquetaModuloElasticoMedioPerpen
                      ModuloElasticoMedioPerpen = New Magnitud(nodo)

                  Case zzEtiquetaModuloElasticoMedioParal
                      ModuloElasticoMedioParal = New Magnitud(nodo)

                  Case zzEtiquetaDiametroTuboInterior
                      DiametroTuboInterior = New Magnitud(nodo)

                  Case zzEtiquetaModuloYoung
                      ModuloYoung = New Magnitud(nodo)

                  Case zzEtiquetaGenerico
                      Generico = nodo.InnerText

                  Case zzEtiquetaAreaCortante
                      AreaCortante = New Magnitud(nodo)

                  Case zzEtiquetaModuloResistente
                      ModuloResistente = New Magnitud(nodo)

                  Case zzEtiquetaResisCortantePerpend
                      ResisCortantePerpend = New Magnitud(nodo)

                  Case zzEtiquetaResisCortanteParal
                      ResisCortanteParal = New Magnitud(nodo)

                  Case zzEtiquetaModuloRigidezMedioPerpend
                      ModuloRigidezMedioPerpend = New Magnitud(nodo)

                  Case zzEtiquetaModuloRigidezMedioParalelo
                      ModuloRigidezMedioParalelo = New Magnitud(nodo)

                End Select
            Next

        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerDatosDesde.Name & "). El nombre del nodo ha de ser " & zzXMLEtiquetaNodoDatosadicionales)
        End If
    End Sub






    Public Overrides Function Equals(ByVal otroObjeto As Object) As Boolean Implements IEquatable(Of Object).Equals
        If otroObjeto.GetType.Equals(Me.GetType) Then
            Dim otroElemento = CType(otroObjeto, ElementoDatosadicionales)
            If otroElemento.TipoElemento.Equals(Me.TipoElemento) _
            And otroElemento.CodElemento.Equals(Me.CodElemento) _
            And otroElemento.Peso.Equals(Me.Peso) _
            And otroElemento.AreaEncofrado.Equals(Me.AreaEncofrado) _
            And otroElemento.Material.Equals(Me.Material) _
            And otroElemento.UnidadStock.Equals(Me.UnidadStock) _
            And otroElemento.Generico.Equals(Me.Generico) _
            And otroElemento.Longitud.Equals(Me.Longitud) _
            And otroElemento.Anchura.Equals(Me.Anchura) _
            And otroElemento.Area.Equals(Me.Area) _
            And otroElemento.Altura.Equals(Me.Altura) _
            And otroElemento.AlturaMin.Equals(Me.AlturaMin) _
            And otroElemento.AlturaMax.Equals(Me.AlturaMax) _
            And otroElemento.Espesor.Equals(Me.Espesor) _
            And otroElemento.Cortante.Equals(Me.Cortante) _
            And otroElemento.Inercia.Equals(Me.Inercia) _
            And otroElemento.Momento.Equals(Me.Momento) _
            And otroElemento.Densidad.Equals(Me.Densidad) _
            And otroElemento.ResisFlexionPerpen.Equals(Me.ResisFlexionPerpen) _
            And otroElemento.ResisFlexionParal.Equals(Me.ResisFlexionParal) _
            And otroElemento.ModuloElasticoMedioPerpen.Equals(Me.ModuloElasticoMedioPerpen) _
            And otroElemento.ModuloElasticoMedioParal.Equals(Me.ModuloElasticoMedioParal) _
            And otroElemento.DiametroTuboInterior.Equals(Me.DiametroTuboInterior) _
            And otroElemento.ModuloYoung.Equals(Me.ModuloYoung) _
            And otroElemento.AreaCortante.Equals(Me.AreaCortante) _
            And otroElemento.ModuloResistente.Equals(Me.ModuloResistente) _
            And otroElemento.ResisCortantePerpend.Equals(Me.ResisCortantePerpend) _
            And otroElemento.ResisCortanteParal.Equals(Me.ResisCortanteParal) _
            And otroElemento.ModuloRigidezMedioPerpend.Equals(Me.ModuloRigidezMedioPerpend) _
            And otroElemento.ModuloRigidezMedioParalelo.Equals(Me.ModuloRigidezMedioParalelo) _
            Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Public Function CompareTo(ByVal otroElemento As ElementoDatosadicionales) As Integer Implements IComparable(Of ElementoDatosadicionales).CompareTo
        If otroElemento.CodElemento.Equals(Me.CodElemento) Then
            Return 0
        ElseIf otroElemento.CodElemento < Me.CodElemento Then
            Return 1
        Else
            Return -1
        End If
    End Function



End Class
