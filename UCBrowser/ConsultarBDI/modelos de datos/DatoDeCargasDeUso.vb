Imports System.Xml
Imports Magnitudes


Public Class DatoDeCargasDeUso : Implements IEquatable(Of Object) : Implements IComparable(Of DatoDeCargasDeUso)

    Public TipoElemento As Elemento.TiposDeElemento
    Public CodElemento As String
    Public Altura As Magnitud
    Public TuboInterior As OrientacionDelTuboInterior
    Public Carga As Magnitud
    Public CargaPlantaHormigonada As Magnitud
    Public CargaReapuntalado As Magnitud
    Public Rigidez As Magnitud
    Public Viento As Double

    Public Enum OrientacionDelTuboInterior
        indistinto = 0
        arriba = 1
        abajo = 2
    End Enum


    Public Sub New()
        TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
        CodElemento = ""
        Viento = 0.0
        TuboInterior = OrientacionDelTuboInterior.indistinto
        Altura = New Magnitud(Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Carga = New Magnitud(Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        CargaPlantaHormigonada = New Magnitud(Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        CargaReapuntalado = New Magnitud(Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
        Rigidez = New Magnitud(Double.NaN, unidaddemedida:=Magnitud.DESCONOCIDA)
    End Sub

    Public Sub New(cargaDeLaQueCopiar As DatoDeCargasDeUso)
        Me.TipoElemento = cargaDeLaQueCopiar.TipoElemento
        Me.CodElemento = cargaDeLaQueCopiar.CodElemento
        Me.Viento = cargaDeLaQueCopiar.Viento
        Me.TuboInterior = cargaDeLaQueCopiar.TuboInterior
        Me.Altura = cargaDeLaQueCopiar.Altura
        Me.Carga = cargaDeLaQueCopiar.Carga
        Me.CargaPlantaHormigonada = cargaDeLaQueCopiar.CargaPlantaHormigonada
        Me.CargaReapuntalado = cargaDeLaQueCopiar.CargaReapuntalado
        Me.Rigidez = cargaDeLaQueCopiar.Rigidez
    End Sub


    Public Overrides Function ToString() As String
        Return "{" & TipoElemento & "[" & CodElemento & "] ti " & TuboInterior.ToString & "  ==> " & Altura.ToString & " " & Carga.ToString("F4") & " }"
    End Function


    Public Const zzXML_EtiquetaNodo As String = "workingload"

    Public Const zzEtiquetaElemento As String = "element"
    Public Const zzAtributoTipo As String = "type"
    Public Const zzEtiquetaElementoCodigo As String = "elementCode"
    Public Const zzEtiquetaAltura As String = "height"
    Public Const zzEtiquetaTuboInterior As String = "internalTube"
    Public Const zzEtiquetaCarga As String = "load"
    Public Const zzEtiquetaCargaPlantaHormigonada As String = "afterpouringLoad"
    Public Const zzEtiquetaCargaReapuntalado As String = "repropLoad"
    Public Const zzEtiquetaRigidez As String = "stiffness"
    Public Const zzEtiquetaViento As String = "wind"

    Function toXML(ByVal documentoDondeSeVaAInsertar As XmlDocument, ByVal omitirDatosVaciosONoValidos As Boolean) As XmlNode

        Dim nCargaDeUso As XmlNode
        nCargaDeUso = documentoDondeSeVaAInsertar.CreateElement(zzXML_EtiquetaNodo)

        Dim nElemento As XmlNode
        nElemento = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaElemento)
        nElemento.InnerText = CodElemento
        Dim aTipoElemento As XmlAttribute
        aTipoElemento = documentoDondeSeVaAInsertar.CreateAttribute(zzAtributoTipo)
        aTipoElemento.InnerText = CType(TipoElemento, Integer).ToString
        nElemento.Attributes.Append(aTipoElemento)
        nCargaDeUso.AppendChild(nElemento)

        Dim nTuboInterior As XmlNode
        nTuboInterior = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaTuboInterior)
        nTuboInterior.InnerText = TuboInterior.ToString
        nCargaDeUso.AppendChild(nTuboInterior)

        nCargaDeUso.AppendChild(Altura.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaAltura))

        nCargaDeUso.AppendChild(Carga.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaCarga))

        If Not (omitirDatosVaciosONoValidos And CargaPlantaHormigonada.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nCargaDeUso.AppendChild(CargaPlantaHormigonada.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaCargaPlantaHormigonada))
        End If

        If Not (omitirDatosVaciosONoValidos And CargaReapuntalado.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nCargaDeUso.AppendChild(CargaReapuntalado.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaCargaReapuntalado))
        End If

        If Not (omitirDatosVaciosONoValidos And Rigidez.unidaddemedida.Equals(Magnitud.DESCONOCIDA)) Then
            nCargaDeUso.AppendChild(Rigidez.ToXML(documentoDondeSeVaAInsertar, zzEtiquetaRigidez))
        End If

        Dim nViento As XmlNode
        nViento = documentoDondeSeVaAInsertar.CreateElement(zzEtiquetaViento)
        nViento.InnerText = Viento.ToString(Globalization.CultureInfo.InvariantCulture)
        nCargaDeUso.AppendChild(nViento)

        Return nCargaDeUso
    End Function


    Sub fromXML(ByVal LeerCargaDeUsoDesde As XmlNode)
        If LeerCargaDeUsoDesde.Name = zzXML_EtiquetaNodo Then
            Dim nodo As XmlNode
            For Each nodo In LeerCargaDeUsoDesde.ChildNodes
                Select Case nodo.Name

                    Case zzEtiquetaElemento
                        CodElemento = nodo.InnerText
                        Dim tipo As Integer
                        Integer.TryParse(nodo.Attributes.GetNamedItem(zzAtributoTipo).InnerText, tipo)
                        If [Enum].IsDefined(GetType(Elemento.TiposDeElemento), tipo) Then
                            TipoElemento = CType([Enum].Parse(GetType(Elemento.TiposDeElemento), _
                                                              nodo.Attributes.GetNamedItem(zzAtributoTipo).InnerText), 
                                                 Elemento.TiposDeElemento)
                        Else
                            TipoElemento = Elemento.TiposDeElemento.DESCONOCIDO
                        End If

                    Case zzEtiquetaTuboInterior
                        TuboInterior = CType([Enum].Parse(GetType(OrientacionDelTuboInterior), nodo.InnerText), OrientacionDelTuboInterior)

                    Case zzEtiquetaAltura
                        Altura = New Magnitud(nodo)

                    Case zzEtiquetaCarga
                        Carga = New Magnitud(nodo)

                    Case zzEtiquetaCargaPlantaHormigonada
                        CargaPlantaHormigonada = New Magnitud(nodo)

                    Case zzEtiquetaCargaReapuntalado
                        CargaReapuntalado = New Magnitud(nodo)

                    Case zzEtiquetaRigidez
                        Rigidez = New Magnitud(nodo)

                    Case zzEtiquetaViento
                        Viento = Double.Parse(nodo.InnerText, Globalization.CultureInfo.InvariantCulture)

                End Select
            Next
        Else
            Throw New ArgumentException("Se ha recibido un nodo erroneo (" & LeerCargaDeUsoDesde.Name & "). El nombre del nodo ha de ser " & zzXML_EtiquetaNodo)
        End If
    End Sub



    Public Overrides Function Equals(ByVal otroObjeto As Object) As Boolean Implements IEquatable(Of Object).Equals
        If otroObjeto.GetType.Equals(Me.GetType) Then
            Dim otraCargaDeUso = CType(otroObjeto, DatoDeCargasDeUso)
            If otraCargaDeUso.TipoElemento.Equals(Me.TipoElemento) _
            And otraCargaDeUso.CodElemento.Equals(Me.CodElemento) _
            And otraCargaDeUso.Altura.Equals(Me.Altura) _
            And otraCargaDeUso.TuboInterior.Equals(Me.TuboInterior) _
            And otraCargaDeUso.Carga.Equals(Me.Carga) _
            And otraCargaDeUso.CargaPlantaHormigonada.Equals(Me.CargaPlantaHormigonada) _
            And otraCargaDeUso.CargaReapuntalado.Equals(Me.CargaReapuntalado) _
            And otraCargaDeUso.Rigidez.Equals(Me.Rigidez) _
            And otraCargaDeUso.Viento.Equals(Me.Viento) _
           Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function


    Public Function CompareTo(ByVal otraCargaDeUso As DatoDeCargasDeUso) As Integer Implements IComparable(Of DatoDeCargasDeUso).CompareTo
        If otraCargaDeUso.Equals(Me) Then
            Return 0
        ElseIf otraCargaDeUso.CodElemento < Me.CodElemento Then
            Return 1
        Else
            Return -1
        End If
    End Function


End Class
