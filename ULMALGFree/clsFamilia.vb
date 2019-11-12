Public Class clsFamilia
    ''
    Public QUEESTADO As ESTADO = ESTADO.NOENCONTRADO
    Public QUETIPO As Tipo = Tipo.ESTATICO
    Public colCV As Hashtable               ' Key=Nombre propiedad, Value=Valor
    Public arrCriterios As ArrayList        ' Criterios para buscar estas propiedades en Familiinstances
    Public arrPropEscribir As ArrayList     ' Las propiedades escribiremos en FamiliInstances.
    'Public count As Integer = 0             ' Cantidad total de esta FamilyInstance
    Public Shared totalFamilyInstances As Integer          ' Cantidad total de todas las FamilyInstances en el dibujo (De todos los tipos)
    'Public Shared totalPeso As Integer
    Public Shared colProp As Hashtable = Nothing           ' Key=Nº columna, Value=Nombre de la propiedad
    Public fila As Integer = -1
    Public modificado As Boolean = True        '' Para controlar si se han modificado los datos iniciales. No procesamos este objeto si modificado=false
    Public ITEM_LENGTH_UNIT As String = "mm"
    Public ITEM_AREA_UNIT As String = "m2"
    Public ITEM_WEIGHT_UNIT As String = "kg"
    Public ITEM_GENERIC_UNIT As String = "kg"
    ''
    Public Sub New(queFila As Integer, colColumnas As Hashtable)
        fila = queFila
        colCV = New Hashtable          '' Hashtable el nombre de las propiedades y su valor (key=parametro, value=valor parametro)
        '' Solo ponemos las propiedades que se escribirán.
        colCV.Add("TYPE", "")                '' colCV.Add(nType.ToUpper, "")
        colCV.Add("FAMILY_CODE", "")
        colCV.Add("ITEM_CODE", "")
        colCV.Add("ITEM_CODE_L", "")
        colCV.Add("ITEM_LENGTH", "")
        colCV.Add("ITEM_LENGTH_SUMA", "")
        colCV.Add("ITEM_WIDTH", "")
        colCV.Add("ITEM_HEIGHT", "")
        colCV.Add("ITEM_LENGTH1", "")
        colCV.Add("ITEM_WIDTH1", "")
        colCV.Add("ITEM_HEIGHT1", "")

        colCV.Add("ITEM_WEIGHT", "")
        colCV.Add("ITEM_WEIGHT_L", "")
        colCV.Add("WEIGHT", "")

        colCV.Add("ITEM_DESCRIPTION", "")
        colCV.Add("COUNT", "")               '' colCV.Add(nCount.ToUpper, "")
        colCV.Add("ITEM_MARKET", 0)
        colCV.Add("W_MARKET", "")
        colCV.Add("ITEM_GENERIC", "")               '' Poner 1 si no tiene largo, ancho, alto.
        colCV.Add("AREA", "")
        colCV.Add("AREA_UNID", "")
        colCV.Add("PESOFIN", "")
        colCV.Add("PESOFIN_SUMA", "")
        colCV.Add("AREAFIN", "")
        colCV.Add("AREAFIN_SUMA", "")
        '' Llenamos colProp con colColumnas, si no existe
        If colProp Is Nothing Then colProp = New Hashtable
        If colProp.Count = 0 Then
            For Each queCol As Integer In colColumnas.Keys
                If queCol = 3 Then
                    colProp.Add(3, "COUNT")
                ElseIf queCol = 5 Then
                    colProp.Add(5, "TYPE")
                Else
                    colProp.Add(queCol, colColumnas(queCol))
                End If
            Next
        End If
        ''
        '' TYPE (nType), FAMILY_CODE, ITEM_CODE, ITEM_WEIGHT, ITEM_LENGTH, ITEM_WIDTH, ITEM_HEIGHT, ITEM_WEIGHT, ITEM_DESCRIPTION, COUNT (nCount), W_MARKET, ITEM_GENERIC
    End Sub
    ''
    Public Sub PonCriterios(queCriterios As String())
        If arrCriterios Is Nothing Then
            arrCriterios = New ArrayList
        Else
            arrCriterios.Clear()
        End If
        ''
        For Each queCriterio As String In queCriterios
            If arrCriterios.Contains(queCriterio) = False AndAlso colCV.ContainsKey(queCriterio) Then
                arrCriterios.Add(queCriterio)
            End If
        Next
    End Sub
    ''
    Public Sub PonPropEscribir(queProps As String())
        If arrPropEscribir Is Nothing Then
            arrPropEscribir = New ArrayList
        Else
            arrPropEscribir.Clear()
        End If
        ''
        For Each queProp As String In queProps
            If arrPropEscribir.Contains(queProp) = False AndAlso colCV.ContainsKey(queProp) Then
                arrPropEscribir.Add(queProp)
            End If
        Next
    End Sub
    ''
    Public Sub PonDato(queCol As Integer, queValor As String, Optional control As Boolean = False)
        ''
        If colProp.ContainsKey(queCol) = False Then Exit Sub
        ''
        If control = False Then
            '' Convertir las unidades leidas de la lista de piezas BOM-WK, según las unidades de Revit.
            ''
            '' Cambiar a mm largo, ancho, alto.
            If (colProp(queCol).ToString = "ITEM_LENGTH" OrElse
                colProp(queCol).ToString = "ITEM_WIDTH" OrElse
                colProp(queCol).ToString = "ITEM_HEIGHT") AndAlso
                (queValor.Trim <> "" AndAlso IsNumeric(queValor.Trim)) = True Then
                ''
                queValor = Unidades_DameMilimetros(queValor)
            End If
            '' En la carga inicial de la clase, modificiado = false y control = false. Escribimos los valores leidos de BOM-WK
            If colCV.ContainsKey(colProp(queCol).ToString.ToUpper) Then
                colCV(colProp(queCol)) = queValor
            End If
        Else
            If colProp(queCol).ToString = "ITEM_WEIGHT" Then
                '' Si control=true. Comprobamos antes el dato y pondremos modificado=true en caso de que fuera diferentes y lo tengamos que escribir.
                If colCV.ContainsKey(colProp(queCol).ToString.ToUpper) AndAlso ValoresIgualesPeso(queValor) = False Then
                    If queValor <> "" AndAlso queValor.Contains(" ") Then
                        Try
                            Dim valorfin() As String = queValor.Split(" "c)
                            colCV(colProp(queCol)) = CDbl(FormatNumber(valorfin(0).ToString, 2)) & " " & valorfin(1)
                        Catch ex As Exception
                            Debug.Print(queValor)
                        End Try
                    Else
                        colCV(colProp(queCol)) = queValor
                    End If
                    '' Con 1 sólo dato que modifiquemos. Pondremos modificado=true
                    If arrPropEscribir.Contains(colProp(queCol)) Then
                        Me.modificado = True
                    End If
                End If
            Else
                '' Si control=true. Comprobamos antes el dato y pondremos modificado=true en caso de que fuera diferentes y lo tengamos que escribir.
                If colCV.ContainsKey(colProp(queCol).ToString.ToUpper) AndAlso colCV(colProp(queCol)).ToString <> queValor Then
                    colCV(colProp(queCol)) = queValor
                    '' Con 1 sólo dato que modifiquemos. Pondremos modificado=true
                    If arrPropEscribir.Contains(colProp(queCol)) Then
                        Me.modificado = True
                    End If
                End If
            End If
        End If
        ''
        '' Aumentamos total (variable compartida por todos los objetos de la clase)
        '' total = total de elementos dibujo  /  colCV("COUNT") = cantidad de elementos sólo de este objeto (Columna Count de BOM-WK)
        If colProp(queCol).ToString.ToUpper = nCount.ToUpper AndAlso IsNumeric(queValor) Then
            totalFamilyInstances += CInt(queValor)
        End If
    End Sub
    ''
    Public Function ValoresIgualesPeso(queValor As String) As Boolean
        '' Valor actual del peso guardado en colCV("ITEM_WEIGHT")
        Dim valorenclase As String = colCV("ITEM_WEIGHT").ToString
        'Try
        '    If valorenclase <> "" AndAlso (valorenclase.Contains(" ") AndAlso valorenclase.Substring(0, 1) <> " ") Then
        '        valorenclase = colCV("ITEM_WEIGHT").ToString.Split(" "c)(0).Replace(",", ".")
        '    End If
        'Catch ex As Exception
        '    valorenclase = ""
        'End Try
        '' Le quitamos puntos, comas y espacios. Para compararlo con el valor que leemos.
        valorenclase = valorenclase.Replace(",", "").Replace(".", "").Replace(" ", "")
        Dim valorrecibo As String = queValor.Replace(",", "").Replace(".", "").Replace(" ", "")
        '' Esta variable sólo para el peso. Ya que puede venir con decimales.
        'Dim valores() As String = queValor.Split(" "c)
        'Try
        '    If valores.Count = 2 AndAlso valores(0).ToString <> "" Then
        '        queValor = valores(0).Replace(",", ".")
        '    Else
        '        queValor = ""
        '    End If
        'Catch ex As Exception
        '    Debug.Print(ex.Message)
        'End Try
        ' ''
        'Dim v1 As Double = 0.0
        'Dim v2 As Double = 0.0
        ' ''
        'If IsNumeric(valorenclase) Then v1 = CDbl(valorenclase)
        'If IsNumeric(queValor) Then v2 = CDbl(queValor)
        ' ''
        'If v1 = v2 Then
        '    Return True
        'Else
        '    Return False
        'End If
        If valorenclase = valorrecibo Then
            Return True
        Else
            Return False
        End If
    End Function
    ''
    Public Sub PonTipo(queTipo As Tipo)
        '' TYPE, FAMILY_CODE, ITEM_CODE, ITEM_LENGTH, ITEM_WIDTH, ITEM_HEIGHT, ITEM_WEIGHT, FILTER_ID, ITEM_DESCRIPTION, COUNT, W_MARKET
        Dim arrCriterios() As String = Nothing
        Dim arrPropEscribir() As String = Nothing
        Select Case queTipo
            Case Tipo.ESTATICO
                arrCriterios = New String() {"ITEM_CODE"}
                arrPropEscribir = New String() {"ITEM_DESCRIPTION", "ITEM_MARKET", "W_MARKET", "ITEM_GENERIC", "ITEM_CODE_L", "ITEM_WEIGHT", "ITEM_WEIGHT_L", "WEIGHT"}    ' Incluimos WEIGHT 
                Me.QUETIPO = queTipo
                Me.colCV("ITEM_GENERIC") = ""
            Case Tipo.DINAMICO
                arrCriterios = New String() {"FAMILY_CODE", "ITEM_LENGTH", "ITEM_WIDTH", "ITEM_HEIGHT"}
                arrPropEscribir = New String() {"ITEM_CODE", "ITEM_MARKET", "ITEM_DESCRIPTION", "W_MARKET", "ITEM_GENERIC", "ITEM_CODE_L", "ITEM_WEIGHT", "ITEM_WEIGHT_L", "WEIGHT"}    ' Incluimos WEIGHT
                'arrPropEscribir = New String() {"ITEM_CODE", "PESOFIN", "ITEM_DESCRIPTION", "W_MARKET"}, "ITEM_CODE_L"
                Me.QUETIPO = queTipo
                'Me.colCV("ITEM_GENERIC") = "0"
            Case Tipo.GENERICO
                arrCriterios = New String() {"FAMILY_CODE"}
                arrPropEscribir = New String() {"ITEM_CODE", "ITEM_MARKET", "ITEM_DESCRIPTION", "W_MARKET", "ITEM_GENERIC", "ITEM_CODE_L", "ITEM_WEIGHT", "ITEM_WEIGHT_L", "WEIGHT"}    ' Incluimos WEIGHT
                'arrPropEscribir = New String() {"ITEM_CODE", "PESOFIN", "ITEM_DESCRIPTION", "W_MARKET"}, "ITEM_CODE_L"
                Me.QUETIPO = queTipo
                'Me.colCV("ITEM_GENERIC") = "1"  '' Sera 1 (Lineales) o 2 (Area)  ** Lo cambiaremos a 2 si tiene area.
            Case Tipo.CONERROR
                arrCriterios = New String() {}
                arrPropEscribir = New String() {}
                'Me.PonCriterios(arrCriterios)
                'Me.PonPropEscribir(arrPropEscribir)
                Me.QUEESTADO = ESTADO.CONERRORES
                Me.colCV("ITEM_GENERIC") = ""
        End Select
        ''
        Me.PonCriterios(arrCriterios)
        Me.PonPropEscribir(arrPropEscribir)
    End Sub
    ''
    '' ***** Calculamos el peso final de FamilyInstante
    Public Sub PonPesoFin()
        If colCV("ITEM_WEIGHT").ToString = "" Then Exit Sub
        ''
        PonArea()
        '
        If colCV("AREAFIN").ToString = "" Then
            ' ITEM_GENERIC es 0
            '' Si no tenemos Largo ni Alto salimos. Habrá que ver cual de los 2 tenemos
            If colCV("ITEM_LENGTH").ToString = "" AndAlso colCV("ITEM_WIDTH").ToString = "" AndAlso colCV("ITEM_HEIGHT").ToString = "" Then
                colCV("PESOFIN") = colCV("ITEM_WEIGHT").ToString & " " & Me.ITEM_GENERIC_UNIT
                colCV("WEIGHT") = colCV("ITEM_WEIGHT").ToString & " " & Me.ITEM_GENERIC_UNIT
                Exit Sub
            End If
            ''
            Dim LargoAlto As Double = 0     '' Largo, Alto y, si no tienen, Ancho.
            If colCV("ITEM_LENGTH").ToString <> "" AndAlso IsNumeric(colCV("ITEM_LENGTH").ToString) Then
                LargoAlto = CDbl(colCV("ITEM_LENGTH").ToString)
            ElseIf colCV("ITEM_HEIGHT").ToString <> "" AndAlso IsNumeric(colCV("ITEM_HEIGHT").ToString) Then
                LargoAlto = CDbl(colCV("ITEM_HEIGHT").ToString)
            ElseIf colCV("ITEM_WIDTH").ToString <> "" AndAlso IsNumeric(colCV("ITEM_WIDTH").ToString) Then
                LargoAlto = CDbl(colCV("ITEM_WIDTH").ToString)
            End If
            '
            ' ITEM_GENERIC es 1. O no lleva area o es un elemento (ML)
            If colCV("ITEM_DESCRIPTION").ToString.ToUpper.Contains("(ML)") OrElse colCV("ITEM_GENERIC").ToString = "1" Then
                '' Articulo con (ML) Calcular el peso en función de la longitud
                Dim solopeso As String = colCV("ITEM_WEIGHT").ToString.Split(" "c)(0)
                ''
                If solopeso = "" OrElse IsNumeric(solopeso) = False Then Exit Sub
                '' *** CALCULAR LONGITUD segun genericArticleUnit
                Dim largounit As Double
                If Me.ITEM_GENERIC_UNIT = "" Then
                    largounit = CDbl(FormatNumber(LargoAlto / 1000, 2))    '' De milimetros a metros
                Else
                    largounit = CDbl(FormatNumber(U_DameString__DesdeCualquiera(LargoAlto.ToString, ITEM_LENGTH_UNIT, Me.ITEM_GENERIC_UNIT), 2))
                End If
                Dim pesofin As Double = (CDbl(solopeso) * largounit)
                colCV("PESOFIN") = FormatNumber(pesofin, 2) & " " & Me.ITEM_GENERIC_UNIT
                colCV("WEIGHT") = FormatNumber(pesofin, 2) & " " & Me.ITEM_GENERIC_UNIT
                'Else
                '    '' Articulo normal, no lleva (ML) ni lleva area
                '    colCV("PESOFIN") = colCV("ITEM_WEIGHT")
            End If
            '' Pondremos modificado=true, si ha cambiado el peso
            If colCV("ITEM_WEIGHT").ToString <> colCV("PESOFIN").ToString Then
                Me.modificado = True
            End If
        Else
            ' ITEM_GENERIC es 2
            ''
            '' Calculamos el peso, en función del area.
            Dim areafin As Double = CDbl(colCV("AREAFIN").ToString.Split(" "c)(0))
            '' Regla de 3 para calcular el peso. Si AREA m2 = ITEM_WEIGHT, XXXX m2 = PESOFIN
            '' PESOFIN = (AREAFIN * ITEM_WEIGHT) / AREA
            Dim solopeso As String = colCV("ITEM_WEIGHT").ToString.Split(" "c)(0)
            ''
            If solopeso = "" OrElse IsNumeric(solopeso) = False Then Exit Sub
            '' *** CALCULAR AREA segun genericArticleUnit (El area se ha calculado en m2)
            Dim areaunit As Double
            If Me.ITEM_GENERIC_UNIT = "" Then
                areaunit = (areafin * CDbl(solopeso)) '/ CDbl(colCV("AREA").ToString)
            Else
                areaunit = CDbl(FormatNumber(U_DameString__DesdeCualquiera(areafin.ToString, "M2", Me.ITEM_GENERIC_UNIT), 2))
            End If
            Dim pesofin As Double = (CDbl(solopeso) * areaunit)
            colCV("PESOFIN") = FormatNumber(pesofin, 2) & " " & Me.ITEM_GENERIC_UNIT
            colCV("WEIGHT") = FormatNumber(pesofin, 2) & " " & Me.ITEM_GENERIC_UNIT
        End If
        '' Pondremos modificado=true, si ha cambiado el peso
        If colCV("ITEM_WEIGHT").ToString <> colCV("PESOFIN").ToString Then
            Me.modificado = True
        End If
    End Sub
    ''
    Public Sub PonArea()
        '' Si no tenemos unidades de area. Salimos sin hacer nada.
        If Me.colCV("AREA_UNID").ToString = "" OrElse
            Me.colCV("AREA_UNID").ToString = "0" Then
            Exit Sub
        End If
        '' Si faltan datos para calcular el area, tambien salimos.
        If Me.colCV("ITEM_LENGTH").ToString = "" OrElse
            Me.colCV("ITEM_LENGTH").ToString = "0" OrElse
            Me.colCV("ITEM_WIDTH").ToString = "" OrElse
            Me.colCV("ITEM_WIDTH").ToString = "0" OrElse
            Me.colCV("AREA").ToString = "" Then
            Exit Sub
        End If
        ''
        '' Calculamos el area (en m2. Dividir para 1000 porque viene en mm.)
        Dim area As Double = CDbl(colCV("ITEM_LENGTH").ToString) / 1000 * CDbl(colCV("ITEM_WIDTH").ToString) / 1000
        'Dim areaTemp As String = utilesRevitUnidades.U_DameString__DesdeCualquiera(area.ToString, "m2", Me.colCV("AREA_UNID").ToString)
        'area = CDbl(areaTemp)
        '' Area con 2 decimales "#.00"
        colCV("AREAFIN") = area.ToString("#.00", System.Globalization.CultureInfo.InvariantCulture) '& " " & Me.colCV("AREA_UNID").ToString
        '' Area (2) - Lineales (1) ** Llevan el 1 por defecto los genericos.
        ' Me.colCV("ITEM_GENERIC") = "2"
    End Sub
    ''
    '' Nos devolverá los totales multiplicados
    '' PesoTotal = (Me.colCV("COUNT") * Me.colCV("ITEM_WEIGHT")) 
    Public Sub PonTotalesGenerales()
        Dim nFam As Integer = CInt(colCV(nCount.ToUpper))
        Dim lTot As Integer = 0
        Dim aTot As Double = 0.0
        Dim pTot As Double = 0.0
        '' Largo total
        If colCV("ITEM_LENGTH").ToString <> "" AndAlso IsNumeric(colCV("ITEM_LENGTH")) Then
            lTot = CInt(colCV("ITEM_LENGTH")) * nFam
            colCV("ITEM_LENGTH_SUMA") = lTot.ToString
        End If
        '' Area total
        Dim aT As String = colCV("AREAFIN").ToString.Trim.Split(" "c)(0)
        If aT <> "" AndAlso IsNumeric(aT) Then
            aTot = CDbl(aT) * nFam
            colCV("AREAFIN_SUMA") = aTot.ToString
            '' Area (2) - Lineales (1) ** Llevan el 1 por defecto los genericos.
            'Me.colCV("ITEM_GENERIC") = "2"
        End If
        '' Peso total
        Dim pT As String = colCV("PESOFIN").ToString.Trim.Split(" "c)(0)
        If pT <> "" AndAlso IsNumeric(pT) Then
            pTot = CDbl(pT) * nFam
            'colCV("PESOFIN_SUMA") = pTot.ToString
        End If
    End Sub
    ''
    Public Enum ESTADO
        CONERRORES
        ENCONTRADO
        NOENCONTRADO
    End Enum
    ''
    Public Enum Tipo
        ESTATICO        '' FAMILY_CODE vacío, SPECIAL = NO y SPECIAL = YES 
        DINAMICO        '' FAMILY_CODE relleno, con valor en ITEM_LENGHT y/o ITEM_WIDTH y/o ITEM_HEIGHT
        GENERICO        '' FAMILY_CODE relleno, vacíos ITEM_LENGHT y ITEM_WIDTH y ITEM_HEIGHT
        CONERROR           '' FAMILY_CODE vacio e ITEM_CODE vacio  (Esto es un error en la familia)
    End Enum
End Class

