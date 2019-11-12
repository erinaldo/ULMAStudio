#Region "Imported Namespaces"
Imports Autodesk.Revit.ApplicationServices
Imports Autodesk.Revit.Attributes
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection
Imports Autodesk.Revit.DB.Architecture

Imports System.Collections.Generic
Imports System.Xaml
Imports System.Diagnostics '– used for debug 
'Imports System.Windows.Forms
Imports System.IO '– used for reading folders 
Imports System.Windows.Media.Imaging '– used for bitmap images 
Imports adWin = Autodesk.Windows
Imports System.Globalization
#End Region


Module utilesRevitUnidades
    Public Function Unidades_DameDouble(queValor As String) As Double
        Dim oCulture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("en-US")
        'Dim culture As System.Globalization.CultureInfo = New System.Globalization.CultureInfo("es-ES")
        Dim resultado As Double = 0
        queValor = queValor.Split(" "c)(0).ToString
        If IsNumeric(queValor) Then
            resultado = Math.Round(Convert.ToDouble(resultado, oCulture), 2)
        Else
            If queValor = "" Then
                ' No hacemos nada. Se devolverá 0
            Else
                If queValor.Contains(" ") Then
                    Dim parteValor As String = queValor.Split(" "c)(0)
                    If IsNumeric(parteValor) AndAlso CDbl(parteValor) = 0 Then
                        ' No hacemos nada. Se devolverá 0
                    ElseIf IsNumeric(parteValor) Then
                        resultado = Math.Round(Convert.ToDouble(parteValor, oCulture), 2)
                    End If
                Else
                    Dim nTemp As String = ""
                    For x As Integer = 0 To queValor.Length - 1
                        Dim oCh As Char = queValor.Chars(x)
                        If IsNumeric(oCh) OrElse oCh.ToString = "," OrElse oCh.ToString = "." Then
                            nTemp &= oCh.ToString
                        End If
                    Next
                    If IsNumeric(nTemp) Then
                        resultado = Math.Round(Convert.ToDouble(nTemp, oCulture), 2)
                    End If
                End If
            End If
        End If
        '
        Return resultado
    End Function
    ''
    '' Crear unidades para Length (mm) y para Weight (Kg) y aplicarlas al documento.
    Public Sub UDocumento_PonMilimetros(ByRef od As Document)
        '' Guardar las Unidades del documento
        unidDoc = od.GetUnits          '' Todas las unidades del documento.
        ''
        '' Unidades en mm y con 2 decimales. Objeto preparado para cambio de unidades.
        If unidDocMM Is Nothing Then
            unidDocMM = New Units(UnitSystem.Metric)
            unidDocMM.SetFormatOptions(UnitType.UT_Length, New FormatOptions(DisplayUnitType.DUT_MILLIMETERS, 0.01))
            unidDocMM.SetFormatOptions(UnitType.UT_Weight, New FormatOptions(DisplayUnitType.DUT_KILOGRAMS_MASS, 0.01))
        End If
        ''
        Dim oTr As Transaction = Nothing
        ''
        'If unidDoc.GetFormatOptions(UnitType.UT_Length).DisplayUnits <> unidDocMM.GetFormatOptions(UnitType.UT_Length).DisplayUnits Then
        If unidDoc.Equals(unidDocMM) = False Then
            Try
                oTr = New Transaction(evRevit.evAppUI.ActiveUIDocument.Document, "CHANGE UNITS DOCUMENT TO MM")
                oTr.Start()
                od.SetUnits(unidDocMM)
                od.Regenerate()
                oTr.Commit()
            Catch ex As Exception
                If oTr IsNot Nothing AndAlso oTr.GetStatus = TransactionStatus.Started Then oTr.RollBack()
            End Try
        End If
        oTr = Nothing
    End Sub
    ''
    Public Sub UDocumento_PonOriginales(ByRef oD As Document)
        ''
        Dim oTr As Transaction = Nothing
        'If oD.GetUnits.GetFormatOptions(UnitType.UT_Length).DisplayUnits <> unidDoc.GetFormatOptions(UnitType.UT_Length).DisplayUnits Then
        If unidDoc IsNot Nothing AndAlso oD.GetUnits.Equals(unidDoc) = False Then
            Try
                oTr = New Transaction(oD, "CHANGE UNITS DOCUMENT TO ORIGINAL UNITS")
                oTr.Start()
                oD.SetUnits(unidDoc)
                oD.Regenerate()
                oTr.Commit()
            Catch ex As Exception
                If oTr IsNot Nothing AndAlso oTr.GetStatus = TransactionStatus.Started Then oTr.RollBack()
            End Try
        End If
        oTr = Nothing
    End Sub
    ''
    ''' <summary>
    ''' Convierte un valor desde una unidad origen X (uniOri) a otra unidad destino Y (uniDes)
    ''' </summary>
    ''' <param name="queValor">Valor a convertir, como texto (Si lleva unidades, las quita)</param>
    ''' <param name="uniOri">Texto de unidades origen</param>
    ''' <param name="uniDes">Texto de unidades destino</param>
    ''' <returns></returns>
    Public Function U_DameString__DesdeCualquiera(queValor As String, uniOri As String, uniDes As String) As String
        ' Solo el texto con el valor numérico
        Dim resultado As String = queValor
        Dim txtValor As String = queValor.Split(" "c)(0)
        '' Si es "" o "0" o no es numérico o son igual uniOri y uniDes devolvemos resultado (Solo numérico) y salimos
        If queValor.Trim = "" OrElse queValor.Trim = "0" OrElse IsNumeric(queValor) = False OrElse uniOri.ToUpper = uniDes.ToUpper OrElse uniOri = "" OrElse uniDes = "" Then
            Return txtValor
            Exit Function
        End If
        ' Convertimos el texto numérico a double
        Dim dblValor As Double = Convert.ToDouble(txtValor)
        '
        Dim enumOri As DisplayUnitType = Nothing
        Dim enumDes As DisplayUnitType = Nothing
        ' Unificar unidades de Area
        If uniOri.EndsWith("2") = True AndAlso uniDes.EndsWith("2") = False Then
            uniDes &= "^2"
        ElseIf uniOri.EndsWith("2") = False AndAlso uniDes.EndsWith("2") = True Then
            uniOri &= "^2"
        End If
        '
        ' DisplayUnitType de uniOri
        Select Case uniOri.ToUpper
            ' Unidades de LONGITUD
            Case "MM"
                enumOri = DisplayUnitType.DUT_MILLIMETERS
            Case "CM"
                enumOri = DisplayUnitType.DUT_CENTIMETERS
            Case "DM"
                enumOri = DisplayUnitType.DUT_DECIMETERS
            Case "M"
                enumOri = DisplayUnitType.DUT_METERS
            Case "IN"
                enumOri = DisplayUnitType.DUT_DECIMAL_INCHES
            Case "F", "FT"
                enumOri = DisplayUnitType.DUT_DECIMAL_FEET
                '
                ' Unidades de PESO
            Case "KG"
                enumOri = DisplayUnitType.DUT_KILOGRAMS_MASS
            Case "LB"
                enumOri = DisplayUnitType.DUT_POUNDS_MASS
                '
                ' Unidades de AREA
            Case "MM2", "MM^2"
                enumOri = DisplayUnitType.DUT_SQUARE_MILLIMETERS
            Case "CM2", "CM^2"
                enumOri = DisplayUnitType.DUT_SQUARE_CENTIMETERS
            'Case "DM2", "DM^2"
            '    enumOri = DisplayUnitType.DUT_SQUARE_DECIMETERS
            Case "M2", "M^2"
                enumOri = DisplayUnitType.DUT_SQUARE_METERS
            Case "IN2", "IN^2"
                enumOri = DisplayUnitType.DUT_SQUARE_INCHES
            Case "F2", "FT2", "F^2", "FT^2"
                enumOri = DisplayUnitType.DUT_SQUARE_FEET
                '
            Case "GR"   ' Unidades de ANGULOS
                enumOri = DisplayUnitType.DUT_GRADS
            Case "RAD"
                enumOri = DisplayUnitType.DUT_RADIANS
        End Select
        '
        ' DisplayUnitType de uniDes
        Select Case uniDes.ToUpper
            ' Unidades de LONGITUD
            Case "MM"
                enumDes = DisplayUnitType.DUT_MILLIMETERS
            Case "CM"
                enumDes = DisplayUnitType.DUT_CENTIMETERS
            Case "DM"
                enumDes = DisplayUnitType.DUT_DECIMETERS
            Case "M"
                enumDes = DisplayUnitType.DUT_METERS
            Case "IN"
                enumDes = DisplayUnitType.DUT_DECIMAL_INCHES
            Case "F", "FT"
                enumDes = DisplayUnitType.DUT_DECIMAL_FEET
                '
                ' Unidades de AREA
            Case "MM2", "MM^2"
                enumDes = DisplayUnitType.DUT_SQUARE_MILLIMETERS
            Case "CM2", "CM^2"
                enumDes = DisplayUnitType.DUT_SQUARE_CENTIMETERS
            'Case "DM2", "DM^2"
            '    enumDes = DisplayUnitType.DUT_SQUARE_DECIMETERS
            Case "M2", "M^2"
                enumDes = DisplayUnitType.DUT_SQUARE_METERS
            Case "IN2", "IN^2"
                enumDes = DisplayUnitType.DUT_SQUARE_INCHES
            Case "F2", "FT2", "F^2", "FT^2"
                enumDes = DisplayUnitType.DUT_SQUARE_FEET
                '
            Case "GR"   ' Unidades de ANGULOS
                enumDes = DisplayUnitType.DUT_GRADS
            Case "RAD"
                enumDes = DisplayUnitType.DUT_RADIANS
        End Select
        resultado = UnitUtils.Convert(dblValor, enumOri, enumDes).ToString
        '
        Return resultado
    End Function
    '
    'Public Function U_DameDouble__DesdeCualquiera(queValor As String, uniOri As String, uniDes As String) As Object
    '    Dim resultadoTemp As String = U_DameString__DesdeCualquiera(queValor, uniOri, uniDes)
    '    Dim resultado As Object = queValor
    '    Try
    '        resultado = Convert.ToDouble(resultadoTemp, System.Globalization.CultureInfo.CurrentCulture)
    '    Catch ex As Exception
    '        '
    '    End Try
    '    Return resultado
    'End Function
    Public Function U_DameGrados_DesdeRadianes(radianes As Double, Optional formateadecimales As Boolean = True, Optional nDec As Integer = 2) As Double
        Dim resultado As Double = (radianes * 180) / Math.PI
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        If formateadecimales Then
            resultado = CDbl(FormatNumber(resultado, nDec, TriState.True, TriState.False, TriState.False))
        End If
        ''
        Return resultado
    End Function

    Public Function U_DameRadianes_DesdeGrados(degrees As Double, Optional formateadecimales As Boolean = True, Optional nDec As Integer = 2) As Double
        Dim resultado As Double = (degrees * Math.PI) / 180
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        If formateadecimales Then
            resultado = CDbl(FormatNumber(resultado, nDec, TriState.True, TriState.False, TriState.False))
        End If
        ''
        Return resultado
    End Function
    ''
    Public Function Unidades_DameMilimetros(queValor As String, unidades As String) As String
        Dim resultado As String = queValor.Split(" "c)(0).ToString
        '' Si es "" o "0" o no es numérico devolvemos "" y salimos
        If queValor.Trim = "" OrElse queValor.Trim = "0" OrElse IsNumeric(queValor) = False Then
            Return resultado
            Exit Function
        End If
        '
        If unidades.Trim.ToUpper = "MM" Then
            resultado = FormatNumber(queValor.Trim, 2, TriState.True, TriState.False, TriState.False)
        ElseIf unidades.Trim.ToUpper = "CM" Then
            resultado = Unidades_DameMilimetros(queValor.Trim, DisplayUnitType.DUT_CENTIMETERS)
        ElseIf unidades.Trim.ToUpper = "DM" Then
            resultado = Unidades_DameMilimetros(queValor.Trim, DisplayUnitType.DUT_DECIMETERS)
        ElseIf unidades.Trim.ToUpper = "M" Then
            resultado = Unidades_DameMilimetros(queValor.Trim, DisplayUnitType.DUT_METERS)
        ElseIf unidades.Trim.ToUpper = "IN" Then
            resultado = Unidades_DameMilimetros(queValor.Trim, DisplayUnitType.DUT_DECIMAL_INCHES)
        ElseIf unidades.Trim.ToUpper = "F" OrElse queValor.Trim.ToUpper = "FT" Then
            resultado = Unidades_DameMilimetros(queValor.Trim, DisplayUnitType.DUT_DECIMAL_FEET)
        End If
        ''
        Return resultado
    End Function
    Public Function Unidades_DameMilimetros(queValor As String, Optional decimales As Integer = 2) As String
        If queValor.Trim = "" OrElse queValor.Trim = "0" Then   ' Or IsNumeric(queValor.Trim) = False Then
            Return queValor
            Exit Function
        End If
        '
        queValor = queValor.Split(" "c)(0).ToString
        Dim resultado As String = ""
        ''
        'InicioAplicacionObligatorio()
        ''
        '' Poner el milimetros queValor. De unidades documento a milimetros. Con 2 decimales para mejorar conversión.
        resultado = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, decimales)), unidLDoc.DisplayUnits, DisplayUnitType.DUT_MILLIMETERS).ToString
        ''
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        resultado = FormatNumber(resultado, decimales, TriState.True, TriState.False, TriState.False)
        '' Formatear resultado sin decimales, ni separador de miles.
        'resultado = FormatNumber(resultado, 0, TriState.True, TriState.False, TriState.False)
        ''
        Return resultado
    End Function
    Public Function Unidades_DameMilimetros(queValor As String, queUnidOri As DisplayUnitType, Optional decimales As Integer = 2) As String
        If queValor.Trim = "" OrElse queValor.Trim = "0" Or IsNumeric(queValor.Trim) = False Then
            Return queValor
            Exit Function
        End If
        '
        queValor = queValor.Split(" "c)(0).ToString
        Dim resultado As String = ""
        ''
        'InicioAplicacionObligatorio()
        ''
        '' Poner el milimetros queValor. De unidades documento a milimetros. Con 2 decimales para mejorar conversión.
        If queUnidOri = DisplayUnitType.DUT_MILLIMETERS Then
            resultado = FormatNumber(queValor, decimales, TriState.True, TriState.False, TriState.False)
        Else
            resultado = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, decimales)), queUnidOri, DisplayUnitType.DUT_MILLIMETERS).ToString
        End If
        ''
        '' Formatear resultado sin decimales, ni separador de miles. Si era sin decimales
        '' O poner los mismos decimales que había.
        ''
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        resultado = FormatNumber(resultado, decimales, TriState.True, TriState.False, TriState.False)
        ''
        Return resultado
    End Function
    Public Function Unidades_DameMetrosCuadrados(queValor As String, Optional decimales As Integer = 2) As String
        If queValor.Trim = "" OrElse IsNumeric(queValor.Trim) = False Then
            Return queValor
            Exit Function
        End If
        ''
        queValor = queValor.Split(" "c)(0).ToString
        Dim resultado As String = ""
        ''
        'InicioAplicacionObligatorio()
        ''
        '' Poner el milimetros queValor. De unidades documento a milimetros. Con 2 decimales para mejorar conversión.
        resultado = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, decimales)), unidADoc.DisplayUnits, DisplayUnitType.DUT_SQUARE_METERS).ToString
        ''
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        resultado = FormatNumber(resultado, decimales, TriState.True, TriState.False, TriState.False)
        '' *****
        '' Formatear resultado sin decimales, ni separador de miles.
        'resultado = FormatNumber(resultado, 0, TriState.True, TriState.False, TriState.False)
        ''
        Return resultado
    End Function
    ''
    Public Function Unidades_DameMetrosCuadrados_DesdePiesCuadrados(metroscuadrados As Double, Optional decimales As Integer = 2) As String
        '' 1 Metro cuadrado = 10.7639 Pies cuadrados
        Dim multiplicar As Double = 10.7639
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return Math.Round(metroscuadrados * multiplicar, decimales).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Function UnidadesDocumento_DameFormatOptions(queUnitType As UnitType) As FormatOptions
        '' Llamarlo con Dim ut As UnitType = UnitUtils.GetValidUnitTypes(UnitType.UT_Length)
        'Dim ut As UnitType = UnitUtils.GetValidUnitTypes(queUnitType)
        Dim units As Units = evRevit.evAppUI.ActiveUIDocument.Document.GetUnits
        'Dim unitTypes As IList(Of UnitType) = UnitUtils.GetValidUnitTypes()
        ''
        Return units.GetFormatOptions(queUnitType)
    End Function
    ''
    Public Function UnidadesDocumento_DameUnidType(queUnitType As UnitType) As DisplayUnitType
        '' Llamarlo con Dim ut As UnitType = UnitUtils.GetValidUnitTypes(UnitType.UT_Length)
        'Dim ut As UnitType = UnitUtils.GetValidUnitTypes(queUnitType)
        Dim units As Units = evRevit.evAppUI.ActiveUIDocument.Document.GetUnits
        'Dim unitTypes As IList(Of UnitType) = UnitUtils.GetValidUnitTypes()
        ''
        Return units.GetFormatOptions(queUnitType).DisplayUnits
    End Function
    ''
    Public Function UnidadesDocumento_DameSimbolo(queUnitType As UnitType) As String
        '' Llamarlo con Dim ut As UnitType = UnitUtils.GetValidUnitTypes(UnitType.UT_Length)
        'Dim ut As UnitType = UnitUtils.GetValidUnitTypes(queUnitType)
        Dim resultado As String = ""
        Dim units As Units = evRevit.evAppUI.ActiveUIDocument.Document.GetUnits
        Dim oFOp As FormatOptions = units.GetFormatOptions(queUnitType)
        ''
        If oFOp.CanHaveUnitSymbol = False Then
            resultado = ""
        ElseIf oFOp.UnitSymbol = UnitSymbolType.UST_NONE Then
            resultado = ""
        Else
            resultado = LabelUtils.GetLabelFor(oFOp.UnitSymbol)
        End If
        ''
        Return resultado
    End Function

    ''
    Public Function Unidades_DamePiesCuadrados_DesdeMetrosCuadrados(piescuadrados As Double, Optional decimales As Integer = 2) As String
        '' 1 Pie cuadrado = 0.092903 Metros cuadrados
        Dim multiplicar As Double = 0.092903
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return Math.Round(piescuadrados * multiplicar, decimales).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Function Unidades_DamePulgadas_DesdeMilimetros(dblPulgadas As Double, Optional decimales As Integer = 2) As String
        '' 1 Pulgada = 25,4 milimetros
        Dim multiplicar As Double = 25.4
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return (dblPulgadas * multiplicar).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Function Unidades_DamePulgadas_DesdeMilimetros(queValor As String) As String

        If queValor.Trim = "" OrElse IsNumeric(queValor.Trim) = False Then
            Return queValor
            Exit Function
        End If
        ''
        queValor = queValor.Split(" "c)(0).ToString
        Dim resultado As String = ""
        ''
        'InicioAplicacionObligatorio()
        ''
        '' Poner el milimetros queValor. De unidades documento a milimetros. Con 2 decimales para mejorar conversión.
        resultado = UnitUtils.Convert(CDbl(FormatNumber(queValor.Trim, 2)), DisplayUnitType.DUT_DECIMAL_INCHES, DisplayUnitType.DUT_MILLIMETERS).ToString
        '' ***** SIEMPRE LO EVALUAREMOS CON 2 DECIMALES (Cambiar en todo el desarrollo)
        resultado = FormatNumber(resultado, 2, TriState.True, TriState.False, TriState.False)
        '' Formatear resultado sin decimales, ni separador de miles.
        'resultado = FormatNumber(resultado, 0, TriState.True, TriState.False, TriState.False)
        ''
        Return resultado
    End Function
    ''
    Public Function Unidades_DameMilimetros_DesdePulgadas(dblMilimetros As Double, Optional decimales As Integer = 0) As String
        '' 1 milimetro = 0,0393701 pulgadas
        Dim multiplicar As Double = 0.0393701
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return (dblMilimetros * multiplicar).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Function Unidades_DameLibras_DesdeKilos(dblLibras As Double, Optional decimales As Integer = 2) As String
        '' 1 Libra = 0,453592 Kilogramos
        Dim multiplicar As Double = 0.453592
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return (dblLibras * multiplicar).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Function Unidades_DameKilos_DesdeLibras(dblKilos As Double, Optional decimales As Integer = 3) As String
        '' 1 Kilo = 2,20462 Libras
        Dim multiplicar As Double = 2.20462
        Dim formato As String = ""
        If decimales = 0 Then
            formato = "#"
        Else
            formato = "#." & StrDup(decimales, "0")
        End If
        ''
        Return (dblKilos * multiplicar).ToString(formato, System.Globalization.CultureInfo.InvariantCulture)
    End Function
    ''
    Public Sub Unidades_ListaEscribeAllRevitUnits()
        Dim units As Units = evRevit.evAppUI.ActiveUIDocument.Document.GetUnits
        Dim unitTypes As IList(Of UnitType) = UnitUtils.GetValidUnitTypes()
        Dim info As String = "UNIDADES aplicadas en el documento actual" & Environment.NewLine
        ''
        For Each un As UnitType In unitTypes
            Dim fmtOpts As FormatOptions = units.GetFormatOptions(un)
            info &= Unidades_DameInformacion(un, fmtOpts, Environment.NewLine)
        Next
        ''
        'Dim queFicheroFin As String = IO.Path.Combine(My.Computer.FileSystem.SpecialDirectories.Temp, "RevitUnidades.txt")
        Dim queDirectorio As String = "C:\DESARROLLO"
        Dim queFichero As String = "RevitUnidades.txt"
        If IO.Directory.Exists(queDirectorio) = False Then
            Try
                queDirectorio = My.Computer.FileSystem.SpecialDirectories.Temp
            Catch ex As Exception
                Exit Sub
            End Try
        End If
        ''
        Dim ficherocompleto As String = IO.Path.Combine(queDirectorio, queFichero)
        Using sw As New IO.StreamWriter(ficherocompleto)
            Try
                sw.WriteLine(info)
            Catch ex As Exception
                'TaskDialog.Show("ATENCION", "ListaEscribeAllRevitUnits--> " & ex.Message & Environment.NewLine)
            End Try
            sw.Close()
        End Using
        ''
        If IO.File.Exists(ficherocompleto) Then Process.Start(ficherocompleto)
    End Sub
    '
    Public Function Unidades_DameInformacion(ut As UnitType, obj As FormatOptions, indent As String) As String
        Dim msg As String = String.Format(indent & "{0} ({1}): " & Environment.NewLine, LabelUtils.GetLabelFor(ut), ut)
        ''
        On Error Resume Next
        msg &= String.Format(indent & vbTab & "Accuracy: {0}" & Environment.NewLine, obj.Accuracy)
        msg &= String.Format(indent & vbTab & "Unit display: {0} ({1})" & Environment.NewLine, LabelUtils.GetLabelFor(obj.DisplayUnits), obj.DisplayUnits)
        msg &= String.Format(indent & vbTab & "Unit symbol: {0}" & Environment.NewLine,
                             IIf(obj.CanHaveUnitSymbol(), String.Format("{0} ({1})",
                                        IIf(obj.UnitSymbol = UnitSymbolType.UST_NONE, "", LabelUtils.GetLabelFor(obj.UnitSymbol)),
                                                    obj.UnitSymbol), "n/a"))
        msg &= String.Format(indent & vbTab & "Use default: {0}" & Environment.NewLine, obj.UseDefault)
        msg &= String.Format(indent & vbTab & "Use grouping: {0}" & Environment.NewLine, obj.UseDigitGrouping)
        msg &= String.Format(indent & vbTab & "Use digit grouping: {0}" & Environment.NewLine, obj.UseDigitGrouping)
        msg &= String.Format(indent & vbTab & "Use plus prefix: {0}" & Environment.NewLine,
                             IIf(obj.CanUsePlusPrefix(), obj.UsePlusPrefix.ToString(), "n/a"))
        msg &= String.Format(indent & vbTab & "Suppress spaces: {0}" & Environment.NewLine,
                             IIf(obj.CanSuppressSpaces(), obj.SuppressSpaces.ToString(), "n/a"))
        msg &= String.Format(indent & vbTab & "Suppress leading zeros: {0}" & Environment.NewLine,
                             IIf(obj.CanSuppressLeadingZeros(), obj.SuppressLeadingZeros.ToString(), "n/a"))
        msg &= String.Format(indent & vbTab & "Suppress trailing zeros: {0}" & Environment.NewLine,
                             IIf(obj.CanSuppressTrailingZeros, obj.SuppressTrailingZeros.ToString(), "n/a"))
        ''
        Return msg
    End Function

    '
    Public Function FtToMm(queft As Double, Optional enDouble As Boolean = True) As Object
        ' used to convert stored units from ft to mm for display
        Dim resultado As Object = queft * 304.8 ' * 0.00328
        If enDouble = False Then
            resultado = resultado.ToString & " mm"
        End If
        Return resultado
    End Function
    '
    Public Function MmToFt(queMm As Double, Optional enDouble As Boolean = True) As Object
        ' used to convert stored units from ft to mm for display
        Dim resultado As Object = queMm / 304.8
        If enDouble = False Then
            resultado = resultado.ToString & " ft"
        End If
        Return resultado
    End Function

    '' Ponemos milímetros a las longitudes, que son doubles y no llevan unidades.
    'If oPar.StorageType = StorageType.Double AndAlso ParametroLlevaUnidades(valor) = False Then
    'Select Case oPar.DisplayUnitType
    'Case DisplayUnitType.DUT_MILLIMETERS
    '                        valor &= " mm"
    '                    Case DisplayUnitType.DUT_CENTIMETERS
    '                        valor &= " cm"
    '                    Case DisplayUnitType.DUT_DECIMETERS
    '                        valor &= " dm"
    '                    Case DisplayUnitType.DUT_METERS
    '                        valor &= " m"
    '                    Case DisplayUnitType.DUT_SQUARE_MILLIMETERS
    '                        valor &= " mm2"
    '                    Case DisplayUnitType.DUT_SQUARE_CENTIMETERS
    '                        valor &= " cm2"
    '                    Case DisplayUnitType.DUT_SQUARE_METERS
    '                        valor &= " m2"
    '                    Case DisplayUnitType.DUT_DECIMAL_INCHES
    '                        valor &= " in"
    '                    Case DisplayUnitType.DUT_DECIMAL_FEET
    '                        valor &= " ft"
    '                    Case DisplayUnitType.DUT_PERCENTAGE
    '                        valor &= " %"
    '                End Select
    'End If
    '
    ''
    ''MsgBox Fraction("3/4") & vbCrLf & Fraction("2 7/8") & Fraction("2 - 7/8")
    Public Function Fraction(ByVal TextIn As String) As Double
        Dim nResult As Double
        Dim x As Integer
        x = InStr(1, TextIn, " ")
        If x > 1 Then
            If IsNumeric(Left$(TextIn, x - 1)) Then
                nResult = CDbl(Left$(TextIn, x - 1))
            End If
            TextIn = Mid$(TextIn, x)
        End If
        x = InStr(1, TextIn, "/")
        If x > 1 Then
            If IsNumeric(Left$(TextIn, x - 1)) Then
                If IsNumeric(Mid$(TextIn, x + 1)) Then
                    nResult = nResult + CDbl(Left$(TextIn, x - 1)) / CDbl(Mid$(TextIn, x +
              1))
                End If
            End If
        End If
        Fraction = nResult
    End Function
End Module