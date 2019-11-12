Public Class clsArticulos
    Public active As Boolean = True
    '
    Public articleCode As String = ""
    Public articleCodeL As String = ""
    Public colDescritions As Hashtable         '' Key=es, Value=descripción en el idio de Key.
    Public material As String = ""
    Public weight As String = ""
    Public weightUnit As String = ""
    Public weightL As String = ""
    Public weightLUnit As String = ""
    Public formArea As String = ""
    Public formAreaUnit As String = ""
    Public stockUnit As String = ""
    Public genericArticleUnit As String = ""
    ' Con las unidades, como texto, después del valor
    Public weightAll As String = ""
    Public weightLAll As String = ""
    Public formAreaAll As String = ""
    Public weightEND As String = ""
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
    '
    Public Sub PesoArea_ConUnidades()
        ' Rellenar weightAll, weightLAll y formAreaAll con las unidades correspondientes.
        If Me.weight <> "" AndAlso Me.weightUnit <> "" Then
            Me.weightAll = Me.weight & " " & Me.weightUnit
        ElseIf Me.weight <> "" AndAlso Me.weightUnit = "" Then
            Me.weightAll = Me.weight & " kg"      ' kg, por defecto, si no tiene weightUnit
        Else
            Me.weightAll = Me.weight                        ' Vacio, si no hay ni valor ni unidades
        End If
        '
        If Me.weightL <> "" AndAlso Me.weightLUnit <> "" Then
            Me.weightLAll = Me.weightL & " " & Me.weightLUnit
        ElseIf Me.weightL <> "" AndAlso Me.weightLUnit = "" Then
            Me.weightLAll = Me.weightL & " kg"    ' kg, por defecto, si no tiene weightLUnit
        Else
            Me.weightLAll = Me.weightLAll                       ' Vacio, si no hay ni valor ni unidades
        End If
        '
        If Me.formArea <> "" AndAlso Me.formAreaUnit <> "" Then
            Me.formAreaAll = Me.formArea & " " & Me.formAreaUnit
        ElseIf Me.formArea <> "" AndAlso Me.formAreaUnit = "" Then
            Me.formAreaAll = Me.formArea & " m2"  ' m2, por defecto, si no tiene formAreaUnit
        Else
            Me.formAreaAll = ""                      ' Vacio, si no hay ni valor ni unidades
        End If
    End Sub
    '
    Public Function weightEnd_Dame(Optional conUnits As Boolean = False) As String
        PesoArea_ConUnidades()
        Dim resultado As String = ""
        If conUnits = False Then
            resultado = Me.weight
            'If weightL <> "" AndAlso IsNumeric(weightL) AndAlso weightL <> weight Then
            '    resultado = weightL
            'End If
        ElseIf conUnits = True Then
            resultado = Me.weightAll
            'If weightL <> "" AndAlso IsNumeric(weightL) AndAlso weightL <> weight Then
            '    resultado = weightLAll
            'End If
        End If
        Return resultado
    End Function
End Class
