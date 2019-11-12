
Public Class clsFamCode
    Public W_MARKET As String = ""
    Public IMARKET As String = ""
    Public FCODE As String = ""
    Public ICODE As String = ""
    Public ILENGTH As String = ""
    Public IWIDTH As String = ""
    Public IHEIGHT As String = ""
    Public IGENERIC As String = ""
    Public RFCODE As String = ""
    Public fileName As String = ""
    Public IDESCRIPTIONES As String = ""
    Public IDESCRIPTIONEN As String = ""
    Public IUnid As String = ""
    Public clave As String = ""
    Public claveGEN As String = ""
    Public cabeceracorrecta As Boolean = True
    ''
    ''
    Public Sub New()
        '' No hacemos nada. Tras rellenar cada clsFamCode ejecutar PonMedidas para asegurar
        '' quetendrán 2 decimales y estará en las unidades correctas
    End Sub
    ''
    Public Sub PonMedidas()
        '' Las unidades de largo, ancho y alto con punto decimal y 2 decimales.
        Me.ILENGTH = Unidades_DameMilimetros(Me.ILENGTH.Trim, IUnid)
        Me.IWIDTH = Unidades_DameMilimetros(Me.IWIDTH.Trim, IUnid)
        Me.IHEIGHT = Unidades_DameMilimetros(Me.IHEIGHT.Trim, IUnid)
        '' Las claves de búsqueda de esta familia (Para dinamicos y genéricos)
        'If IGENERIC = "0" Then
        '' Es DINÁMICO. Tendrá sus medidas  (Tiene ILENGTH o IWIDTH)
        'Me.clave = Me.IMARKET & Me.FCODE.Trim & Me.ILENGTH & Me.IWIDTH & Me.IHEIGHT
        'ElseIf IGENERIC = "1" Then
        '' Generico LINEAL (Tiene ILENGTH o IWIDTH)
        'ElseIf IGENERIC = "2" Then
        '' Generico AREA
        'End If
        ''
        Me.clave = Me.IMARKET & Me.FCODE.Trim & Me.ILENGTH & Me.IWIDTH & Me.IHEIGHT
    End Sub
End Class
