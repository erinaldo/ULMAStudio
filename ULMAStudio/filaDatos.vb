Imports System.Collections.Generic
Public Class filaDatos
    Public Property ImageByte As Byte()
    Public Property ImagePath As String
    Public Property Name As String
    Public Property Code As String
    Public Property Weight As Double
    Public Property Quantity As Integer
    Public Property esulma As Boolean = False
    Public Property categoria As String = ""

    Public Sub New(pathImage As String, oName As String, oCode As String, oWeight As Double, oQuantity As Integer,
                   ulma As Boolean,
                   oCategoria As String)
        Me.ImagePath = pathImage
        If IO.File.Exists(pathImage) Then
            Me.ImageByte = GetBytes(System.Drawing.Image.FromFile(pathImage))
        Else
            Me.ImageByte = Nothing
        End If
        Me.ImageByte = GetBytes(System.Drawing.Image.FromFile(pathImage))
        Me.Name = oName
        Me.Code = oCode
        Me.Weight = oWeight
        Me.Quantity = oQuantity
        Me.esulma = ulma
        Me.categoria = oCategoria
    End Sub

    Public Function GetBytes(imageIn As System.Drawing.Image) As Byte()

        'Usamos la clase MemoryStream para contener los bytes que compone la imagen
        Dim ms As IO.MemoryStream = New IO.MemoryStream()
        imageIn.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
        '/Retornamos el arreglo de bytes
        Return ms.ToArray()
    End Function
End Class
