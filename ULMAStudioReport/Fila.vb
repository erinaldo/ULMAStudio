Imports System.Collections.Generic
Imports ULMAStudioReport

Public Class fila
    Implements IComparable(Of fila)
    Public Property number As Integer
    Public Property ImageByte As Byte()
    Public Property Name As String
    Public Property Code As String
    Public Property Weight As Double
    Public Property Quantity As Integer
    Public Property EsUlma As Boolean

    Public Sub New(n As Integer, oImage As Image, oName As String, oCode As String, oWeight As Double, oQuantity As Integer, oEsUlma As Boolean)
        Me.number = n
        Me.ImageByte = GetBytes(oImage)
        Me.Name = oName
        Me.Code = oCode
        Me.Weight = oWeight
        Me.Quantity = oQuantity
        Me.EsUlma = oEsUlma
    End Sub

    Public Function GetBytes(imageIn As Image) As Byte()

        'Usamos la clase MemoryStream para contener los bytes que compone la imagen
        Dim ms As IO.MemoryStream = New IO.MemoryStream()
        imageIn.Save(ms, Imaging.ImageFormat.Png)
        '/Retornamos el arreglo de bytes
        Return ms.ToArray()
    End Function

    Public Function CompareTo(other As fila) As Integer Implements IComparable(Of fila).CompareTo
        Return Me.number.CompareTo(other.number)
    End Function
End Class
