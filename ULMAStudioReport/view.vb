Imports System.Collections.Generic
Public Class view
    Public Property view As String
    Public Property filas As New List(Of Fila)

    Public Sub New(oView As String)
        Me.view = oView
    End Sub

    Public Sub PonFila(oFila As Fila)
        Me.filas.Add(oFila)
    End Sub
End Class
