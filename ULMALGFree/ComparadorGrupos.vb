Public Class ComparadorGrupos
    Implements IComparer(Of Group)

    Public Sub New()
    End Sub


    Public Function Compare(x As Group, y As Group) As Integer Implements IComparer(Of Group).Compare
        Return x.gDescription.CompareTo(y.gDescription)
    End Function
End Class

