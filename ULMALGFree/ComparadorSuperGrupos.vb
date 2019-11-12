Public Class ComparadorSuperGrupos
    Implements IComparer(Of SuperGroup)

    Public Sub New()
    End Sub


    Public Function Compare(x As SuperGroup, y As SuperGroup) As Integer Implements IComparer(Of SuperGroup).Compare
        Return x.sgCode.CompareTo(y.sgCode)
    End Function
End Class
