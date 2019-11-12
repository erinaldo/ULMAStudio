Module modVar
    Public _appPath As String = System.Reflection.Assembly.GetExecutingAssembly.Location
    Public _appFolder As String = IO.Path.GetDirectoryName(_appPath)
    Public _appOptions As String = IO.Path.ChangeExtension(_appPath, "txt")
End Module
