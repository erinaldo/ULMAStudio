Imports System.IO
Imports System.IO.Compression
Imports System.IO.Compression.ZipFile
Module modComprimeDescomprime
#Region "GZIPSTREAM"
    'Private directoryPath As String = "c:\temp"
    'Public Sub Main()
    '    Dim directorySelected As New DirectoryInfo(directoryPath)
    '    Compress(directorySelected)

    '    For Each fileToDecompress As FileInfo In directorySelected.GetFiles("*.gz")
    '        Decompress(fileToDecompress)
    '    Next
    'End Sub
    Public Sub CompressGZipStream_CadaFichero(directorySelected As DirectoryInfo)
        For Each fileToCompress As FileInfo In directorySelected.GetFiles("*.*", SearchOption.AllDirectories)
            Using originalFileStream As FileStream = fileToCompress.OpenRead()
                If (File.GetAttributes(fileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden And fileToCompress.Extension <> ".zip" Then
                    Using compressedFileStream As FileStream = File.Create(fileToCompress.FullName & ".zip") ' & ".gz")
                        Using compressionStream As New GZipStream(compressedFileStream, CompressionMode.Compress)
                            originalFileStream.CopyTo(compressionStream)
                        End Using
                    End Using
                    Dim info As New FileInfo(fileToCompress.FullName & ".zip") '& ".gz")
                    Console.WriteLine("Compressed {0} from {1} to {2} bytes.", fileToCompress.Name,
                                      fileToCompress.Length.ToString(), info.Length.ToString())

                End If
            End Using
        Next
    End Sub
    Public Sub CompressGZipStream_EnUnSoloFichero(directorySelected As DirectoryInfo)
        Dim fiFin As String = directorySelected.FullName & ".zip"
        Using compressedFileStream As FileStream = File.Create(fiFin)
            Using compressionStream As New GZipStream(compressedFileStream, CompressionMode.Compress)
                For Each fileToCompress As FileInfo In directorySelected.GetFiles("*.*", SearchOption.AllDirectories)
                    Using originalFileStream As FileStream = fileToCompress.OpenRead()
                        If (File.GetAttributes(fileToCompress.FullName) And FileAttributes.Hidden) <> FileAttributes.Hidden Then
                            ' & ".gz")
                            originalFileStream.CopyTo(compressionStream)
                            Dim info As New FileInfo(fileToCompress.FullName & ".zip") '& ".gz")
                            'Console.WriteLine("Compressed {0} from {1} to {2} bytes.", fileToCompress.Name,
                            'fileToCompress.Length.ToString(), info.Length.ToString())
                            Console.WriteLine("Compressed {0} from {1}.", fileToCompress.Name,
                                                  fileToCompress.Length.ToString())

                        End If
                    End Using
                Next
            End Using
        End Using
    End Sub


    Public Sub DecompressGZipStream(ByVal fileToDecompress As FileInfo)
        Using originalFileStream As FileStream = fileToDecompress.OpenRead()
            Dim currentFileName As String = fileToDecompress.FullName
            Dim newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length)

            Using decompressedFileStream As FileStream = File.Create(newFileName)
                Using decompressionStream As GZipStream = New GZipStream(originalFileStream, CompressionMode.Decompress)
                    decompressionStream.CopyTo(decompressedFileStream)
                    Console.WriteLine("Decompressed: {0}", fileToDecompress.Name)
                End Using
            End Using
        End Using
    End Sub
#End Region

#Region "ZIPFILE"
    Public Sub CompressZipFile_carpeta(carpetaOrigen As String, ficheroFin As String)
        ZipFile.CreateFromDirectory(carpetaOrigen, ficheroFin)
    End Sub
    Public Sub DecompressZipFile_carpeta(ficheroOrigen As String, carpetaFin As String)
        'Dim startPath As String = "c:\example\start"
        'Dim zipPath As String = "c:\example\result.zip"
        'Dim extractPath As String = "c:\example\extract"

        ZipFile.ExtractToDirectory(ficheroOrigen, carpetaFin)
    End Sub
#End Region
End Module
