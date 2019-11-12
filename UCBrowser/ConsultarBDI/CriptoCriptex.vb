Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Security.Cryptography


Public Class CriptoCriptex

  Private gCriptoSistema As New RijndaelManaged

  '/**
  '* Los caracteres "extraños" suelen dar problemas en sistemas 
  '* con juegos de caracteres no-Western (paises del este, paises asiáticos,...)
  '**/
  Private Sub ComprobarQueSoloHayCaracteresASCIIpuros(texto As String)
    Dim patronDeBusqueda As Regex
    patronDeBusqueda = New Regex("[^ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789]")
    Dim trozoEncontrado As Match
    trozoEncontrado = patronDeBusqueda.Match(texto)
    If trozoEncontrado.Success Then
      Throw New ArgumentOutOfRangeException("strClave" & _
                                            Environment.NewLine & Environment.NewLine & _
                                            "Not valid [" & trozoEncontrado.Value & "] on the key." & _
                                            Environment.NewLine & Environment.NewLine & _
                                            "Valid characters are: [ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789]") 
    End If
  End Sub


  '/**
  '* constructor del criptosistema
  '**/
  Public Sub New(ByVal strClave As String) 
    ComprobarQueSoloHayCaracteresASCIIpuros(strClave)

    Dim strClaveRevuelta As String = ""
    If strClave.Length > 8 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(7)
    If strClave.Length > 3 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(2)
    If strClave.Length > 16 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(15)
    If strClave.Length > 2 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(1)
    If strClave.Length > 4 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(3)
    If strClave.Length > 7 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(6)
    If strClave.Length > 9 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(8)
    If strClave.Length > 5 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(4)
    If strClave.Length > 10 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(9)
    If strClave.Length > 1 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(0)
    If strClave.Length > 13 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(12)
    If strClave.Length > 17 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(16)
    If strClave.Length > 20 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(19)
    If strClave.Length > 39 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(38)
    If strClave.Length > 12 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(11)
    If strClave.Length > 6 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(5)
    If strClave.Length > 11 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(10)
    If strClave.Length > 18 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(17)
    If strClave.Length > 19 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(18)
    If strClave.Length > 24 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(23)
    If strClave.Length > 36 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(35)
    If strClave.Length > 14 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(13)
    If strClave.Length > 35 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(34)
    If strClave.Length > 32 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(31)
    If strClave.Length > 15 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(14)
    If strClave.Length > 21 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(20)
    If strClave.Length > 26 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(25)
    If strClave.Length > 41 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(40)
    If strClave.Length > 25 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(24)
    If strClave.Length > 22 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(21)
    If strClave.Length > 27 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(26)
    If strClave.Length > 31 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(30)
    If strClave.Length > 29 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(28)
    If strClave.Length > 28 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(27)
    If strClave.Length > 40 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(39)
    If strClave.Length > 33 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(32)
    If strClave.Length > 38 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(37)
    If strClave.Length > 23 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(22)
    If strClave.Length > 42 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(41)
    If strClave.Length > 34 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(33)
    If strClave.Length > 37 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(36)
    If strClave.Length > 30 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(29)
    If strClave.Length > 43 Then strClaveRevuelta = strClaveRevuelta & strClave.Chars(42)
    If strClave.Length > 44 Then strClaveRevuelta = strClaveRevuelta & strClave.Substring(43)

    Dim binClave() As Byte
    ReDim binClave(gCriptoSistema.Key.Length - 1)
    Dim i, j As Integer
    j = 0
    For i = 0 To binClave.Length - 1
      binClave(i) = CByte(Asc(strClaveRevuelta.Chars(j)))
      If j < strClaveRevuelta.Length - 1 Then
        j = j + 1
      Else
        j = 0
      End If
    Next
    gCriptoSistema.Key = binClave

    Dim binVectorIni() As Byte
    ReDim binVectorIni(gCriptoSistema.IV.Length - 1)
    j = 0
    For i = 0 To binVectorIni.Length - 1
      binVectorIni(i) = CByte(Asc(strClaveRevuelta.Chars(j)))
      If j < strClaveRevuelta.Length - 1 Then
        j = j + 1
      Else
        j = 0
      End If
    Next
    gCriptoSistema.IV = binVectorIni

  End Sub


  '/**
  '* Cifra una cadena de texto, y nos devuelve una cadena de números hexadecimales
  '**/
  Public Function Cifrar(ByVal strTextoClaro As String) As String
    If strTextoClaro = "" Then Return ""

    '// contenedor -binario- del texto claro:
    Dim conversorTexto As New ASCIIEncoding
    Dim binTextoClaro As Byte() = conversorTexto.GetBytes(strTextoClaro)

    '// cifrador:
    Dim TransformacionCifradora As ICryptoTransform
    TransformacionCifradora = gCriptoSistema.CreateEncryptor
    Dim FlujoCifrador As CryptoStream
    Dim FlujoMem As New MemoryStream
    FlujoCifrador = New CryptoStream(FlujoMem, TransformacionCifradora, CryptoStreamMode.Write)
    FlujoCifrador.Write(binTextoClaro, 0, binTextoClaro.Length)
    FlujoCifrador.FlushFinalBlock()

    '// contenedor del texto cifrado:
    Dim strTextoCifrado As String 
    strTextoCifrado = ""
    Dim binTextoCifrado As Byte()= FlujoMem.ToArray
    Dim binChar As Byte
    For Each binChar In binTextoCifrado
      strTextoCifrado = strTextoCifrado & binChar.ToString("X2")
    Next

    Return strTextoCifrado
  End Function

  '/**
  '* Dada una cadena de numeros hexadecimales, la descifra 
  '* y nos devuelve la cadena de texto correspondiente
  '**/
  Public Function Descifrar(ByVal strTextoCifrado As String) As String
    If strTextoCifrado = "" Then Return ""

    '// contenedor -binario- del texto cifrado:
    Dim binTextoCifrado() As Byte = New Byte(CInt(Math.Ceiling(strTextoCifrado.Length / 2) - 1)) {}
    Dim chrChar(2) As Char
    Dim i As Integer = 0
    Dim j As Integer = 0
    While j + 2 <= strTextoCifrado.Length
      binTextoCifrado(i) = Byte.Parse(strTextoCifrado.Substring(j, 2), _
                                      System.Globalization.NumberStyles.AllowHexSpecifier)
      i = i + 1
      j = j + 2
    End While

    '// descifrador:
    Dim TransformacionDescifradora As ICryptoTransform 
    TransformacionDescifradora = gCriptoSistema.CreateDecryptor
    Dim FlujoMem As MemoryStream
    FlujoMem = New MemoryStream(binTextoCifrado)
    Dim FlujoDescifrador As CryptoStream
    FlujoDescifrador = New CryptoStream(FlujoMem, TransformacionDescifradora, CryptoStreamMode.Read)

    '// contenedor del texto claro:
    Dim binTextoClaro As Byte() = New Byte(CInt(Math.Ceiling(strTextoCifrado.Length / 2) - 1)) {}
    Try
      FlujoDescifrador.Read(binTextoClaro, 0, binTextoClaro.Length)
        Catch ex As Exception
            MsgBox(Prompt:="Problema grave al desencriptar la cadena:" _
                           & Environment.NewLine & strTextoCifrado & Environment.NewLine _
                           & "No se ha podido obtener el texto claro correspondiente." _
                           & Environment.NewLine & ex.Message _
                           & Environment.NewLine & Environment.NewLine & ex.StackTrace, _
                   Buttons:=MsgBoxStyle.Critical, _
                   Title:="ERROR")
    End Try

    Dim conversorTexto As New ASCIIEncoding
    Dim strTextoClaro As String = conversorTexto.GetString(binTextoClaro)
    Dim strTextoClaroLimpio As String = ""
    i = 0
    While i < strTextoClaro.Length - 1 And Asc(strTextoClaro.Chars(i)) > 0
      strTextoClaroLimpio = strTextoClaroLimpio & strTextoClaro.Chars(i)
      i = i + 1
    End While
    Return strTextoClaroLimpio

  End Function



End Class
