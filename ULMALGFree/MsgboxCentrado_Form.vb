
Imports System.Text
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class MsgboxCentrado_Form
    Implements IDisposable

    Private mTries As Integer = 0
    Private mOwner As Form

    Public Sub New(ByVal owner As Form)
        mOwner = owner
        owner.BeginInvoke(New MethodInvoker(AddressOf findDialog))
    End Sub

    Private Sub findDialog()
        ' Enumerate windows to find the message box
        If mTries < 0 Then
            Return
        End If
        Dim callback As New EnumThreadWndProc(AddressOf checkWindow)
        If EnumThreadWindows(GetCurrentThreadId(), callback, IntPtr.Zero) Then
            If System.Threading.Interlocked.Increment(mTries) < 10 Then
                mOwner.BeginInvoke(New MethodInvoker(AddressOf findDialog))
            End If
        End If
    End Sub
    Private Function checkWindow(ByVal hWnd As IntPtr, ByVal lp As IntPtr) As Boolean
        ' Checks if <hWnd> is a dialog
        Dim sb As New StringBuilder(260)
        GetClassName(hWnd, sb, sb.Capacity)
        If sb.ToString() <> "#32770" Then
            Return True
        End If
        ' Got it
        Dim frmRect As New Rectangle(mOwner.Location, mOwner.Size)
        Dim dlgRect As RECT
        GetWindowRect(hWnd, dlgRect)
        MoveWindow(hWnd, frmRect.Left + (frmRect.Width - dlgRect.Right + dlgRect.Left) \ 2, frmRect.Top + (frmRect.Height - dlgRect.Bottom + dlgRect.Top) \ 2, dlgRect.Right - dlgRect.Left, dlgRect.Bottom - dlgRect.Top, True)
        Return False
    End Function
    Public Sub Dispose() Implements IDisposable.Dispose
        GC.SuppressFinalize(Me)
        mTries = -1
    End Sub

    ' P/Invoke declarations
    Private Delegate Function EnumThreadWndProc(ByVal hWnd As IntPtr, ByVal lp As IntPtr) As Boolean

    <DllImport("user32.dll")>
    Private Shared Function EnumThreadWindows(ByVal tid As Integer, ByVal callback As EnumThreadWndProc, ByVal lp As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll")>
    Private Shared Function GetCurrentThreadId() As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetClassName(ByVal hWnd As IntPtr, ByVal buffer As StringBuilder, ByVal buflen As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Private Shared Function GetWindowRect(ByVal hWnd As IntPtr, ByRef rc As RECT) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function MoveWindow(ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal h As Integer, ByVal repaint As Boolean) As Boolean
    End Function

    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
End Class
