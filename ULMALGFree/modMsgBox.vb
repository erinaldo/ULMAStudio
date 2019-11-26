Imports System.Windows.Forms

Module modMsgBox
    Public Function MsgBox(Prompt As String, Optional Buttons As MsgBoxStyle = MsgBoxStyle.OkOnly, Optional Titulo As String = "") As MsgBoxResult

        Dim hInst As Long
        Dim Thread As Long

        'Set up the CBT hook
        hInst = clsAPI.GetWindowLongPtr(Me.hwnd, clsAPI.GWL.GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProc2, hInst,
                                 Thread)

        'Display the message box
        MsgBox = VBA.MsgBox Prompt, Buttons, Title, HelpFile, Context
End Function
End Module
