Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi
Imports Windows.Win32.System

Public Class Form1

    ' ===== CoreAudio (system mic) =====
    Private mmDevice As MMDevice
    Private isTalking As Boolean = False
    Private _iconDefault As Icon
    Private _iconTalk As Icon
    Private _iconMute As Icon

    ' ===== Tray & minimize =====
    Private startMinimized As Boolean = True

    ' ===== Low-level keyboard hook =====
    Private hookId As IntPtr = IntPtr.Zero
    Private hookProc As LowLevelKeyboardProcDelegate

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101

    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private ReadOnly PttKeys As Keys() = {Keys.LControlKey, Keys.RControlKey}

    <StructLayout(LayoutKind.Sequential)>
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInteger
        Public scanCode As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    Private Delegate Function LowLevelKeyboardProcDelegate(
        nCode As Integer,
        wParam As IntPtr,
        lParam As IntPtr
    ) As IntPtr

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(
        idHook As Integer,
        lpfn As LowLevelKeyboardProcDelegate,
        hMod As IntPtr,
        dwThreadId As UInteger
    ) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function CallNextHookEx(
        hhk As IntPtr,
        nCode As Integer,
        wParam As IntPtr,
        lParam As IntPtr
    ) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    ' ===== LOW-LEVEL KEYBOARD HOOK =====
    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private Function LowLevelKeyboardProc(
        nCode As Integer,
        wParam As IntPtr,
        lParam As IntPtr
    ) As IntPtr

        If nCode >= 0 Then
            Dim data = Marshal.PtrToStructure(Of KBDLLHOOKSTRUCT)(lParam)
            Dim key = CType(data.vkCode, Keys)

            If PttKeys.Contains(key) Then
                If wParam = CType(WM_KEYDOWN, IntPtr) Then
                    PressPTT()
                ElseIf wParam = CType(WM_KEYUP, IntPtr) Then
                    ReleasePTT()
                End If
            End If
        End If

        Return CallNextHookEx(hookId, nCode, wParam, lParam)
    End Function

    ' ===== PTT HELPERS =====
    Private Sub PressPTT()
        If isTalking Then Exit Sub

        SetMicMute(False)
        isTalking = True

        Try
            Label1.ForeColor = Color.Green
            Label1.Text = "Talking… (PTT active)"
            PictureBox1.Image = My.Resources.Mic_Green

            NotifyIcon1.Visible = True
            NotifyIcon1.Icon = Icon.FromHandle(CType(My.Resources.Talk, Bitmap).GetHicon())
            NotifyIcon1.Text = "Microphone: ACTIVE (PTT held)"

            Me.Icon = _iconTalk
        Catch
        End Try
    End Sub

    Private Sub ReleasePTT()
        If Not isTalking Then Exit Sub

        SetMicMute(True)
        isTalking = False

        Try
            Label1.ForeColor = Color.Red
            Label1.Text = "Muted (PTT inactive)"
            PictureBox1.Image = My.Resources.Mic_Red

            NotifyIcon1.Visible = True
            NotifyIcon1.Icon = Icon.FromHandle(CType(My.Resources.Mute, Bitmap).GetHicon())
            NotifyIcon1.Text = "Microphone: Muted (PTT inactive)"

            Me.Icon = _iconMute
        Catch
        End Try
    End Sub

    ' ===== FORM LOAD =====
    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label2.Text = "Version: " & My.Application.Info.Version.ToString()

        _iconDefault = Icon.FromHandle(CType(My.Resources.Mute, Bitmap).GetHicon())
        _iconTalk = Icon.FromHandle(CType(My.Resources.Talk, Bitmap).GetHicon())
        _iconMute = Icon.FromHandle(CType(My.Resources.Mute, Bitmap).GetHicon())

        ' Pick audio device
        Dim enumerator As New MMDeviceEnumerator()
        Try
            mmDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications)
        Catch
            mmDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console)
        End Try

        SetMicMute(True)

        ' Keyboard hook
        hookProc = AddressOf LowLevelKeyboardProc
        hookId = SetWindowsHookEx(
            WH_KEYBOARD_LL,
            hookProc,
            GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName),
            0
        )

        Label1.ForeColor = Color.Red
        Label1.Text = "Muted (PTT inactive)"
        PictureBox1.Image = My.Resources.Mic_Red
        Me.Icon = _iconMute

        ' ===== Start minimized to tray =====
        If startMinimized Then
            Me.WindowState = FormWindowState.Minimized
            Me.ShowInTaskbar = False

            ' Assign default icon first
            NotifyIcon1.Icon = _iconMute
            NotifyIcon1.Text = "Microphone: Muted (PTT inactive)"
            NotifyIcon1.Visible = True
        End If
    End Sub

    ' ===== HANDLE MINIMIZE TO TRAY =====
    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.ShowInTaskbar = False
            NotifyIcon1.Visible = True
        End If
    End Sub

    ' ===== RESTORE WINDOW ON TRAY DOUBLE-CLICK =====
    Private Sub NotifyIcon1_DoubleClick(sender As Object, e As EventArgs) Handles NotifyIcon1.DoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
        Me.Activate()
    End Sub

    ' ===== FORM CLOSING =====
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If hookId <> IntPtr.Zero Then UnhookWindowsHookEx(hookId)
        NotifyIcon1.Visible = False
        SetMicMute(False)
    End Sub

    ' ===== MIC CONTROL =====
    Private Sub SetMicMute(mute As Boolean)
        Try
            If mmDevice IsNot Nothing Then
                mmDevice.AudioEndpointVolume.Mute = mute
            End If
        Catch
        End Try
    End Sub

End Class
