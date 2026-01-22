Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports NAudio.CoreAudioApi
Imports Windows.Win32.System

Public Class Form1
    ' ===== CoreAudio (system mic) =====
    Private mmDevice As MMDevice
    Private isTalking As Boolean = False

    ' ===== Low-level keyboard hook =====
    Private hookId As IntPtr = IntPtr.Zero
    Private hookProc As LowLevelKeyboardProcDelegate

    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const WM_KEYUP As Integer = &H101

    ' Choose your push-to-talk key (you can change this)
    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private ReadOnly PttKeys As Keys() = {Keys.LControlKey, Keys.RControlKey}

    ' KBDLLHOOKSTRUCT for WH_KEYBOARD_LL
    <StructLayout(LayoutKind.Sequential)>
    Private Structure KBDLLHOOKSTRUCT
        Public vkCode As UInteger
        Public scanCode As UInteger
        Public flags As UInteger
        Public time As UInteger
        Public dwExtraInfo As IntPtr
    End Structure

    Private Delegate Function LowLevelKeyboardProcDelegate(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelKeyboardProcDelegate, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private Function LowLevelKeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 Then
            Dim data = Marshal.PtrToStructure(Of KBDLLHOOKSTRUCT)(lParam)
            Dim key = CType(data.vkCode, Keys)

            If PttKeys.Contains(key) Then
                If wParam = CType(WM_KEYDOWN, IntPtr) Then
                    ' Avoid repeat from key auto-repeat
                    If Not isTalking Then
                        SetMicMute(False)
                        isTalking = True
                        Try
                            Label1.ForeColor = System.Drawing.Color.Green
                            Me.BeginInvoke(Sub() Label1.Text = "Talking… (hold Ctrl)")
                            PictureBox1.Image = My.Resources.Mic_Green

                            NotifyIcon1.Visible = False
                            NotifyIcon1.Icon = Icon.FromHandle(CType(My.Resources.Talk, Bitmap).GetHicon())
                            NotifyIcon1.Visible = True

                            NotifyIcon1.Text = "Microphone: ACTIVE (PTT held)"

                        Catch
                        End Try
                    End If

                ElseIf wParam = CType(WM_KEYUP, IntPtr) Then
                    ' If you’re using both Ctrl keys, you might want to track counts.
                    If isTalking Then
                        SetMicMute(True)
                        isTalking = False
                        Try
                            Label1.ForeColor = System.Drawing.Color.Red
                            Me.BeginInvoke(Sub() Label1.Text = "Muted (hold Ctrl to talk)")
                            PictureBox1.Image = My.Resources.Mic_Red

                            NotifyIcon1.Visible = False
                            NotifyIcon1.Icon = Icon.FromHandle(CType(My.Resources.Mute, Bitmap).GetHicon())

                            NotifyIcon1.Visible = True

                            NotifyIcon1.Text = "Microphone: Muted (hold Ctrl to talk)"


                        Catch
                        End Try
                    End If
                End If
            End If
        End If
        Return CallNextHookEx(hookId, nCode, wParam, lParam)
    End Function


    ' Add the SupportedOSPlatform attribute to the Form1_Load method to suppress CA1416
    <Global.System.Runtime.Versioning.SupportedOSPlatform("windows6.1")>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Text = "Version: " & My.Application.Info.Version.ToString()
        ' Pick audio device
        Dim enumerator = New MMDeviceEnumerator()
        Try
            mmDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Communications)
        Catch
            mmDevice = enumerator.GetDefaultAudioEndpoint(DataFlow.Capture, Role.Console)
        End Try

        SetMicMute(True)

        ' === INSTALL KEYBOARD HOOK (YOU WERE MISSING THIS) ===
        hookProc = AddressOf LowLevelKeyboardProc
        hookId = SetWindowsHookEx(
        WH_KEYBOARD_LL,
        hookProc,
        GetModuleHandle(Process.GetCurrentProcess().MainModule.ModuleName),
        0
)

        Label1.ForeColor = Color.Red
        Label1.Text = "Muted (hold Ctrl to talk)"
        PictureBox1.Image = My.Resources.Mic_Red
    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If hookId <> IntPtr.Zero Then UnhookWindowsHookEx(hookId)
        ' Optional: restore to unmuted on exit (or remember prior state yourself)
        SetMicMute(False)
    End Sub

    Private Sub SetMicMute(mute As Boolean)
        Try
            If mmDevice IsNot Nothing Then
                mmDevice.AudioEndpointVolume.Mute = mute
            End If
        Catch ex As Exception
            ' Handle errors (device removed, access policies, etc.)
            ' You might want to log ex.Message in a production app.
        End Try
    End Sub




End Class
