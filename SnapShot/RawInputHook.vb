Public Class RawInputHook
    Inherits RawInputNativeMethods
    Private m_SimpleMessageOnlyWindow As SimpleMessageOnlyWindow
    Public Event OnRawInputFromMouse(ByVal riHeader As RAWINPUTHEADER, ByVal riMouse As RAWMOUSE)
    Public Event OnRawInputFromKeyboard(ByVal riHeader As RAWINPUTHEADER, ByVal riKeyboard As RAWKEYBOARD)
    Public Event OnRawInputFromHID(ByVal riHeader As RAWINPUTHEADER, ByVal riHID As RAWHID)
    Public Event OnRawInputFromUnknown(ByVal riUnknownType As RawInput)
    Private Shared key_names As String() = {"", "{Left mouse button}", "{Right mouse button}", "{Control-break processing}", "{Middle mouse button}", "{2000/XP: X1 mouse button}", "{2000/XP: X2 mouse button}", "", "{BACKSPACE}", "{TAB}", "", "", "{CLEAR}", "{ENTER}", "", "", "{SHIFT}", "{CTRL}", "{ALT}", "{PAUSE}", "{CAPS LOCK}", "{Kana/Hangul/Hanguel mode}", "", "{Junja mode}", "{Final mode}", "{Hanja/Kanji mode}", "", "{ESC}", "{IME Convert}", "{IME Nonconvert}", "{IME Accept}", "{IME Mode Change Request}", "{SPACEBAR}", "{PAGE UP}", "{PAGE DOWN}", "{END}", "{HOME}", "{LEFT ARROW}", "{UP ARROW}", "{RIGHT ARROW}", "{DOWN ARROW}", "{SELECT}", "{PRINT}", "{EXECUTE}", "{PRINT SCREEN}", "{INSERT}", "{DELETE}", "{HELP}", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "", "", "", "", "", "", "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "{Left Windows/Start}", "{Right Windows key}", "{Applications OR Context-Menu}", "", "{Computer Sleep key}", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "*", "+", "|", "-", ".", "/", "{F1}", "{F2}", "{F3}", "{F4}", "{F5}", "{F6}", "{F7}", "{F8}", "{F9}", "{F10}", "{F11}", "{F12}", "{F13}", "{F14}", "{F15}", "{F16}", "{F17}", "{F18}", "{F19}", "{F20}", "{F21}", "{F22}", "{F23}", "{F24}", "", "", "", "", "", "", "", "", "{NUM LOCK}", "{SCROLL LOCK}", "=", "{OEM OR 'Unregister word' key}", "{OEM OR 'Register word' key}", "{OEM OR 'Left OYAYUBI' key}", "{OEM OR 'Right OYAYUBI' key}", "", "", "", "", "", "", "", "", "", "{Left SHIFT key}", "{Right SHIFT key}", "{Left CONTROL key}", "{Right CONTROL key}", "{Left MENU key}", "{Right MENU key}", "{2000/XP: Browser Back key}", "{2000/XP: Browser Forward key}", "{2000/XP: Browser Refresh key}", "{2000/XP: Browser Stop key}", "{2000/XP: Browser Search key}", "{2000/XP: Browser Favorites key}", "{2000/XP: Browser Start and Home key}", "{2000/XP: Volume Mute key}", "{2000/XP: Volume Down key}", "{2000/XP: Volume Up key}", "{2000/XP: Next Track key}", "{2000/XP: Previous Track key}", "{2000/XP: Stop Media key}", "{2000/XP: Play/Pause Media key}", "{2000/XP: Start Mail key}", "{2000/XP: Select Media key}", "{2000/XP: Start Application 1 key}", "{2000/XP: Start Application 2 key}", "", "", "{OEM_1, the ';:' key}", "{OEM_PLUS}", "{OEM_COMMA}", "{OEM_MINUS}", "{OEM_PERIOD}", "{OEM_2, the '/?' key}", "{OEM_3, the '~' key}", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "{2000/XP, the '[{' key}", "{2000/XP, the '\|' key}", "{2000/XP, the ']}' key}", "{2000/XP, the 'single/double-quote' key}", "{OEM_8}", "{ICO_F17 reserved}", "{OEM_AX key on Japanese AX kbd (OEM specific)}", "{OEM_102 OR OEM102 '<>' or '\|' on RT 102-key kbd.}", "{Help key on ICO (OEM specific)}", "{00 key on ICO (OEM specific)}", "{95/98/Me/NT4/2000/XP: IME PROCESS key}", "{ICO_CLEAR OEM specific}", "{2000/XP: Unicode, VK_PACKET = low 32-bits VKey}", "", "{OEM_RESET}", "{OEM_JUMP}", "{OEM_PA1}", "{OEM_PA2}", "{OEM_PA3}", "{OEM_WSCTRL}", "{OEM_CUSEL}", "{OEM_ATTN}", "{OEM_FINISH OEM_FINNISH}", "{OEM_COPY}", "{OEM_AUTO}", "{OEM_ENLW}", "{OEM_BACKTAB}", "{Attn key}", "{CrSel key}", "{ExSel key}", "{Erase EOF key}", "{Play key}", "{Zoom key}", "{Reserved}", "{PA1 key}", "{Clear key}", ""}

    Public Sub New()
        m_SimpleMessageOnlyWindow = New SimpleMessageOnlyWindow()
        AddHandler m_SimpleMessageOnlyWindow.OnWindowsMessageCreate, AddressOf SimpleMessageWindowCreate
        AddHandler m_SimpleMessageOnlyWindow.OnWindowsMessageInput, AddressOf SimpleMessageWindowInput
    End Sub

    Public Function FriendlyKeyname(ByVal KeyCode As Integer) As String
        Return key_names(KeyCode And 255)
    End Function

    Public Sub SimpleMessageWindowCreate(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
        'Dim rid() As RAWINPUTDEVICE = {New RAWINPUTDEVICE}
        Dim rid(1) As RAWINPUTDEVICE
        rid(0).dwFlags = RawInputDeviceFlags.InputSink Or RawInputDeviceFlags.NoHotKeys 'keyboard flags     RawInputDeviceFlags.NoLegacy Or 
        rid(0).usUsagePage = HIDUsagePage.Generic
        rid(0).usUsage = HIDUsage.Keyboard
        rid(0).hwndTarget = hWnd

        rid(1).dwFlags = RawInputDeviceFlags.InputSink
        rid(1).usUsagePage = HIDUsagePage.Generic
        rid(1).usUsage = HIDUsage.Mouse
        rid(1).hwndTarget = hWnd
        RegisterRawInputDevices(rid, rid.Length, System.Runtime.InteropServices.Marshal.SizeOf(rid(0)))
    End Sub

    Public Sub SimpleMessageWindowInput(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
        Dim ri As RawInput = Nothing
        Dim blen As Integer = 0
        Dim hlen As Integer = System.Runtime.InteropServices.Marshal.SizeOf(GetType(RAWINPUTHEADER))
        Dim bytesRead As Integer = -1

        bytesRead = GetRawInputData(lParam, RawInputCommand.Input, IntPtr.Zero, blen, hlen)
        If ((bytesRead = -1) OrElse (blen < 1)) Then
            Debug.WriteLine("GetRawInputData Error Retreiving Buffer size.")
        Else
            Dim riBuffer As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(blen)
            bytesRead = GetRawInputData(lParam, RawInputCommand.Input, riBuffer, blen, hlen)
            If bytesRead <> blen Then
                Debug.WriteLine("GetRawInputData Error Retreiving Buffer data.")
            Else
                Dim bb(bytesRead - 1) As Byte
                Try
                    ri = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(riBuffer, GetType(RawInput)), RawInput)
                    System.Runtime.InteropServices.Marshal.Copy(riBuffer, bb, 0, bb.Length)
                Catch ex As Exception
                    ri = Nothing
                End Try
                If ri.Equals(Nothing) Then
                    Debug.WriteLine("Error Casting Marshalled Buffer into RawInput structure.")
                Else
                    Dim deviceType As String
                    Select Case ri.Header.dwType
                        Case 0
                            deviceType = "MOUSE"
                            If ((Not OnRawInputFromMouseEvent Is Nothing) AndAlso (OnRawInputFromMouseEvent.GetInvocationList().Length > 0)) Then
                                RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnRawInputFromMouseEvent, New Object() {ri.Header, ri.Data.Mouse})
                            End If
                        Case 1
                            deviceType = "KEYBOARD"
                            If ((Not OnRawInputFromKeyboardEvent Is Nothing) AndAlso (OnRawInputFromKeyboardEvent.GetInvocationList().Length > 0)) Then
                                RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnRawInputFromKeyboardEvent, New Object() {ri.Header, ri.Data.Keyboard})
                            End If
                        Case 2
                            deviceType = "HID"
                            If ((Not OnRawInputFromHIDEvent Is Nothing) AndAlso (OnRawInputFromHIDEvent.GetInvocationList().Length > 0)) Then
                                RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnRawInputFromHIDEvent, New Object() {ri.Header, ri.Data.HID})
                            End If
                        Case Else
                            deviceType = "UNKNOWN"
                            If ((Not OnRawInputFromUnknownEvent Is Nothing) AndAlso (OnRawInputFromUnknownEvent.GetInvocationList().Length > 0)) Then
                                RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnRawInputFromUnknownEvent, New Object() {ri})
                            End If
                    End Select
                End If
            End If
            Try
                System.Runtime.InteropServices.Marshal.FreeHGlobal(riBuffer)
            Catch ex As Exception
                Debug.WriteLine("Error Freeing Marshalled Buffer for RawInput data.")
            End Try
        End If
    End Sub

    Public ReadOnly Property lastWin32Error() As Integer
        Get
            Return m_lastWin32Error
        End Get
    End Property

    Public ReadOnly Property errorMessage() As String
        Get
            Return m_errorMessage
        End Get
    End Property

End Class

