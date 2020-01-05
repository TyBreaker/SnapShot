Public MustInherit Class RawInputNativeMethods
    Friend Shared x As Integer = 0
    Friend Shared m_lastWin32Error As Integer = 0
    Friend Shared m_errorMessage As String = ""
    Friend Shared mybase_BackgroundWorker As System.ComponentModel.BackgroundWorker = Nothing
    Private Class UnSafeNativeMethods
        <System.Runtime.InteropServices.DllImport("user32.dll")> _
        Friend Shared Function GetRawInputDeviceList(ByVal pRawInputDeviceList As IntPtr, ByRef uiNumDevices As Integer, ByVal cbSize As Integer) As Integer
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll")> _
        Friend Shared Function GetRawInputDeviceInfo(ByVal hDevice As IntPtr, ByVal uiCommand As Integer, ByVal pData As IntPtr, ByRef pcbSize As Integer) As Integer
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=Runtime.InteropServices.CharSet.Auto, EntryPoint:="RegisterRawInputDevices")> _
        Friend Shared Function RegisterRawInputDevices(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal pRawInputDevices As RAWINPUTDEVICE(), ByVal uiNumDevices As Integer, ByVal cbSize As Integer) As Boolean
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll")> _
        Friend Shared Function GetRawInputData(ByVal hRawInput As IntPtr, ByVal uiCommand As RawInputCommand, ByVal pData As IntPtr, ByRef pcbSize As Integer, ByVal cbSizeHeader As Integer) As Integer
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function CreateWindowEx(ByVal dwExStyle As Integer, ByVal lpszClassName As String, ByVal lpszWindowName As String, ByVal style As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal hWndParent As IntPtr, ByVal hMenu As IntPtr, ByVal hInst As IntPtr, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.AsAny)> ByVal pvParam As Object) As IntPtr
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function RegisterClass(ByVal wc As WNDCLASS) As Short
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True)> _
        Friend Shared Function TranslateMessage(<System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out()> ByRef msg As MSG) As Boolean
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function DispatchMessage(<System.Runtime.InteropServices.In()> ByRef msg As MSG) As Integer
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function DefWindowProc(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto, ExactSpelling:=True)> _
        Friend Shared Sub PostQuitMessage(ByVal nExitCode As Integer)
        End Sub
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function GetMessage(<System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out()> ByRef msg As MSG, ByVal hWnd As IntPtr, ByVal uMsgFilterMin As Integer, ByVal uMsgFilterMax As Integer) As Integer
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
        End Function
        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
        Friend Shared Function GetForegroundWindow() As IntPtr
        End Function
    End Class
    Friend Shared Sub OnWin32Error(ByVal ex As Exception)
        Debug.WriteLine(ex.ToString)
        m_errorMessage = ex.ToString
        m_lastWin32Error = System.Runtime.InteropServices.Marshal.GetLastWin32Error
        If ((Not mybase_BackgroundWorker Is Nothing) AndAlso (Not mybase_BackgroundWorker.CancellationPending)) Then
            mybase_BackgroundWorker.CancelAsync()
        End If
    End Sub
    Friend Shared Function CreateWindowEx(ByVal dwExStyle As Integer, ByVal lpszClassName As String, ByVal lpszWindowName As String, ByVal style As Integer, ByVal x As Integer, ByVal y As Integer, ByVal width As Integer, ByVal height As Integer, ByVal hWndParent As IntPtr, ByVal hMenu As IntPtr, ByVal hInst As IntPtr, <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.AsAny)> ByVal pvParam As Object) As IntPtr
        Dim result As IntPtr = IntPtr.Zero
        Try
            result = UnSafeNativeMethods.CreateWindowEx(dwExStyle, lpszClassName, lpszWindowName, style, x, y, width, height, hWndParent, hMenu, hInst, pvParam)
        Catch ex As Exception
            OnWin32Error(ex)
            result = IntPtr.Zero
        End Try
        Return result
    End Function
    Friend Shared Function RegisterClass(ByVal wc As WNDCLASS) As Short
        Dim result As Short = -1
        Try
            result = UnSafeNativeMethods.RegisterClass(wc)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function TranslateMessage(<System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out()> ByRef msg As MSG) As Boolean
        Dim result As Boolean = False
        Try
            result = UnSafeNativeMethods.TranslateMessage(msg)
        Catch ex As Exception
            OnWin32Error(ex)
            result = False
        End Try
        Return result
    End Function
    Friend Shared Function DispatchMessage(<System.Runtime.InteropServices.In()> ByRef msg As MSG) As Integer
        Dim result As Integer = -1
        Try
            result = UnSafeNativeMethods.DispatchMessage(msg)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function DefWindowProc(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        Dim result As IntPtr = IntPtr.Zero
        Try
            result = UnSafeNativeMethods.DefWindowProc(hWnd, msg, wParam, lParam)
        Catch ex As Exception
            OnWin32Error(ex)
            result = IntPtr.Zero
        End Try
        Return result
    End Function
    Friend Shared Sub PostQuitMessage(ByVal nExitCode As Integer)
        Try
            UnSafeNativeMethods.PostQuitMessage(nExitCode)
        Catch ex As Exception
            OnWin32Error(ex)
        End Try
    End Sub
    Friend Shared Function GetMessage(<System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out()> ByRef msg As MSG, ByVal hWnd As IntPtr, ByVal uMsgFilterMin As Integer, ByVal uMsgFilterMax As Integer) As Integer
        Dim result As Integer = -1
        Try
            result = UnSafeNativeMethods.GetMessage(msg, hWnd, uMsgFilterMin, uMsgFilterMax)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function RegisterRawInputDevices(<System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal pRawInputDevices As RAWINPUTDEVICE(), ByVal uiNumDevices As Integer, ByVal cbSize As Integer) As Boolean
        Dim result As Boolean = False
        Try
            result = UnSafeNativeMethods.RegisterRawInputDevices(pRawInputDevices, uiNumDevices, cbSize)
            If Not result Then
                OnWin32Error(New Exception("Native Method RegisterRawInputDevices returned False"))
            End If
        Catch ex As Exception
            OnWin32Error(ex)
            result = False
        End Try
        Return result
    End Function
    Friend Shared Function GetRawInputData(ByVal hRawInput As IntPtr, ByVal uiCommand As RawInputCommand, ByVal pData As IntPtr, ByRef pcbSize As Integer, ByVal cbSizeHeader As Integer) As Integer
        Dim result As Integer = -1
        Try
            result = UnSafeNativeMethods.GetRawInputData(hRawInput, uiCommand, pData, pcbSize, cbSizeHeader)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function GetRawInputDeviceList(ByVal pRawInputDeviceList As IntPtr, ByRef uiNumDevices As Integer, ByVal cbSize As Integer) As Integer
        Dim result As Integer = -1
        Try
            result = UnSafeNativeMethods.GetRawInputDeviceList(pRawInputDeviceList, uiNumDevices, cbSize)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function GetRawInputDeviceInfo(ByVal hDevice As IntPtr, ByVal uiCommand As Integer, <System.Runtime.InteropServices.In()> <System.Runtime.InteropServices.Out()> <System.Runtime.InteropServices.Optional()> ByVal pData As IntPtr, <System.Runtime.InteropServices.In()> <System.Runtime.InteropServices.Out()> ByRef pcbSize As Integer) As Integer
        Dim result As Integer = -1
        Try
            result = UnSafeNativeMethods.GetRawInputDeviceInfo(hDevice, uiCommand, pData, pcbSize)
        Catch ex As Exception
            OnWin32Error(ex)
            result = -1
        End Try
        Return result
    End Function
    Friend Shared Function PostMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
        Debug.WriteLine(hWnd.ToString("X16") & " : " & Msg.ToString("X16") & " : " & wParam.ToString("X16") & " : " & lParam.ToString("X16"))
        x = x + 1
        If x = 30 Then
            Debug.WriteLine("---------")
        End If
        Dim result As Boolean = False
        Try
            result = UnSafeNativeMethods.PostMessage(hWnd, Msg, wParam, lParam)
        Catch ex As Exception
            OnWin32Error(ex)
            result = False
        End Try
        Return result
    End Function
    Friend Shared Function GetForegroundWindow() As IntPtr
        Dim result As IntPtr = IntPtr.Zero
        Try
            result = UnSafeNativeMethods.GetForegroundWindow()
        Catch ex As Exception
            OnWin32Error(ex)
            result = IntPtr.Zero
        End Try
        Return result
    End Function
End Class




Public Delegate Function WndProc(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr


<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure MSG
    Public hwnd As IntPtr
    Public message As Integer
    Public wParam As IntPtr
    Public lParam As IntPtr
    Public time As Integer
    Public pt_x As Integer
    Public pt_y As Integer
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet:=System.Runtime.InteropServices.CharSet.Auto)> _
Public Class WNDCLASS
    Public style As Integer = 0
    Public lpfnWndProc As WndProc = Nothing
    Public cbClsExtra As Integer = 0
    Public cbWndExtra As Integer = 0
    Public hInstance As IntPtr = IntPtr.Zero
    Public hIcon As IntPtr = IntPtr.Zero
    Public hCursor As IntPtr = IntPtr.Zero
    Public hbrBackground As IntPtr = IntPtr.Zero
    Public lpszMenuName As String = ""
    Public lpszClassName As String = ""
End Class

''' <summary>HID usage page flags.</summary>
Public Enum HIDUsagePage As UShort
    ''' <summary>Unknown usage page.</summary>
    Undefined = &H0
    ''' <summary>Generic desktop controls.</summary>
    Generic = &H1
    ''' <summary>Simulation controls.</summary>
    Simulation = &H2
    ''' <summary>Virtual reality controls.</summary>
    VR = &H3
    ''' <summary>Sports controls.</summary>
    Sport = &H4
    ''' <summary>Games controls.</summary>
    Game = &H5
    ''' <summary>Keyboard controls.</summary>
    Keyboard = &H7
    ''' <summary>LED controls.</summary>
    LED = &H8
    ''' <summary>Button.</summary>
    Button = &H9
    ''' <summary>Ordinal.</summary>
    Ordinal = &HA
    ''' <summary>Telephony.</summary>
    Telephony = &HB
    ''' <summary>Consumer.</summary>
    Consumer = &HC
    ''' <summary>Digitizer.</summary>
    Digitizer = &HD
    ''' <summary>Physical interface device.</summary>
    PID = &HF
    ''' <summary>Unicode.</summary>
    Unicode = &H10
    ''' <summary>Alphanumeric display.</summary>
    AlphaNumeric = &H14
    ''' <summary>Medical instruments.</summary>
    Medical = &H40
    ''' <summary>Monitor page 0.</summary>
    MonitorPage0 = &H80
    ''' <summary>Monitor page 1.</summary>
    MonitorPage1 = &H81
    ''' <summary>Monitor page 2.</summary>
    MonitorPage2 = &H82
    ''' <summary>Monitor page 3.</summary>
    MonitorPage3 = &H83
    ''' <summary>Power page 0.</summary>
    PowerPage0 = &H84
    ''' <summary>Power page 1.</summary>
    PowerPage1 = &H85
    ''' <summary>Power page 2.</summary>
    PowerPage2 = &H86
    ''' <summary>Power page 3.</summary>
    PowerPage3 = &H87
    ''' <summary>Bar code scanner.</summary>
    BarCode = &H8C
    ''' <summary>Scale page.</summary>
    Scale = &H8D
    ''' <summary>Magnetic strip reading devices.</summary>
    MSR = &H8E
End Enum

''' <summary>Enumeration containing the HID usage values.</summary>
Public Enum HIDUsage As UShort
    ''' <summary></summary>
    Pointer = &H1
    ''' <summary></summary>
    Mouse = &H2
    ''' <summary></summary>
    Joystick = &H4
    ''' <summary></summary>
    Gamepad = &H5
    ''' <summary></summary>
    Keyboard = &H6
    ''' <summary></summary>
    Keypad = &H7
    ''' <summary></summary>
    SystemControl = &H80
    ''' <summary></summary>
    X = &H30
    ''' <summary></summary>
    Y = &H31
    ''' <summary></summary>
    Z = &H32
    ''' <summary></summary>
    RelativeX = &H33
    ''' <summary></summary>    
    RelativeY = &H34
    ''' <summary></summary>
    RelativeZ = &H35
    ''' <summary></summary>
    Slider = &H36
    ''' <summary></summary>
    Dial = &H37
    ''' <summary></summary>
    Wheel = &H38
    ''' <summary></summary>
    HatSwitch = &H39
    ''' <summary></summary>
    CountedBuffer = &H3A
    ''' <summary></summary>
    ByteCount = &H3B
    ''' <summary></summary>
    MotionWakeup = &H3C
    ''' <summary></summary>
    VX = &H40
    ''' <summary></summary>
    VY = &H41
    ''' <summary></summary>
    VZ = &H42
    ''' <summary></summary>
    VBRX = &H43
    ''' <summary></summary>
    VBRY = &H44
    ''' <summary></summary>
    VBRZ = &H45
    ''' <summary></summary>
    VNO = &H46
    ''' <summary></summary>
    SystemControlPower = &H81
    ''' <summary></summary>
    SystemControlSleep = &H82
    ''' <summary></summary>
    SystemControlWake = &H83
    ''' <summary></summary>
    SystemControlContextMenu = &H84
    ''' <summary></summary>
    SystemControlMainMenu = &H85
    ''' <summary></summary>
    SystemControlApplicationMenu = &H86
    ''' <summary></summary>
    SystemControlHelpMenu = &H87
    ''' <summary></summary>
    SystemControlMenuExit = &H88
    ''' <summary></summary>
    SystemControlMenuSelect = &H89
    ''' <summary></summary>
    SystemControlMenuRight = &H8A
    ''' <summary></summary>
    SystemControlMenuLeft = &H8B
    ''' <summary></summary>
    SystemControlMenuUp = &H8C
    ''' <summary></summary>
    SystemControlMenuDown = &H8D
    ''' <summary></summary>
    KeyboardNoEvent = &H0
    ''' <summary></summary>
    KeyboardRollover = &H1
    ''' <summary></summary>
    KeyboardPostFail = &H2
    ''' <summary></summary>
    KeyboardUndefined = &H3
    ''' <summary></summary>
    KeyboardaA = &H4
    ''' <summary></summary>
    KeyboardzZ = &H1D
    ''' <summary></summary>
    Keyboard1 = &H1E
    ''' <summary></summary>
    Keyboard0 = &H27
    ''' <summary></summary>
    KeyboardLeftControl = &HE0
    ''' <summary></summary>
    KeyboardLeftShift = &HE1
    ''' <summary></summary>
    KeyboardLeftALT = &HE2
    ''' <summary></summary>
    KeyboardLeftGUI = &HE3
    ''' <summary></summary>
    KeyboardRightControl = &HE4
    ''' <summary></summary>
    KeyboardRightShift = &HE5
    ''' <summary></summary>
    KeyboardRightALT = &HE6
    ''' <summary></summary>
    KeyboardRightGUI = &HE7
    ''' <summary></summary>
    KeyboardScrollLock = &H47
    ''' <summary></summary>
    KeyboardNumLock = &H53
    ''' <summary></summary>
    KeyboardCapsLock = &H39
    ''' <summary></summary>
    KeyboardF1 = &H3A
    ''' <summary></summary>
    KeyboardF12 = &H45
    ''' <summary></summary>
    KeyboardReturn = &H28
    ''' <summary></summary>
    KeyboardEscape = &H29
    ''' <summary></summary>
    KeyboardDelete = &H2A
    ''' <summary></summary>
    KeyboardPrintScreen = &H46
    ''' <summary></summary>
    LEDNumLock = &H1
    ''' <summary></summary>
    LEDCapsLock = &H2
    ''' <summary></summary>
    LEDScrollLock = &H3
    ''' <summary></summary>
    LEDCompose = &H4
    ''' <summary></summary>
    LEDKana = &H5
    ''' <summary></summary>
    LEDPower = &H6
    ''' <summary></summary>
    LEDShift = &H7
    ''' <summary></summary>
    LEDDoNotDisturb = &H8
    ''' <summary></summary>
    LEDMute = &H9
    ''' <summary></summary>
    LEDToneEnable = &HA
    ''' <summary></summary>
    LEDHighCutFilter = &HB
    ''' <summary></summary>
    LEDLowCutFilter = &HC
    ''' <summary></summary>
    LEDEqualizerEnable = &HD
    ''' <summary></summary>
    LEDSoundFieldOn = &HE
    ''' <summary></summary>
    LEDSurroundFieldOn = &HF
    ''' <summary></summary>
    LEDRepeat = &H10
    ''' <summary></summary>
    LEDStereo = &H11
    ''' <summary></summary>
    LEDSamplingRateDirect = &H12
    ''' <summary></summary>
    LEDSpinning = &H13
    ''' <summary></summary>
    LEDCAV = &H14
    ''' <summary></summary>
    LEDCLV = &H15
    ''' <summary></summary>
    LEDRecordingFormatDet = &H16
    ''' <summary></summary>
    LEDOffHook = &H17
    ''' <summary></summary>
    LEDRing = &H18
    ''' <summary></summary>
    LEDMessageWaiting = &H19
    ''' <summary></summary>
    LEDDataMode = &H1A
    ''' <summary></summary>
    LEDBatteryOperation = &H1B
    ''' <summary></summary>
    LEDBatteryOK = &H1C
    ''' <summary></summary>
    LEDBatteryLow = &H1D
    ''' <summary></summary>
    LEDSpeaker = &H1E
    ''' <summary></summary>
    LEDHeadset = &H1F
    ''' <summary></summary>
    LEDHold = &H20
    ''' <summary></summary>
    LEDMicrophone = &H21
    ''' <summary></summary>
    LEDCoverage = &H22
    ''' <summary></summary>
    LEDNightMode = &H23
    ''' <summary></summary>
    LEDSendCalls = &H24
    ''' <summary></summary>
    LEDCallPickup = &H25
    ''' <summary></summary>
    LEDConference = &H26
    ''' <summary></summary>
    LEDStandBy = &H27
    ''' <summary></summary>
    LEDCameraOn = &H28
    ''' <summary></summary>
    LEDCameraOff = &H29
    ''' <summary></summary>    
    LEDOnLine = &H2A
    ''' <summary></summary>
    LEDOffLine = &H2B
    ''' <summary></summary>
    LEDBusy = &H2C
    ''' <summary></summary>
    LEDReady = &H2D
    ''' <summary></summary>
    LEDPaperOut = &H2E
    ''' <summary></summary>
    LEDPaperJam = &H2F
    ''' <summary></summary>
    LEDRemote = &H30
    ''' <summary></summary>
    LEDForward = &H31
    ''' <summary></summary>
    LEDReverse = &H32
    ''' <summary></summary>
    LEDStop = &H33
    ''' <summary></summary>
    LEDRewind = &H34
    ''' <summary></summary>
    LEDFastForward = &H35
    ''' <summary></summary>
    LEDPlay = &H36
    ''' <summary></summary>
    LEDPause = &H37
    ''' <summary></summary>
    LEDRecord = &H38
    ''' <summary></summary>
    LEDError = &H39
    ''' <summary></summary>
    LEDSelectedIndicator = &H3A
    ''' <summary></summary>
    LEDInUseIndicator = &H3B
    ''' <summary></summary>
    LEDMultiModeIndicator = &H3C
    ''' <summary></summary>
    LEDIndicatorOn = &H3D
    ''' <summary></summary>
    LEDIndicatorFlash = &H3E
    ''' <summary></summary>
    LEDIndicatorSlowBlink = &H3F
    ''' <summary></summary>
    LEDIndicatorFastBlink = &H40
    ''' <summary></summary>
    LEDIndicatorOff = &H41
    ''' <summary></summary>
    LEDFlashOnTime = &H42
    ''' <summary></summary>
    LEDSlowBlinkOnTime = &H43
    ''' <summary></summary>
    LEDSlowBlinkOffTime = &H44
    ''' <summary></summary>
    LEDFastBlinkOnTime = &H45
    ''' <summary></summary>
    LEDFastBlinkOffTime = &H46
    ''' <summary></summary>
    LEDIndicatorColor = &H47
    ''' <summary></summary>
    LEDRed = &H48
    ''' <summary></summary>
    LEDGreen = &H49
    ''' <summary></summary>
    LEDAmber = &H4A
    ''' <summary></summary>
    LEDGenericIndicator = &H3B
    ''' <summary></summary>
    TelephonyPhone = &H1
    ''' <summary></summary>
    TelephonyAnsweringMachine = &H2
    ''' <summary></summary>
    TelephonyMessageControls = &H3
    ''' <summary></summary>
    TelephonyHandset = &H4
    ''' <summary></summary>
    TelephonyHeadset = &H5
    ''' <summary></summary>
    TelephonyKeypad = &H6
    ''' <summary></summary>
    TelephonyProgrammableButton = &H7
    ''' <summary></summary>
    SimulationRudder = &HBA
    ''' <summary></summary>
    SimulationThrottle = &HBB
End Enum

''' <summary>RawInput device flags.</summary>
<Flags()> _
Public Enum RawInputDeviceFlags As Integer
    ''' <summary>No flags.</summary>
    None = 0
    ''' <summary>If set, this removes the top level collection from the inclusion list. This tells the operating system to stop reading from a device which matches the top level collection.</summary>
    Remove = &H1
    ''' <summary>If set, this specifies the top level collections to exclude when reading a complete usage page. This flag only affects a TLC whose usage page is already specified with PageOnly.</summary>
    Exclude = &H10
    ''' <summary>If set, this specifies all devices whose top level collection is from the specified usUsagePage. Note that Usage must be zero. To exclude a particular top level collection, use Exclude.</summary>
    PageOnly = &H20
    ''' <summary>If set, this prevents any devices specified by UsagePage or Usage from generating legacy messages. This is only for the mouse and keyboard.</summary>
    NoLegacy = &H30
    ''' <summary>If set, this enables the caller to receive the input even when the caller is not in the foreground. Note that WindowHandle must be specified.</summary>
    InputSink = &H100
    ''' <summary>If set, the mouse button click does not activate the other window.</summary>
    CaptureMouse = &H200
    ''' <summary>If set, the application-defined keyboard device hotkeys are not handled. However, the system hotkeys; for example, ALT+TAB and CTRL+ALT+DEL, are still handled. By default, all keyboard hotkeys are handled. NoHotKeys can be specified even if NoLegacy is not specified and WindowHandle is NULL.</summary>
    NoHotKeys = &H200
    ''' <summary>If set, application keys are handled.  NoLegacy must be specified.  Keyboard only.</summary>
    AppKeys = &H400
End Enum

''' <summary> Enumeration containing the type device the raw input is coming from. </summary>
Public Enum RawInputType As Integer
    ''' <summary> Mouse input. </summary>
    RIM_TYPEMOUSE = 0
    ''' <summary> Keyboard input. </summary>
    RIM_TYPEKEYBOARD = 1
    ''' <summary> Another device that is not the keyboard or the mouse. </summary>
    RIM_TYPEHID = 2
End Enum

''' <summary> Enumeration contanining the command types to issue. </summary>
Public Enum RawInputCommand As Integer
    ''' <summary> Get input data. </summary>
    Input = &H10000003
    ''' <summary> Get header data. </summary>
    Header = &H10000005
End Enum

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RAWINPUTDEVICELIST
    Public hDevice As IntPtr
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwType As Integer
End Structure

'<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)> _
'Public Structure RAWINPUT
'    <System.Runtime.InteropServices.FieldOffset(0)> _
'    Public header As RAWINPUTHEADER
'    <System.Runtime.InteropServices.FieldOffset(16)> _
'    Public mouse As RAWMOUSE
'    <System.Runtime.InteropServices.FieldOffset(16)> _
'    Public keyboard As RAWKEYBOARD
'    <System.Runtime.InteropServices.FieldOffset(16)> _
'    Public hid As RAWHID
'End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RawInput
    Public Header As RAWINPUTHEADER
    Public Data As Union
    <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)> _
    Public Structure Union
        <System.Runtime.InteropServices.FieldOffset(0)> _
        Public Mouse As RAWMOUSE
        <System.Runtime.InteropServices.FieldOffset(0)> _
        Public Keyboard As RAWKEYBOARD
        <System.Runtime.InteropServices.FieldOffset(0)> _
        Public HID As RAWHID
    End Structure
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RAWINPUTHEADER
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwType As Integer
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwSize As Integer
    Public hDevice As IntPtr
    Public wParam As IntPtr
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RAWHID
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwSizeHID As Integer
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwCount As Integer
    Public pData As IntPtr
    Public ReadOnly Property Length() As Integer
        Get
            Return dwSizeHID
        End Get
    End Property
    Public ReadOnly Property data(ByVal index As Integer) As Byte()
        Get
            Dim result() As Byte = New Byte(0) {}
            If ((dwCount > 0) AndAlso (index < dwSizeHID)) Then
                result = New Byte(dwCount - 1) {}
                System.Runtime.InteropServices.Marshal.Copy(pData, result, CInt(index * dwCount), dwCount)
            End If
            Return result
        End Get
    End Property
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure BUTTONSSTR
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public usButtonFlags As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public usButtonData As UShort
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)> _
Public Structure RAWMOUSE
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    <System.Runtime.InteropServices.FieldOffset(0)> _
    Public usFlags As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    <System.Runtime.InteropServices.FieldOffset(4)> _
    Public ulButtons As UInteger
    <System.Runtime.InteropServices.FieldOffset(4)> _
    Public buttonsStr As BUTTONSSTR
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    <System.Runtime.InteropServices.FieldOffset(8)> _
    Public ulRawButtons As UInteger
    <System.Runtime.InteropServices.FieldOffset(12)> _
    Public lLastX As Integer
    <System.Runtime.InteropServices.FieldOffset(16)> _
    Public lLastY As Integer
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    <System.Runtime.InteropServices.FieldOffset(20)> _
    Public ulExtraInformation As UInteger
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RAWKEYBOARD
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public MakeCode As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public Flags As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public Reserved As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public VKey As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public Message As UInteger
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public ExtraInformation As UInteger
End Structure

<System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)> _
Public Structure RAWINPUTDEVICE
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public usUsagePage As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U2)> _
    Public usUsage As UShort
    <System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.U4)> _
    Public dwFlags As Integer
    Public hwndTarget As IntPtr
End Structure













Public Structure DeviceInfo
    Public Size As Integer
    Public Type As Integer
    Public MouseInfo As DeviceInfoMouse
    Public KeyboardInfo As DeviceInfoKeyboard
    Public HIDInfo As DeviceInfoHID
End Structure

Public Structure DeviceInfoMouse
    Public ID As UInteger
    Public NumberOfButtons As UInteger
    Public SampleRate As UInteger
End Structure

Public Structure DeviceInfoKeyboard
    Public Type As UInteger
    Public SubType As UInteger
    Public KeyboardMode As UInteger
    Public NumberOfFunctionKeys As UInteger
    Public NumberOfIndicators As UInteger
    Public NumberOfKeysTotal As UInteger
End Structure

Public Structure DeviceInfoHID
    Public VendorID As UInteger
    Public ProductID As UInteger
    Public VersionNumber As UInteger
    Public UsagePage As UShort
    Public Usage As UShort
End Structure

