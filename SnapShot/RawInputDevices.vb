Public Class RawInputDevices
    Inherits RawInputNativeMethods
    Private Shared RIDEV_INPUTSINK As Integer = &H100
    Private Shared RID_INPUT As Integer = &H10000003
    Private Shared FAPPCOMMAND_MASK As Integer = &HF000
    Private Shared FAPPCOMMAND_MOUSE As Integer = &H8000
    Private Shared FAPPCOMMAND_OEM As Integer = &H1000
    Private Shared RIM_TYPEMOUSE As Integer = 0
    Private Shared RIM_TYPEKEYBOARD As Integer = 1
    Private Shared RIM_TYPEHID As Integer = 2
    Private Shared RIDI_DEVICENAME As Integer = &H20000007
    Private Shared RIDI_DEVICEINFO As Integer = &H2000000B
    Private Shared WM_KEYDOWN As Integer = &H100
    Private Shared WM_SYSKEYDOWN As Integer = &H104
    Private Shared WM_INPUT As Integer = &HFF
    Private Shared VK_OEM_CLEAR As Integer = &HFE
    Private Shared VK_LAST_KEY As Integer = VK_OEM_CLEAR 'this is a made up value used as a sentinel
    Private m_deviceNameList As New Hashtable()
    Private m_deviceInfoList As New Hashtable()
    Private m_usagePagesList As New Hashtable()

    Public Class DeviceNameInfo
        Public deviceName As String
        Public deviceType As String
        Public deviceHandle As IntPtr
        Public Name As String
        Public source As String
        Public key As UShort
        Public vKey As String
    End Class


    Private Function ReadReg(ByVal item As String, ByRef isKeyboard As Boolean) As String ' Example Device Identification string @"\??\ACPI#PNP0303#3&13c0b0c5&0#{884b96c3-56ef-11d1-bc8c-00a0c91405dd}";
        item = item.Substring(4) ' remove the \??\
        Dim split As String() = item.Split("#"c)
        Dim id_01 As String = split(0) ' ACPI (Class code)
        Dim id_02 As String = split(1) ' PNP0303 (SubClass code)
        Dim id_03 As String = split(2) ' 3&13c0b0c5&0 (Protocol code)
        Dim OurKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine 'Open the appropriate key as read-only so no permissions are needed.
        Dim findme As String = String.Format("System\CurrentControlSet\Enum\{0}\{1}\{2}", id_01, id_02, id_03)
        OurKey = OurKey.OpenSubKey(findme, False) 'Retrieve the desired information and set isKeyboard
        Dim deviceDesc As String = DirectCast(OurKey.GetValue("DeviceDesc"), String)
        Dim deviceClass As String = DirectCast(OurKey.GetValue("Class"), String)
        If deviceClass.ToUpper().Equals("KEYBOARD") Then
            isKeyboard = True
        Else
            isKeyboard = False
        End If
        Return deviceDesc
    End Function

    Private Function GetDeviceType(ByVal device As Integer) As String
        Dim deviceType As String
        Select Case device
            Case RIM_TYPEMOUSE
                deviceType = "MOUSE"
                Exit Select
            Case RIM_TYPEKEYBOARD
                deviceType = "KEYBOARD"
                Exit Select
            Case RIM_TYPEHID
                deviceType = "HID"
                Exit Select
            Case Else
                deviceType = "UNKNOWN"
                Exit Select
        End Select
        Return deviceType
    End Function

    Public Function EnumerateDevices() As Integer
        Dim result As Integer = 0
        Dim deviceCount As Integer = 0
        Dim dwSize As Integer = (System.Runtime.InteropServices.Marshal.SizeOf(GetType(RAWINPUTDEVICELIST)))
        If GetRawInputDeviceList(IntPtr.Zero, deviceCount, CInt(dwSize)) = 0 Then
            Dim pRawInputDeviceList As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(CInt(dwSize * deviceCount))
            GetRawInputDeviceList(pRawInputDeviceList, deviceCount, CInt(dwSize))
            For i As Long = 0 To deviceCount - 1
                Dim dName As DeviceNameInfo
                Dim deviceName As String
                Dim pcbSizeA As Integer = 0
                Dim pcbSizeB As Integer = 0
                Dim rid As RAWINPUTDEVICELIST = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStructure(New IntPtr((pRawInputDeviceList.ToInt32() + (dwSize * i))), GetType(RAWINPUTDEVICELIST)), RAWINPUTDEVICELIST)
                GetRawInputDeviceInfo(rid.hDevice, RIDI_DEVICENAME, IntPtr.Zero, pcbSizeA)
                If pcbSizeA > 0 Then
                    Dim pDataA As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(CInt(pcbSizeA))
                    GetRawInputDeviceInfo(rid.hDevice, RIDI_DEVICENAME, pDataA, pcbSizeA)
                    deviceName = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStringAnsi(pDataA), String)
                    If Not deviceName.ToUpper().Contains("ROOT") Then ' Drop the "root" keyboard and mouse devices used for Terminal Services and the Remote Desktop
                        dName = New DeviceNameInfo()
                        dName.deviceName = DirectCast(System.Runtime.InteropServices.Marshal.PtrToStringAnsi(pDataA), String)
                        dName.deviceHandle = rid.hDevice
                        dName.deviceType = GetDeviceType(rid.dwType)
                        GetRawInputDeviceInfo(rid.hDevice, RIDI_DEVICEINFO, IntPtr.Zero, pcbSizeB)
                        If pcbSizeB > 0 Then
                            Dim pDataB As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(CInt(pcbSizeB))
                            System.Runtime.InteropServices.Marshal.Copy(New Integer() {pcbSizeB}, 0, pDataB, 1)
                            GetRawInputDeviceInfo(rid.hDevice, RIDI_DEVICEINFO, pDataB, pcbSizeB)
                            Dim deviceInfoBuffer(pcbSizeB - 1) As Byte
                            System.Runtime.InteropServices.Marshal.Copy(pDataB, deviceInfoBuffer, 0, pcbSizeB)
                            System.Runtime.InteropServices.Marshal.FreeHGlobal(pDataB)
                            Dim dInfo As New DeviceInfo
                            dInfo.Size = BitConverter.ToInt32(deviceInfoBuffer, 0)
                            dInfo.Type = BitConverter.ToInt32(deviceInfoBuffer, 4)
                            Dim pcbSizeC As Integer = pcbSizeB - 8
                            If ((dInfo.Type = RIM_TYPEKEYBOARD) AndAlso (pcbSizeC >= System.Runtime.InteropServices.Marshal.SizeOf(GetType(DeviceInfoKeyboard)))) Then
                                dInfo.KeyboardInfo.Type = BitConverter.ToUInt32(deviceInfoBuffer, 8)
                                dInfo.KeyboardInfo.SubType = BitConverter.ToUInt32(deviceInfoBuffer, 12)
                                dInfo.KeyboardInfo.KeyboardMode = BitConverter.ToUInt32(deviceInfoBuffer, 16)
                                dInfo.KeyboardInfo.NumberOfFunctionKeys = BitConverter.ToUInt32(deviceInfoBuffer, 20)
                                dInfo.KeyboardInfo.NumberOfIndicators = BitConverter.ToUInt32(deviceInfoBuffer, 24)
                                dInfo.KeyboardInfo.NumberOfKeysTotal = BitConverter.ToUInt32(deviceInfoBuffer, 28)
                            ElseIf ((dInfo.Type = RIM_TYPEHID) AndAlso (pcbSizeC >= System.Runtime.InteropServices.Marshal.SizeOf(GetType(DeviceInfoHID)))) Then
                                dInfo.HIDInfo.VendorID = BitConverter.ToUInt32(deviceInfoBuffer, 8)
                                dInfo.HIDInfo.ProductID = BitConverter.ToUInt32(deviceInfoBuffer, 12)
                                dInfo.HIDInfo.VersionNumber = BitConverter.ToUInt32(deviceInfoBuffer, 16)
                                dInfo.HIDInfo.UsagePage = BitConverter.ToUInt16(deviceInfoBuffer, 20)
                                dInfo.HIDInfo.Usage = BitConverter.ToUInt16(deviceInfoBuffer, 22)
                                If Not m_usagePagesList.ContainsKey(dInfo.HIDInfo.UsagePage) Then
                                    m_usagePagesList.Add(dInfo.HIDInfo.UsagePage, dInfo.HIDInfo.UsagePage)
                                End If
                            ElseIf ((dInfo.Type = RIM_TYPEMOUSE) AndAlso (pcbSizeC >= System.Runtime.InteropServices.Marshal.SizeOf(GetType(DeviceInfoMouse)))) Then
                                dInfo.MouseInfo.ID = BitConverter.ToUInt32(deviceInfoBuffer, 8)
                                dInfo.MouseInfo.NumberOfButtons = BitConverter.ToUInt32(deviceInfoBuffer, 12)
                                dInfo.MouseInfo.SampleRate = BitConverter.ToUInt32(deviceInfoBuffer, 16)
                            End If
                            result += 1
                            m_deviceInfoList.Add(rid.hDevice, dInfo)
                            m_deviceNameList.Add(rid.hDevice, dName)
                        End If
                    End If
                    System.Runtime.InteropServices.Marshal.FreeHGlobal(pDataA)
                End If
            Next
            System.Runtime.InteropServices.Marshal.FreeHGlobal(pRawInputDeviceList)
        Else
            m_errorMessage = "An error occurred while retrieving the list of raw input devices."
            result = 0
        End If
        Return result
    End Function

    Public Sub New()
        EnumerateDevices()
    End Sub


    Public ReadOnly Property UniqueUsagePages() As UShort()
        Get
            Dim result() As UShort = New UShort() {}
            If m_usagePagesList.Count > 0 Then
                Try
                    m_usagePagesList.Keys.CopyTo(result, 0)
                Catch ex As Exception
                    result = New UShort() {}
                    Debug.WriteLine(ex.ToString)
                End Try
            End If
            Return result
        End Get
    End Property

End Class
