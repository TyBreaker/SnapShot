Public Class SimpleMessageOnlyWindow
    Inherits RawInputNativeMethods
    Private Shared WM_CREATE As Integer = 1
    Private Shared WM_CLOSE As Integer = 16
    Private Shared WM_INPUT As Integer = 255
    Private m_exitCode As Integer = 0
    Private m_messageCount As Integer = 0
    Private m_handle As IntPtr = IntPtr.Zero
    Private WithEvents m_BackgroundWorker As New System.ComponentModel.BackgroundWorker 'Asynchronous Non-Blocking Background Worker
    Public Event OnWindowsMessageCreate(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
    Public Event OnWindowsMessageInput(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
    Public Event OnWindowsMessageClose(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr)
    Public Event OnNativeGetMessage(ByVal Sender As Object, ByVal msgCount As Integer, ByVal iRet As Integer, ByVal msg As MSG)
    Private Structure OnWindowMessageParams
        Public hWnd As System.IntPtr
        Public msg As Integer
        Public wParam As IntPtr
        Public lParam As IntPtr
    End Structure
    Public Function WndProcCallback(ByVal hWnd As IntPtr, ByVal msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        Dim msgParams As New OnWindowMessageParams
        msgParams.hWnd = hWnd
        msgParams.msg = msg
        msgParams.wParam = wParam
        msgParams.lParam = lParam
        Select Case msg
            Case WM_CREATE
                m_BackgroundWorker.ReportProgress(31, msgParams)
            Case WM_INPUT
                m_BackgroundWorker.ReportProgress(32, msgParams)
            Case WM_CLOSE
                m_BackgroundWorker.ReportProgress(33, msgParams)
                PostQuitMessage(m_exitCode) '0)
                If Not m_BackgroundWorker.CancellationPending Then
                    m_BackgroundWorker.CancelAsync()
                End If
            Case Else
                Return DefWindowProc(hWnd, msg, wParam, lParam)
        End Select
        Return IntPtr.Zero
    End Function
    Private Sub m_BackgroundWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles m_BackgroundWorker.DoWork
        Dim msg As New MSG
        Dim wc As New WNDCLASS
        Dim hInstance As System.IntPtr = New System.IntPtr(System.Runtime.InteropServices.Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly.GetModules()(0)).ToInt32())
        wc.lpfnWndProc = AddressOf WndProcCallback
        wc.hInstance = hInstance
        wc.lpszClassName = "SimpleMessageWindow"
        RegisterClass(wc)
        m_handle = CreateWindowEx(0, wc.lpszClassName, Nothing, 0, 0, 0, 0, 0, New IntPtr(-3), IntPtr.Zero, hInstance, 0) '(0, wc.lpszClassName, Nothing, 0, 0, 0, 0, 0, -3, 0, hInstance, 0) 'HWND_MESSAGE = -3
        Dim iRet As Integer = 1
        Try
            Do Until ((iRet = 0) OrElse (iRet = -1))
                iRet = GetMessage(msg, m_handle, 0, 0)
                If iRet = -1 Then 'ERROR (GetLastError)
                    m_lastWin32Error = System.Runtime.InteropServices.Marshal.GetLastWin32Error()
                    m_exitCode = iRet
                ElseIf iRet = 0 Then 'WM_QUIT
                    m_exitCode = msg.wParam.ToInt32
                Else
                    m_BackgroundWorker.ReportProgress(iRet, msg)
                    TranslateMessage(msg)
                    DispatchMessage(msg)
                End If
            Loop
        Catch ex As Exception
            Debug.Write(ex.ToString) 'should never happen... (all unsafe method calls have handlers)
            iRet = -1
        End Try
    End Sub
    Private Sub m_BackgroundWorker_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles m_BackgroundWorker.ProgressChanged
        If (Not e.UserState Is Nothing) Then
            Dim iRet As Integer = e.ProgressPercentage
            If (e.UserState.GetType Is GetType(MSG)) Then
                m_messageCount += 1
                Dim msg As MSG = CType(e.UserState, MSG)
                If ((Not OnNativeGetMessageEvent Is Nothing) AndAlso (OnNativeGetMessageEvent.GetInvocationList().Length > 0)) Then
                    RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnNativeGetMessageEvent, New Object() {Me, m_messageCount, iRet, msg})
                End If
            ElseIf (e.UserState.GetType Is GetType(OnWindowMessageParams)) Then
                Dim msgParams As OnWindowMessageParams = DirectCast(e.UserState, OnWindowMessageParams)
                If iRet = 31 Then
                    If ((Not OnWindowsMessageCreateEvent Is Nothing) AndAlso (OnWindowsMessageCreateEvent.GetInvocationList().Length > 0)) Then
                        RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnWindowsMessageCreateEvent, New Object() {msgParams.hWnd, msgParams.msg, msgParams.wParam, msgParams.lParam})
                    End If
                ElseIf iRet = 32 Then
                    If ((Not OnWindowsMessageInputEvent Is Nothing) AndAlso (OnWindowsMessageInputEvent.GetInvocationList().Length > 0)) Then
                        RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnWindowsMessageInputEvent, New Object() {msgParams.hWnd, msgParams.msg, msgParams.wParam, msgParams.lParam})
                    End If
                ElseIf iRet = 33 Then
                    If ((Not OnWindowsMessageCloseEvent Is Nothing) AndAlso (OnWindowsMessageCloseEvent.GetInvocationList().Length > 0)) Then
                        RaiseEventUtility.RaiseEventAndExecuteItInTheTargetThread(OnWindowsMessageCloseEvent, New Object() {msgParams.hWnd, msgParams.msg, msgParams.wParam, msgParams.lParam})
                    End If
                End If
            End If
        End If
    End Sub
    Public Sub New()
        m_BackgroundWorker.WorkerReportsProgress = True
        m_BackgroundWorker.WorkerSupportsCancellation = True
        mybase_BackgroundWorker = m_BackgroundWorker
        m_BackgroundWorker.RunWorkerAsync()
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
    Public ReadOnly Property exitCode() As Integer
        Get
            Return m_exitCode
        End Get
    End Property
    Public ReadOnly Property messageCount() As Integer
        Get
            Return m_messageCount
        End Get
    End Property
    Public ReadOnly Property handle() As IntPtr
        Get
            Return m_handle
        End Get
    End Property
End Class











