Public Class RaiseEventUtility
    Public Shared Sub RaiseEventAndExecuteItInTheTargetThread(ByVal _event As System.MulticastDelegate, ByVal _ParamArray_args() As Object)
        If Not _event Is Nothing Then
            If _event.GetInvocationList().Length > 0 Then
                Dim _sync As System.ComponentModel.ISynchronizeInvoke = Nothing
                For Each _delegate As System.MulticastDelegate In _event.GetInvocationList()
                    If ((_sync Is Nothing) AndAlso (GetType(System.ComponentModel.ISynchronizeInvoke).IsAssignableFrom(_delegate.Target.GetType())) AndAlso (Not _delegate.Target.GetType().IsAbstract)) Then
                        Try
                            _sync = CType(_delegate.Target, System.ComponentModel.ISynchronizeInvoke)
                        Catch ex As Exception
                            Diagnostics.Debug.WriteLine(ex.ToString())
                            _sync = Nothing
                        End Try
                    End If
                    If _sync Is Nothing Then
                        Try
                            _delegate.DynamicInvoke(_ParamArray_args)
                        Catch ex As Exception
                            Diagnostics.Debug.WriteLine(ex.ToString())
                        End Try
                    Else
                        Try
                            _sync.Invoke(_delegate, _ParamArray_args)
                        Catch ex As Exception
                            Diagnostics.Debug.WriteLine(ex.ToString())
                        End Try
                    End If
                Next
            End If
        End If
    End Sub
End Class
