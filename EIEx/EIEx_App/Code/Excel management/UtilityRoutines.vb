Module UtilityRoutines
    'Pas certain que ce code soit parfait. 
    'Pour une discution détaillée voir https://blogs.msdn.microsoft.com/visualstudio/2010/03/01/marshal-releasecomobject-considered-dangerous/

#Region "ExcelInMemory"
    ''' <summary>
    ''' Determines if Microsoft Excel process is currently in memory.
    ''' </summary>
    Public Function ExcelInMemory() As Boolean
        Return Process.GetProcesses().Any(Function(p) p.ProcessName.Contains("EXCEL"))
    End Function
#End Region

#Region "releaseObject"
    Public Sub releaseObject(ByVal obj As Object, Optional ByVal Collect As Boolean = False)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            If Collect Then
                GC.Collect()
            End If
        End Try
    End Sub
#End Region

End Module
