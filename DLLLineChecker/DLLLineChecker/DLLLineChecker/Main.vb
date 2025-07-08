Imports System.Configuration
Imports System.IO


Public Class Main
    Shared WithEvents myTimer As New System.Timers.Timer
    Protected Overrides Sub OnStart(ByVal args() As String)
        ' Add code here to start your service. This method should set things
        ' in motion so your service can do its work.
        myTimer.Interval = System.Configuration.ConfigurationManager.AppSettings("CheckInterval").ToString()
        myTimer.Enabled = True

        Dim strDir As String
        Dim intStart As Integer
        Dim intEnd As Integer

        intEnd = System.Configuration.ConfigurationManager.AppSettings("EndCopyValue").ToString()
        intStart = System.Configuration.ConfigurationManager.AppSettings("StartCopyValue").ToString()
        strDir = System.Configuration.ConfigurationManager.AppSettings("CheckDir").ToString()

        My.Application.Log.WriteEntry("Checking directory: " & strDir, TraceEventType.Information)
        My.Application.Log.WriteEntry("Start Copy Value: " & intStart.ToString(), TraceEventType.Information)
        My.Application.Log.WriteEntry("End Copy Value: " & intEnd.ToString(), TraceEventType.Information)

    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
        'If the service stops, we want the timer to stop
        myTimer.Enabled = False
        myTimer.Stop()
    End Sub

    Shared Sub myTimer_Elapsed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles myTimer.Elapsed
        Dim strDir As String
        Dim intStart As Integer
        Dim intEnd As Integer
        Dim blnDebug As Boolean
        Try
            intEnd = System.Configuration.ConfigurationManager.AppSettings("EndCopyValue").ToString()
            intStart = System.Configuration.ConfigurationManager.AppSettings("StartCopyValue").ToString()
            strDir = System.Configuration.ConfigurationManager.AppSettings("CheckDir").ToString()
            blnDebug = System.Configuration.ConfigurationManager.AppSettings("DebugOutput").ToString()

            Dim myDir As New DirectoryInfo(strDir)
            Dim myFiles As FileInfo
            If blnDebug = True Then
                My.Application.Log.WriteEntry("Beginning processing of files", TraceEventType.Information)
            End If
            For Each myFiles In myDir.GetFileSystemInfos("*_1.DLL")

                If blnDebug = True Then
                    My.Application.Log.WriteEntry("File to process: " & myFiles.FullName.ToString())
                End If

                Dim i As Integer
                Dim newfile As String
                Dim j As Integer
                Dim updatedfile As String

                For i = intStart To intEnd
                    newfile = "_" & i.ToString()
                    If blnDebug = True Then
                        My.Application.Log.WriteEntry("New file: " & newfile)
                    End If

                    If File.Exists(myFiles.FullName.Replace("_1", newfile)) = False Then
                        Try
                            File.Copy(myFiles.FullName.ToString(), myFiles.FullName.Replace("_1", newfile), True)
                        Catch ex As Exception
                            My.Application.Log.WriteEntry("A file copy error has occurred: " & ex.Message, TraceEventType.Error)
                        End Try
                    Else
                        If File.GetLastWriteTime(myFiles.FullName.Replace("_1", newfile)) <> File.GetLastWriteTime(myFiles.FullName) Then
                            For j = 2 To intEnd
                                updatedfile = myFiles.FullName.Replace("_1", "_" & j)
                                File.Copy(myFiles.FullName.ToString(), updatedfile, True)
                            Next

                        End If
                    End If

                Next ' For i = intStart to intEnd
            Next ' For each myFiles
        Catch ex As Exception
            My.Application.Log.WriteEntry("An error has occurred: " & ex.Message, TraceEventType.Error)
        End Try

        If blnDebug = True Then
            My.Application.Log.WriteEntry("End of processing files.", TraceEventType.Information)
        End If

    End Sub
End Class
