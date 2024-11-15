''' <summary>Launch the shortcut's target PowerShell script with the markdown.</summary>
''' <version>0.0.1.2</version>

Imports System.Diagnostics
Imports System.Security.Principal
Imports System.Management

Namespace cvmd2html
  Module Program
    Sub Main(args As String())
      RequestAdminPrivileges(args)

      ' The application execution.
      If MarkdownPathParam IsNot Nothing Then
        Const CMD_EXE As String = "C:\Windows\System32\cmd.exe"
        Const CMD_LINE_FORMAT As String = "/d /c """"{0}"" 2> ""{1}"""""
        IconLink.Create(MarkdownPathParam)
        If WaitForExit(Process.Start(New ProcessStartInfo() With {
            .FileName = CMD_EXE,
            .Arguments = String.Format(CMD_LINE_FORMAT, IconLink.Path, ErrorLog.Path),
            .WindowStyle = ProcessWindowStyle.Hidden
          }).Id) Then
          ErrorLog.Read()
          ErrorLog.Delete()
        End If
        IconLink.Delete()
        Quit(0)
      End If

      ' Configuration and settings.
      If SetConfigParam Xor UnsetConfigParam Then
        If SetConfigParam Then
          SetShortcut()
          If CBool(NoIconConfigParam) Then
            RemoveIcon()
          Else
            AddIcon()
          End If
        ElseIf UnsetConfigParam Then
          UnsetShortcut()
        End If
        Quit(0)
      End If

      Quit(1)
    End Sub

    ''' <summary>Wait for the process exit.</summary>
    ''' <param name="processId">The process identifier.</param>
    ''' <returns>The exit status of the process.</returns>
    Private _
    Function WaitForExit(processId As Integer) As UInteger
      ' The process termination event query. Win32_ProcessStopTrace requires admin rights to be used.
      Dim wqlQuery As String = "SELECT * FROM Win32_ProcessStopTrace WHERE ProcessName='cmd.exe' AND ProcessId=" & processId
      ' Wait for the process to exit.
      Return (New ManagementEventWatcher(wqlQuery)).WaitForNextEvent()("ExitStatus")
    End Function

    ''' <summary>Request administrator privileges.</summary>
    ''' <param name="args">The command line arguments.</param>
    Private _
    Sub RequestAdminPrivileges(args As String())
      If IsCurrentProcessElevated() Then Exit Sub
      Try
        Process.Start(New ProcessStartInfo(AssemblyLocation, If(UBound(args) >= 0, String.Format("""{0}""", Join(args, """ """)), "")) With {
            .UseShellExecute = true,
            .Verb = "runas",
            .WindowStyle = ProcessWindowStyle.Hidden
          })
      Catch
        Quit(1)
      End Try
      Quit(0)
    End Sub

    ''' <summary>Check if the process is elevated.</summary>
    ''' <returns>True if the running process is elevated, false otherwise.</returns>
    Private _
    Function IsCurrentProcessElevated() As Boolean
      Return (New WindowsPrincipal(WindowsIdentity.GetCurrent)).IsInRole(WindowsBuiltInRole.Administrator)
    End Function
  End Module
End Namespace