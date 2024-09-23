''' <summary>Information about the resource files used by the project.</summary>
''' <version>0.0.1.0</version>

Imports System.IO
Imports System.Reflection
Imports Win32 = Microsoft.Win32
Imports IWshRuntimeLibrary

Namespace cvmd2html
  Module Package
    Const POWERSHELL_SUBKEY = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\pwsh.exe"

    ''' <summary>The path to the application.</summary>
    Friend ReadOnly AssemblyLocation As String = Assembly.GetExecutingAssembly.Location

    ''' <summary>The project root path</summary>
    ReadOnly root As String = AppContext.BaseDirectory

    ''' <summary>The powershell core runtime path.</summary>
    Friend ReadOnly PwshExePath As String = Win32.Registry.GetValue(POWERSHELL_SUBKEY, Nothing, Nothing).ToString

    ''' <summary>The project resources directory path.</summary>
    ReadOnly resourcePath As String = Path.Combine(root, "rsc")

    ''' <summary>The shortcut target powershell script path.</summary>
    Friend ReadOnly PwshScriptPath As String = Path.Combine(resourcePath, "cvmd2html.ps1")

    ''' <summary>adapted link object.</summary>
    MustInherit _
    Class IconLink
      ''' <summary>The custom icon link full path.</summary>
      Friend Shared ReadOnly Path As String = GenerateRandomPath(".lnk")

      ''' <summary>Create the custom icon link file.</summary>
      ''' <param name="markdownPath">The input markdown file path.</param>
      Shared _
      Sub Create(markdownPath As String)
        Dim shell As New WshShell
        Dim link As WshShortcut = shell.CreateShortcut(Path)
        With link
          .TargetPath = PwshExePath
          .Arguments = String.Format("-ep Bypass -nop -w Hidden -f ""{0}"" -Markdown ""{1}""", PwshScriptPath, markdownPath)
          .IconLocation = AssemblyLocation
          .Save
        End With
        ReleaseComObject(link)
        ReleaseComObject(shell)
      End Sub

      ''' <summary>Delete the custom icon link file.</summary>
      Shared _
      Sub Delete
        DeleteFile(Path)
      End Sub
    End Class
  End Module
End Namespace