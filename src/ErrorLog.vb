''' <summary>ErrorLog manages the error log file and content.</summary>
''' <version>0.0.1.0</version>

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows

Namespace cvmd2html
  MustInherit _
  Class ErrorLog
    ''' <summary>The path to the generated error log file.</summary>
    Friend Shared ReadOnly Path As String = GenerateRandomPath(".log")

    ''' <summary>Display the content of the error log file in a message box.</summary>
    Shared _
    Sub Read()
      Try
        Using txtStream As StreamReader = File.OpenText(Path)
          ' Read the error message and remove the ANSI escaped character for red coloring.
          Dim errorMessage As String = Regex.Replace(txtStream.ReadToEnd(), "(\x1B\[31;1m)|(\x1B\[0m)", "")
          If Len(errorMessage) Then Popup(errorMessage, MessageBoxImage.Error)
        End Using
      Catch
      End Try
    End Sub

    ''' <summary>Delete the error log file.</summary>
    Shared _
    Sub Delete()
      DeleteFile(Path)
    End Sub
  End Class
End Namespace