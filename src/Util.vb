''' <summary>Some utility methods.</summary>
''' <version>0.0.1.0</version>

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows

Namespace cvmd2html
  Module Util
    ''' <summary>Delete the specified file.</summary>
    ''' <param name="extension">The file extension.</param>
    ''' <returns>A random file name.</returns>
    Function GenerateRandomPath(extension As String) As String
      Return Path.Combine(Path.GetTempPath(), Guid.NewGuid.ToString() & ".tmp" & extension)
    End Function

    ''' <summary>Delete the specified file.</summary>
    ''' <param name="filePath">The file path.</param>
    Sub DeleteFile(filePath As String)
      Try
        File.Delete(filePath)
      Catch
      End Try
    End Sub

    ''' <summary>Show the application message box.</summary>
    ''' <param name="messageText">The message text to show.</param>
    ''' <param name="popupType">The type of popup box.</param>
    ''' <param name="popupButtons">The buttons of the message box.</param>
    Sub Popup(messageText As String, Optional popupType As MessageBoxImage = MessageBoxImage.None, Optional popupButtons As MessageBoxButton = MessageBoxButton.OK)
      MessageBox.Show(messageText, "Convert to HTML", popupButtons, popupType)
    End Sub

    ''' <summary>Release the specified COM object.</summary>
    ''' <param name="comObject">The COM object to destroy.</param>
    Sub ReleaseComObject(Of T)(ByRef comObject As T)
      Marshal.FinalReleaseComObject(comObject)
      comObject = Nothing
    End Sub

    ''' <summary>Clean up and quit.</summary>
    ''' <param name="exitCode">The exit code.</param>
    Sub Quit(exitCode As Integer)
      GC.Collect()
      Environment.Exit(exitCode)
    End Sub
  End Module
End Namespace