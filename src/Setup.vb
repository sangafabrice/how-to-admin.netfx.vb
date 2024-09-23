''' <summary>The methods for managing the shortcut menu option: install and uninstall.</summary>
''' <version>0.0.1.0</version>

Imports Win32 = Microsoft.Win32

Namespace cvmd2html
  Module Setup
    Const SHELL_SUBKEY = "SOFTWARE\Classes\SystemFileAssociations\.md\shell"
    Const VERB = "cthtml"
    ReadOnly VERB_SUBKEY As String = String.Format("{0}\{1}", SHELL_SUBKEY, VERB)
    ReadOnly VERB_KEY As String = String.Format("{0}\{1}", Win32.Registry.CurrentUser, VERB_SUBKEY)
    Const ICON_VALUENAME = "Icon"

    ''' <summary>Configure the shortcut menu in the registry.</summary>
    Sub SetShortcut()
      Dim COMMAND_KEY As String = VERB_KEY & "\command"
      Dim command As String = String.Format("""{0}"" /Markdown:""%1""", AssemblyLocation)
      Win32.Registry.SetValue(COMMAND_KEY, Nothing, command)
      Win32.Registry.SetValue(VERB_KEY, Nothing, "Convert to &HTML")
    End Sub

    ''' <summary>Add an icon to the shortcut menu in the registry.</summary>
    Sub AddIcon()
      Win32.Registry.SetValue(VERB_KEY, ICON_VALUENAME, AssemblyLocation)
    End Sub

    ''' <summary>Remove the shortcut icon menu.</summary>
    Sub RemoveIcon()
      Using VERB_KEY_OBJ As Win32.RegistryKey = Win32.Registry.CurrentUser.OpenSubKey(VERB_SUBKEY, True)
        If VERB_KEY_OBJ IsNot Nothing Then VERB_KEY_OBJ.DeleteValue(ICON_VALUENAME, False)
      End Using
    End Sub

    ''' <summary>Remove the shortcut menu by removing the verb key and subkeys.</summary>
    Sub UnsetShortcut()
      Using SHELL_KEY_OBJ As Win32.RegistryKey = Win32.Registry.CurrentUser.OpenSubKey(SHELL_SUBKEY, True)
        If SHELL_KEY_OBJ IsNot Nothing Then SHELL_KEY_OBJ.DeleteSubKeyTree(VERB, False)
      End Using
    End Sub
  End Module
End Namespace