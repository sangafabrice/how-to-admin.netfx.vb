''' <summary>The parsed parameters.</summary>
''' <version>0.0.1.0</version>

Imports System.Dynamic
Imports System.Text

Namespace cvmd2html
  Module Parameters
    ''' <summary>The parameter object.</summary>
    ReadOnly param As Object = ParseCommandLine(Environment.GetCommandLineArgs.Skip(1).ToArray)

    ''' <summary>The input markdown path.</summary>
    Friend ReadOnly MarkdownPathParam As String = param.MarkdownPath

    ''' <summary>Specify to configure the shortcut in the registry.</summary>
    Friend ReadOnly SetConfigParam As Boolean = param.SetConfig

    ''' <summary>Specify to remove the shortcut menu.</summary>
    Friend ReadOnly UnsetConfigParam As Boolean = param.UnsetConfig

    ''' <summary>Specify to configure the shortcut without the icon.</summary>
    Friend ReadOnly NoIconConfigParam As Boolean? = param.NoIconConfig

    Private _
    Function ParseCommandLine(args() As String) As Object
      Dim paramExpando As Object = New ExpandoObject
      With paramExpando
        .MarkdownPath = Nothing
        .SetConfig = False
        .NoIconConfig = Nothing
        .UnsetConfig = False
      End With
      If UBound(args) = 0 Then
        Dim arg As String = Trim(args(0))
        Dim paramNameValue() As String = Split(arg, ":", 2)
        If UBound(paramNameValue) = 1 And StrComp(paramNameValue(0), "/Markdown", CompareMethod.Text) = 0 Then
          Dim paramMarkdown As String = paramNameValue(1)
          If Len(paramMarkdown) Then
            paramExpando.MarkdownPath = paramMarkdown
            Return paramExpando
          End If
        End If
        Select Case LCase(arg)
          Case "/set"
            With paramExpando
              .SetConfig = True
              .NoIconConfig = False
            End With
            Return paramExpando
          Case "/set:noicon"
            With paramExpando
              .SetConfig = True
              .NoIconConfig = True
            End With
            Return paramExpando
          Case "/unset"
            paramExpando.UnsetConfig = True
            Return paramExpando
          Case Else
            paramExpando.MarkdownPath = arg
            Return paramExpando
        End Select
      ElseIf args.Length = 0 Then
        With paramExpando
          .SetConfig = True
          .NoIconConfig = False
        End With
        Return paramExpando
      End If
      ShowHelp
      Return Nothing
    End Function

    Private _
    Sub ShowHelp
      With New StringBuilder
        .AppendLine("The MarkdownToHtml shortcut launcher.")
        .AppendLine("It starts the shortcut menu target script in a hidden window.")
        .AppendLine
        .AppendLine("Syntax:")
        .AppendLine("  Convert-MarkdownToHtml [/Markdown:]<markdown file path>")
        .AppendLine("  Convert-MarkdownToHtml [/Set[:NoIcon]]")
        .AppendLine("  Convert-MarkdownToHtml /Unset")
        .AppendLine("  Convert-MarkdownToHtml /Help")
        .AppendLine
        .AppendLine("<markdown file path>  The selected markdown's file path.")
        .AppendLine("                 Set  Configure the shortcut menu in the registry.")
        .AppendLine("              NoIcon  Specifies that the icon is not configured.")
        .AppendLine("               Unset  Removes the shortcut menu.")
        .AppendLine("                Help  Show the help doc.")
        Popup(.ToString)
      End With
      Quit(1)
    End Sub
  End Module
End Namespace