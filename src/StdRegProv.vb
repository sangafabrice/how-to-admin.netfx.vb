''' <summary>StdRegProv WMI class as inspired by mgmclassgen.exe.</summary>
''' <version>0.0.1.1</version>

Imports System.Reflection
Imports System.Diagnostics
Imports System.Management

<Assembly: AssemblyTitle("StdRegProv")>

Namespace ROOT.CIMV2
  Public _
  Module StdRegProv
    Dim CreatedClassName As String = GetTypeName

    Function CheckAccess(hDefKey As UInteger, sSubKeyName As String, ByRef bGranted As Boolean) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        Dim outParams As ManagementBaseObject = .InvokeMethod(methodName, inParams, Nothing)
        bGranted = outParams("bGranted")
        Return outParams("ReturnValue")
      End With
    End Function

    Function CreateKey(hDefKey As UInteger, sSubKeyName As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        Return .InvokeMethod(methodName, inParams, Nothing)("ReturnValue")
      End With
    End Function

    Function DeleteKey(hDefKey As UInteger, sSubKeyName As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        Return .InvokeMethod(methodName, inParams, Nothing)("ReturnValue")
      End With
    End Function

    Function DeleteValue(hDefKey As UInteger, sSubKeyName As String, Optional sValueName As String = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        inParams("sValueName") = sValueName
        Return .InvokeMethod(methodName, inParams, Nothing)("ReturnValue")
      End With
    End Function

    Function EnumKey(hDefKey As UInteger, sSubKeyName As String, ByRef sNames() As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        Dim outParams As ManagementBaseObject = .InvokeMethod(methodName, inParams, Nothing)
        Dim sNamesObj = outParams("sNames")
        If IsArray(sNamesObj) Then
          sNames = sNamesObj
        Else
          sNames = Array.Empty(Of String)()
        End If
        Return outParams("ReturnValue")
      End With
    End Function

    Function GetStringValue(ByRef sValue As String, Optional sSubKeyName As String = Nothing, Optional sValueName As String = Nothing, Optional hDefKey? As UInteger = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        If hDefKey.HasValue Then
          inParams("hDefKey") = hDefKey
        End If
        inParams("sSubKeyName") = sSubKeyName
        inParams("sValueName") = sValueName
        Dim outParams As ManagementBaseObject = .InvokeMethod(methodName, inParams, Nothing)
        sValue = outParams("sValue")
        Return outParams("ReturnValue")
      End With
    End Function

    Function SetStringValue(hDefKey As UInteger, sSubKeyName As String, sValue As String, Optional sValueName As String = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      With New ManagementClass(CreatedClassName)
        Dim inParams As ManagementBaseObject = .GetMethodParameters(methodName)
        inParams("hDefKey") = hDefKey
        inParams("sSubKeyName") = sSubKeyName
        inParams("sValueName") = sValueName
        inParams("sValue") = sValue
        Return .InvokeMethod(methodName, inParams, Nothing)("ReturnValue")
      End With
    End Function

    ''' <summary>Remove the key And all descendant subkeys.</summary>
    Function DeleteKeyTree(hDefKey As UInteger, sSubKeyName As String)
      Dim sNames(0) As String
      Dim returnValue As UInteger = EnumKey(hDefKey, sSubKeyName, sNames)
      If sNames.Length Then
        For Each sName As String In sNames
          returnValue += DeleteKeyTree(hDefKey, sSubKeyName & "\" & sName)
        Next
      End If
      returnValue += DeleteKey(hDefKey, sSubKeyName)
      Return returnValue
    End Function

    Private _
    Function GetMethodName(stackTrace As StackTrace) As String
      Return stackTrace.GetFrame(0).GetMethod.Name
    End Function

    Private _
    Function GetTypeName() As String
      Return (New StackTrace).GetFrame(0).GetMethod.DeclaringType.Name
    End Function
  End Module
End Namespace