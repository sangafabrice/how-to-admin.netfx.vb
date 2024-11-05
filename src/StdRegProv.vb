''' <summary>StdRegProv WMI class as inspired by mgmclassgen.exe.</summary>
''' <version>0.0.1.0</version>

Imports WbemScripting
Imports System.Reflection
Imports System.Diagnostics
Imports System.Runtime.InteropServices

<Assembly: AssemblyTitle("StdRegProv")>

Namespace ROOT.CIMV2
  Public _
  Module StdRegProv
    Dim CreatedClassName As String = GetTypeName

    Function CheckAccess(hDefKey As UInteger, sSubKeyName As String, ByRef bGranted As Boolean) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          Dim outParams As SWbemObject = classObj.ExecMethod_(methodName, inParams)
          bGranted = GetParamProperty(Of Boolean)(outParams, "bGranted")
          Try
            Return GetReturnValue(outParams)
          Finally
            ReleaseComObject(outParams)
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function CreateKey(hDefKey As UInteger, sSubKeyName As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          Try
            Return GetReturnValue(classObj.ExecMethod_(methodName, inParams))
          Finally
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function DeleteKey(hDefKey As UInteger, sSubKeyName As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          Try
            Return GetReturnValue(classObj.ExecMethod_(methodName, inParams))
          Finally
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function DeleteValue(hDefKey As UInteger, sSubKeyName As String, Optional sValueName As String = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          SetParamProperty(inParams, "sValueName", sValueName)
          Try
            Return GetReturnValue(classObj.ExecMethod_(methodName, inParams))
          Finally
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function EnumKey(hDefKey As UInteger, sSubKeyName As String, ByRef sNames() As String) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          Dim outParams As SWbemObject = classObj.ExecMethod_(methodName, inParams)
          Dim sNamesObj = GetParamProperty(Of Object)(outParams, "sNames")
          Try
            If TypeOf sNamesObj Is DBNull Then
              sNames = Array.Empty(Of String)()
            Else
              Dim cvobj2str As Converter(Of Object, String) =
                Function(item)
                  Return item.ToString
                End Function
              sNames = Array.ConvertAll(sNamesObj, cvobj2str)
            End If
            Return GetReturnValue(outParams)
          Finally
            ReleaseComObject(outParams)
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function GetStringValue(ByRef sValue As String, Optional sSubKeyName As String = Nothing, Optional sValueName As String = Nothing, Optional hDefKey? As UInteger = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          If hDefKey.HasValue Then
            SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          End If
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          SetParamProperty(inParams, "sValueName", sValueName)
          Dim outParams As SWbemObject = classObj.ExecMethod_(methodName, inParams)
          sValue = GetParamProperty(Of String)(outParams, "sValue")
          Try
            Return GetReturnValue(outParams)
          Finally
            ReleaseComObject(outParams)
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
    End Function

    Function SetStringValue(hDefKey As UInteger, sSubKeyName As String, sValue As String, Optional sValueName As String = Nothing) As UInteger
      Dim methodName As String = GetMethodName(New StackTrace)
      Using Registry = New StdRegProvider(CreatedClassName)
        Dim classObj As SWbemObject = Registry.Provider
        Using Input = New StdRegInput(classObj, methodName)
          Dim inParams As SWbemObject = Input.Params
          SetParamProperty(inParams, "hDefKey", ConvertToInt32(hDefKey))
          SetParamProperty(inParams, "sSubKeyName", sSubKeyName)
          SetParamProperty(inParams, "sValueName", sValueName)
          SetParamProperty(inParams, "sValue", sValue)
          Try
            Return GetReturnValue(classObj.ExecMethod_(methodName, inParams))
          Finally
            ReleaseComObject(inParams)
            ReleaseComObject(classObj)
          End Try
        End Using
      End Using
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
    Class StdRegInput
      Implements IDisposable
      Dim wbemMethodSet As SWbemMethodSet
      Dim wbemMethod As SWbemMethod
      Dim inParamsClass As SWbemObject
      Friend Params As SWbemObject

      Friend _
      Sub New(ByRef classObj As SWbemObject, methodName As String)
        wbemMethodSet = classObj.Methods_
        wbemMethod = wbemMethodSet.Item(methodName)
        inParamsClass = wbemMethod.InParameters
        Params = inParamsClass.SpawnInstance_
      End Sub

      Private _
      Sub IDisposable_Dispose() Implements IDisposable.Dispose
        ReleaseComObject(Params)
        ReleaseComObject(inParamsClass)
        ReleaseComObject(wbemMethod)
        ReleaseComObject(wbemMethodSet)
      End Sub
    End Class

    Private _
    Class StdRegProvider
      Implements IDisposable
      Dim wbemLocator = New SWbemLocator
      Dim wbemService As SWbemServices
      Friend Provider As SWbemObject

      Friend _
      Sub New(className As String)
        wbemService = wbemLocator.ConnectServer
        Provider = wbemService.Get(className)
      End Sub

      Private _
      Sub IDisposable_Dispose() Implements IDisposable.Dispose
        ReleaseComObject(Provider)
        ReleaseComObject(wbemService)
        ReleaseComObject(wbemLocator)
      End Sub
    End Class

    ''' <summary>Release the specified COM object.</summary>
    ''' <param name="comObject">The COM object to destroy.</param>
    Private _
    Sub ReleaseComObject(Of T)(ByRef comObject As T)
      Marshal.FinalReleaseComObject(comObject)
      comObject = Nothing
    End Sub

    Private _
    Sub SetParamProperty(Of T)(ByRef inParams As SWbemObject, propertyName As String, propertyValue As T)
      Dim propertySet As SWbemPropertySet = inParams.Properties_
      Dim propertyObj As SWbemProperty = propertySet.Item(propertyName)
      propertyObj.Value = propertyValue
      ReleaseComObject(propertyObj)
      ReleaseComObject(propertySet)
    End Sub

    Private _
    Function GetParamProperty(Of T)(ByRef outParams As SWbemObject, propertyName As String) As T
      Dim propertySet As SWbemPropertySet = outParams.Properties_
      Dim propertyObj As SWbemProperty = propertySet.Item(propertyName)
      GetParamProperty = propertyObj.Value
      ReleaseComObject(propertyObj)
      ReleaseComObject(propertySet)
    End Function

    Private _
    Function GetReturnValue(outParams As SWbemObject) As UInteger
      GetReturnValue = GetParamProperty(Of UInteger)(outParams, "ReturnValue")
      Marshal.ReleaseComObject(outParams)
      outParams = Nothing
    End Function

    Private _
    Function ConvertToInt32(hDefKey As UInteger) As Integer
      Return BitConverter.ToInt32(BitConverter.GetBytes(hDefKey), 0)
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