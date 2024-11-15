<#PSScriptInfo .VERSION 1.0.3#>

using namespace System.Management.Automation
[CmdletBinding()]
param ()

& {
  Import-Module "$PSScriptRoot\tools"
  Format-ProjectCode @('*.vb','*.ps*1','.gitignore'| ForEach-Object { "$PSScriptRoot\$_" })

  $HostColorArgs = @{
    ForegroundColor = 'Black'
    BackgroundColor = 'Green'
  }

  try { Remove-Item "$(($BinDir = "$PSScriptRoot\bin"))\*" -Recurse -ErrorAction Stop }
  catch [ItemNotFoundException] { Write-Host $_.Exception.Message @HostColorArgs }
  catch {
    $HostColorArgs.BackgroundColor = 'Red'
    Write-Host $_.Exception.Message @HostColorArgs
    return
  }
  New-Item $BinDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
  Copy-Item "$PSScriptRoot\rsc" -Destination $BinDir -Recurse
  Save-ProjectPackage $PSScriptRoot
  Set-ProjectVersion $PSScriptRoot

  # Compile the source code with vbc.exe.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:ProgramFiles\Microsoft Visual Studio\2022\Community\Msbuild\Current\Bin\Roslyn\;$Env:Path"
  vbc.exe /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /win32icon:"$PSScriptRoot\menu.ico" /win32manifest:"$PSScriptRoot\cvmd2html.manifest" /reference:"$BinDir\Interop.IWshRuntimeLibrary.dll" /reference:$(Get-WpfLibrary PresentationFramework) /reference:$(Get-WpfLibrary PresentationCore) /reference:$(Get-WpfLibrary WindowsBase) /reference:System.Xaml.dll /out:$(($ConvertExe = "$BinDir\cvmd2html.exe")) "$(($SrcDir = "$PSScriptRoot\src"))\AssemblyInfo.vb" "$SrcDir\ErrorLog.vb" "$SrcDir\Package.vb" "$SrcDir\Parameters.vb" "$PSScriptRoot\Program.vb" "$SrcDir\Setup.vb" "$SrcDir\Util.vb"
  $Env:Path = $EnvPath

  if ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs -NoNewline
    Format-ExecutableInfo $ConvertExe
  }

  Remove-Module tools
}