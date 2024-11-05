<#PSScriptInfo .VERSION 1.0.2#>

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

  $AssemblyInfoVb = "$(($SrcDir = "$PSScriptRoot\src"))\AssemblyInfo.vb"

  function ImportMgmtClass([string] $ClassName) {
    $FileName = $ClassName.Replace('_', '.')
    vbc.exe /nologo /target:library /out:$(($ClassDll = "$BinDir\$FileName.dll")) /reference:"$BinDir\Interop.WbemScripting.dll" $AssemblyInfoVb "$SrcDir\$FileName.vb"
    return $ClassDll
  }

  # Compile the source code with vbc.exe.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:windir\Microsoft.NET\Framework$(If ([Environment]::Is64BitOperatingSystem) { '64' })\v4.0.30319\;$Env:Path"
  vbc.exe /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /win32icon:"$PSScriptRoot\menu.ico" /reference:$(ImportMgmtClass StdRegProv) /reference:"$BinDir\Interop.WbemScripting.dll" /reference:"$BinDir\Interop.IWshRuntimeLibrary.dll" /reference:$(Get-WpfLibrary PresentationFramework) /reference:$(Get-WpfLibrary PresentationCore) /reference:$(Get-WpfLibrary WindowsBase) /reference:System.Xaml.dll /out:$(($ConvertExe = "$BinDir\cvmd2html.exe")) $AssemblyInfoVb "$SrcDir\ErrorLog.vb" "$SrcDir\Package.vb" "$SrcDir\Parameters.vb" "$PSScriptRoot\Program.vb" "$SrcDir\Setup.vb" "$SrcDir\Util.vb"
  $Env:Path = $EnvPath

  if ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs -NoNewline
    Format-ExecutableInfo $ConvertExe
  }

  Remove-Module tools
}