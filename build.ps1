<#PSScriptInfo .VERSION 1.0.0#>

using namespace System.Management.Automation
[CmdletBinding()]
Param ()

& "$PSScriptRoot\package.ps1"
& "$PSScriptRoot\clean.ps1"

& {
  $HostColorArgs = @{
    ForegroundColor = 'Black'
    BackgroundColor = 'Green'
    NoNewline = $True
  }

  Try {
    Remove-Item ($BinDir = "$PSScriptRoot\bin") -Recurse -ErrorAction Stop
  } Catch [ItemNotFoundException] {
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
  } Catch {
    $HostColorArgs.BackgroundColor = 'Red'
    Write-Host $_.Exception.Message @HostColorArgs
    Write-Host
    Return
  }
  New-Item $BinDir -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
  Copy-Item "$PSScriptRoot\rsc" -Destination $BinDir -Recurse
  Copy-Item "$PSScriptRoot\lib\*" -Destination $BinDir

  # Compile the source code with vbc.exe.
  $EnvPath = $Env:Path
  $Env:Path = "$Env:windir\Microsoft.NET\Framework$(If ([Environment]::Is64BitOperatingSystem) { '64' })\v4.0.30319\;$Env:Path"
  vbc.exe /nologo /target:$($DebugPreference -eq 'Continue' ? 'exe':'winexe') /win32icon:"$PSScriptRoot\menu.ico" /reference:"$BinDir\Interop.IWshRuntimeLibrary.dll" /reference:"$BinDir\PresentationFramework.dll" /reference:"$BinDir\PresentationCore.dll" /reference:"$BinDir\WindowsBase.dll" /reference:System.Xaml.dll /out:$(($ConvertExe = "$BinDir\cvmd2html.exe")) "$(($SrcDir = "$PSScriptRoot\src"))\AssemblyInfo.vb" "$SrcDir\ErrorLog.vb" "$SrcDir\Package.vb" "$SrcDir\Parameters.vb" "$PSScriptRoot\Program.vb" "$SrcDir\Setup.vb" "$SrcDir\Util.vb"
  $Env:Path = $EnvPath

  If ($LASTEXITCODE -eq 0) {
    Write-Host "Output file $ConvertExe written." @HostColorArgs
    (Get-Item $ConvertExe).VersionInfo | Format-List -Property @{
      Name = "AssemblyName";
      Expression = { $_.OriginalFilename }
    },@{
      Name = "AssemblyProduct";
      Expression = { $_.ProductName }
    },@{
      Name = "AssemblyTitle";
      Expression = { $_.FileDescription }
    },@{
      Name = "AssemblyCompany";
      Expression = { $_.CompanyName }
    },@{
      Name = "AssemblyFileVersion";
      Expression = { $_.FileVersion }
    },@{
      Name = "AssemblyProductVersion";
      Expression = { $_.ProductVersion }
    },@{
      Name = "AssemblyCopyright";
      Expression = { $_.LegalCopyright }
    } -Force
  }
}