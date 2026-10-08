$ErrorActionPreference = 'Stop'
$compiler = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
& $compiler /nologo /target:winexe /platform:x64 /optimize+ /utf8output /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:Microsoft.CSharp.dll /out:"$PSScriptRoot\Holography40x64.exe" "$PSScriptRoot\Preview.cs" "$PSScriptRoot\Transport.cs"
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }
