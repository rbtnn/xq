@echo off
rem Build xq.exe with the in-box C# 5 compiler (no SDK, no Roslyn needed).
setlocal
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" set CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe

"%CSC%" /nologo /optimize+ /target:exe /platform:anycpu ^
  /out:xq.exe ^
  /r:System.dll /r:System.Core.dll /r:System.Xml.dll ^
  /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll ^
  src\A1.cs src\Cfb.cs src\Ovba.cs src\Xlsx.cs src\NumFmt.cs src\IBook.cs src\Ptg.cs src\Ptg12.cs src\Biff8.cs src\XlsbBook.cs src\xq.cs

if errorlevel 1 exit /b 1
echo built xq.exe
