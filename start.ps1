if (!(Test-Path ".env" -PathType Leaf)) {
	Write-Error ".env does not exist and is required. Make sure to follow the guidance in the README"
	exit 1
}
if (!(Get-Command "msbuild" -ErrorAction SilentlyContinue)) {
	Write-Error "msbuild is required for this script"
	exit 1
}

dotnet restore
msbuild /t:Build /p:Configuration=Release
start ".\Autogrator\bin\Release\net9.0\Autogrator.exe"