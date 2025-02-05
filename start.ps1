param (
	[string]$EnvFilePath = ".env",
	[string]$Configuration = "Release"
)

if (!(Test-Path $EnvFilePath -PathType Leaf)) {
	Write-Error "The required .env file was not found at '$EnvFilePath'. Make sure to follow the guidance in the README"
	exit 1
}
if (!(Get-Command "msbuild" -ErrorAction SilentlyContinue)) {
	Write-Error "msbuild must be installed and in your PATH for this script to work properly"
	exit 1
}

Write-Host "Restoring dependencies..."
dotnet restore
if ($LASTEXITCODE -ne 0) {
	Write-Error "Failed to restore Autogrator dependencies"
	exit $LASTEXITCODE
}

Write-Host "Building Autogrator with configuration '$Configuration'..."
msbuild /t:Build /p:Configuration=$Configuration
if ($LASTEXITCODE -ne 0) {
	Write-Error "Failed to build Autogrator"
	exit $LASTEXITCODE
}

$executablePath = ".\Autogrator\bin\$Configuration\net9.0\Autogrator.exe"
if (!(Test-Path $executablePath)) {
	Write-Error "The Autogrator executable was not found at $executablePath"
	exit 1
}

Write-Host "Running Autogrator..."
Start-Process $executablePath