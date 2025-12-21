@echo off
@REM Detect container runtime
where podman >nul 2>&1 && set containerRuntime=podman || (
	where docker >nul 2>&1 && set containerRuntime=docker
)
if not defined containerRuntime (
	echo Neither Podman nor Docker found!
	exit /b 1
)
echo Using %containerRuntime% as container runtime

@REM Create container if it does not exists
%containerRuntime% ps -a --format "{{.Names}}" | findstr "LegoSetTracker-builder" >nul || (
	echo Creating LegoSetTracker-builder container...
	%containerRuntime% compose -f container/compose.yml up -d
	echo Installing dependencies...
	%containerRuntime% exec -e CI=true LegoSetTracker-builder pnpm install
	if %ERRORLEVEL% neq 0 (
		echo Failed to install dependencies!
		exit /b 1
	)
)

@REM Remove old build directories
if exist "dist" (
	echo Removing dist
	rmdir /s /q "dist"
)
if exist "builds" (
	echo Removing builds
	rmdir /s /q "builds"
)

@REM Start container if it is not running
%containerRuntime% ps --format "{{.Names}}" | findstr /C:"LegoSetTracker-builder" >nul || (
	 echo Starting LegoSetTracker-builder...
	 %containerRuntime% start LegoSetTracker-builder
)

@REM Ensure dependencies are installed
echo Checking and installing dependencies...
%containerRuntime% exec -e CI=true LegoSetTracker-builder pnpm install --frozen-lockfile
if %ERRORLEVEL% neq 0 (
	echo Failed to install dependencies!
	exit /b 1
)

@REM Run build
echo Running electron-builder...
%containerRuntime% exec -e CI=true LegoSetTracker-builder pnpm run build

@REM Stop container
%containerRuntime% stop LegoSetTracker-builder
