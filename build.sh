#!/bin/bash
set -e
set -o pipefail

# Detect container runtime
containerRuntime=$(command -v podman || command -v docker)
echo "Using $containerRuntime as container runtime"
if [-z "$containerRuntime"]; then
	echo "Neither Podman nor Docker found!"
	exit 1
fi
echo "Using $containerRuntime as container runtime"

# Create container if it does not exists
$containerRuntime ps -a --format "{{.Names}}" | grep -q "^LegoSetTracker-builder$" || (
	echo "Creating LegoSetTracker-builder container..."
	$containerRuntime compose -f container/compose.yml up -d
	echo "Installing dependencies..."
	$containerRuntime exec -e CI=true LegoSetTracker-builder pnpm install
	if [ $? -ne 0 ]; then
		echo "Failed to install dependencies!"
		exit 1
	fi
)

# Remove old build directories
if [ -d "dist" ]; then
	echo "Removing dist"
	rm -rf dist
fi
if [ -d "builds" ]; then
	echo "Removing builds"
	rm -rf builds
fi

# Start container only if not running
if ! $containerRuntime ps --format "{{.Names}}" | grep -q "^LegoSetTracker-builder$"; then
	echo "Starting LegoSetTracker-builder..."
	$containerRuntime start LegoSetTracker-builder
fi

# Ensure dependencies are installed
echo "Checking and installing dependencies..."
$containerRuntime exec -e CI=true LegoSetTracker-builder pnpm install --frozen-lockfile
if [ $? -ne 0 ]; then
	echo "Failed to install dependencies!"
	exit 1
fi

# Run build
echo "Running electron-builder..."
$containerRuntime exec -e CI=true LegoSetTracker-builder pnpm run build

# Stop container
$containerRuntime stop LegoSetTracker-builder
