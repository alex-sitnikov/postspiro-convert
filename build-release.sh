#!/bin/bash

echo "Building SpiroUI for production..."

# Build Tailwind CSS for production (minified)
echo "Building Tailwind CSS..."
npm run css:prod

# Clean previous builds
echo "Cleaning previous builds..."
dotnet clean

# Build the Blazor WebAssembly app in Release mode
echo "Building Blazor WebAssembly app..."
dotnet publish -c Release -o dist

echo "Build complete! Output is in the 'dist' directory."
echo "Deploy the contents of 'dist/wwwroot' to your web server."