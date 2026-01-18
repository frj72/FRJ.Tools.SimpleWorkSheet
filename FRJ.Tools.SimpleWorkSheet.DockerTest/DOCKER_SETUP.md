# Quick Reference: Docker Setup for AutoFit Support

**DISCLAIMER**: This document is AI generated (i.e., written by an idiot) and will probably have terrible text.

This document provides copy-paste ready Docker configurations for using FRJ.Tools.SimpleWorkSheet with AutoFit in production.

## TL;DR - What You Need to Know

1. **dotnet publish /t:PublishContainer DOES NOT WORK** - missing font libraries
2. **You MUST use a custom Dockerfile** - see options below
3. **Both linux-x64 and linux-arm64 are supported** - choose based on your deployment platform
4. **Use appropriate platform flag** when building Docker image (--platform=linux/amd64 or --platform=linux/arm64)

## Option 1: With Microsoft Core Fonts + Aptos (Recommended)

### Fonts Included
- Microsoft Core Fonts: Arial, Arial Black, Andale Mono, Comic Sans MS, Courier New, Georgia, Impact, Times New Roman, Trebuchet MS, Verdana, Webdings (60 font files)
- Aptos Font Family: Aptos (Regular, Bold, Italic, Black, Display, ExtraBold, Light, SemiBold), Aptos Narrow, Aptos Serif, Aptos Mono (28 font files)
- Total: 88 Microsoft font files

### Dockerfile

```dockerfile
FROM mcr.microsoft.com/dotnet/runtime:10.0

ENV DEBIAN_FRONTEND=noninteractive

RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections

RUN apt-get update && apt-get install -y \
    fontconfig \
    fonts-liberation \
    fonts-dejavu-core \
    libfontconfig1 \
    libfreetype6 \
    ttf-mscorefonts-installer \
    curl \
    unzip \
    && rm -rf /var/lib/apt/lists/*

RUN mkdir -p /tmp/aptos && \
    cd /tmp/aptos && \
    curl -L -o aptos.zip 'https://download.microsoft.com/download/8/6/0/860a94fa-7feb-44ef-ac79-c072d9113d69/Microsoft%20Aptos%20Fonts.zip' && \
    unzip -q aptos.zip && \
    mkdir -p /usr/share/fonts/truetype/aptos && \
    cp *.ttf /usr/share/fonts/truetype/aptos/ && \
    cd / && \
    rm -rf /tmp/aptos

RUN fc-cache -fv

WORKDIR /app
COPY publish/ .
ENTRYPOINT ["dotnet", "YourApp.dll"]
```

### Build Commands

```bash
# Publish your .NET app for linux-x64
dotnet publish -c Release --runtime linux-x64 -o ./publish

# Build Docker image (must use linux/amd64 platform)
docker build --platform=linux/amd64 -t your-app:latest .

# Run
docker run --rm your-app:latest
```

### Image Size Impact
- Base runtime: ~200 MB
- Font libraries + Microsoft Core Fonts + Aptos: +41-46 MB
- Total: ~241-246 MB

### When to Use
- Excel files explicitly use Arial, Times New Roman, Courier New, Aptos, etc.
- Maximum font metric accuracy is required
- You need Aptos (the new default Office font)
- Image size is not a critical concern
- You accept the Microsoft Core Fonts EULA and Aptos font license

## Option 2: Without Microsoft Fonts (Smaller)

### Fonts Included
Liberation fonts (Sans, Serif, Mono) + DejaVu fonts (Sans, Sans Mono, Serif)

### Dockerfile

```dockerfile
FROM mcr.microsoft.com/dotnet/runtime:10.0

RUN apt-get update && apt-get install -y \
    fontconfig \
    fonts-liberation \
    fonts-dejavu-core \
    libfontconfig1 \
    libfreetype6 \
    && rm -rf /var/lib/apt/lists/*

RUN fc-cache -fv

WORKDIR /app
COPY publish/ .
ENTRYPOINT ["dotnet", "YourApp.dll"]
```

### Build Commands

**For x64 (AMD64)**:
```bash
# Publish your .NET app for linux-x64
dotnet publish -c Release --runtime linux-x64 -o ./publish

# Build Docker image
docker build --platform=linux/amd64 -t your-app:latest .

# Run
docker run --rm your-app:latest
```

**For ARM64**:
```bash
# Publish your .NET app for linux-arm64
dotnet publish -c Release --runtime linux-arm64 -o ./publish

# Build Docker image
docker build --platform=linux/arm64 -f Dockerfile.arm64 -t your-app:latest .

# Run
docker run --rm your-app:latest
```

### Image Size Impact
- Base runtime: ~200 MB
- Font libraries only: +12.8 MB
- Total: ~213 MB

### Font Fallback Behavior
- Arial → Liberation Sans
- Times New Roman → Liberation Serif
- Courier New → Liberation Mono
- Others → DejaVu Sans

### When to Use
- Smaller Docker image is preferred
- Font fallback with ~10% metric difference is acceptable
- You want to avoid third-party font licensing

## Important Notes

### Platform Requirements
**You MUST**:
1. Publish for matching runtime (`linux-x64` for AMD64, `linux-arm64` for ARM64)
2. Build Docker image with matching platform flag (`--platform=linux/amd64` or `--platform=linux/arm64`)
3. Use custom Dockerfile (PublishContainer alone will not work)

### Platform Support
Both x64 (AMD64) and ARM64 platforms are fully supported. Choose based on your deployment target:
- **x64/AMD64**: Most common, works on standard cloud instances and servers
- **ARM64**: For ARM-based cloud instances (AWS Graviton, Azure ARM VMs, Apple Silicon, etc.)

### Font Licensing
**Microsoft Core Fonts EULA** (Option 1):
- Free to use
- Free to distribute
- Cannot redistribute for profit
- Cannot modify font files

This is acceptable for Docker containers used in application deployment.

## Verification

After building your Docker image, verify AutoFit works:

```bash
# Run your container and generate an Excel file
docker run --rm -v $(pwd)/output:/output your-app:latest

# Open the generated Excel file in Excel/LibreOffice
# Check that columns are properly auto-fitted to content
# Text should not be cut off or have excessive whitespace
```

## Troubleshooting

### Error: "libfontconfig.so.1: cannot open shared object file"
**Cause**: Font libraries not installed in Docker image
**Solution**: Use custom Dockerfile with font packages (see options above)

### Error: "undefined symbol: uuid_generate_random"
**Cause**: This was an issue with older SkiaSharp versions (3.119.1 and later)
**Solution**: Ensure you are using a compatible SkiaSharp version (this library uses 3.116.1)

### Columns are too narrow / text is cut off
**Cause**: AutoFit might be using fallback fonts with different metrics
**Solution**: 
1. Install Microsoft Core Fonts (Option 1)
2. Or use calibration factor (e.g., `.AutoFitAllColumns(1.1)`)

### Font download fails during build
**Cause**: SourceForge might be slow or unreachable
**Solution**: 
1. Retry the build
2. Or use Option 2 (without MS fonts)
3. Or cache the font installer locally

## Quick Start Template

Copy this to your project root:

**Dockerfile**:
```dockerfile
FROM mcr.microsoft.com/dotnet/runtime:10.0
ENV DEBIAN_FRONTEND=noninteractive
RUN echo "ttf-mscorefonts-installer msttcorefonts/accepted-mscorefonts-eula select true" | debconf-set-selections
RUN apt-get update && apt-get install -y fontconfig fonts-liberation fonts-dejavu-core libfontconfig1 libfreetype6 ttf-mscorefonts-installer && rm -rf /var/lib/apt/lists/*
RUN fc-cache -fv
WORKDIR /app
COPY publish/ .
ENTRYPOINT ["dotnet", "YourApp.dll"]
```

**Build script for x64** (`build-docker-x64.sh`):
```bash
#!/bin/bash
dotnet publish -c Release --runtime linux-x64 -o ./publish
docker build --platform=linux/amd64 -t your-app:latest .
```

**Build script for ARM64** (`build-docker-arm64.sh`):
```bash
#!/bin/bash
dotnet publish -c Release --runtime linux-arm64 -o ./publish
docker build --platform=linux/arm64 -f Dockerfile.arm64 -t your-app:latest .
```

Make executable: `chmod +x build-docker-*.sh`

Run: `./build-docker-x64.sh` or `./build-docker-arm64.sh`
