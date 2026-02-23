# Building Windows Executable with Docker

## Overview

You can build a Windows executable (`.exe`) using Docker without needing Windows or PyInstaller installed locally. This works on macOS, Linux, or Windows.

## Prerequisites

- Docker installed on your system
  - [Docker Desktop for Mac](https://docs.docker.com/desktop/install/mac-install/)
  - [Docker Desktop for Linux](https://docs.docker.com/desktop/install/linux-install/)
  - [Docker Desktop for Windows](https://docs.docker.com/desktop/install/windows-install/)

## Building the Executable

### Option 1: Using Docker Directly

```bash
cd /path/to/report_card_generator
docker build -f Dockerfile.windows -t report-card-builder .
docker run --rm -v $(pwd)/dist:/app/dist report-card-builder
```

On Windows CMD:

```cmd
docker build -f Dockerfile.windows -t report-card-builder .
docker run --rm -v %cd%\dist:/app/dist report-card-builder
```

On Windows PowerShell:

```powershell
docker build -f Dockerfile.windows -t report-card-builder .
docker run --rm -v ${PWD}/dist:/app/dist report-card-builder
```

### Option 2: Using Docker Compose

```bash
docker-compose -f docker-compose.windows.yml build
docker-compose -f docker-compose.windows.yml run --rm windows-builder
```

## Output

After the build completes, you'll find your executable at:

```
dist/Report_Card_Generator.exe
```

## What Happens

1. Docker creates a container with Python 3.9 and required dependencies
2. PyInstaller packages your Python code into a Windows executable
3. The `.exe` file is saved to the `dist/` folder
4. You can now run this `.exe` on any Windows machine without Python installed!

## Distribution

To use the executable on Windows:

1. Copy `dist/Report_Card_Generator.exe` to your Windows machine
2. Double-click it to run
3. No Python installation required!

## Troubleshooting

### Build fails with permission denied

Make sure Docker daemon is running and you have proper permissions.

### Can't find dist folder

The dist folder is created during the build. Make sure you're running the command in the project root directory.

### Executable doesn't work on Windows

- Ensure you're using the exact `.exe` file from the Docker build output
- Try running from command line to see any error messages: `Report_Card_Generator.exe`

## Benefits of Using Docker

✅ No need for Python installed locally  
✅ No need for PyInstaller installed locally  
✅ Same build on Mac, Linux, or Windows  
✅ Repeatable, consistent builds  
✅ No environment conflicts

---

For traditional Windows build instructions (without Docker), see [WINDOWS_BUILD_INSTRUCTIONS.md](WINDOWS_BUILD_INSTRUCTIONS.md)
