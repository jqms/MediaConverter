# Format Converter
## 🎮 Quick Media Conversion Utility for Windows

### Overview
Format Converter is a lightweight Windows utility that seamlessly adds context menu options for converting media files between various formats. With just a right-click, you can quickly transform your audio, video, and image files.

### 🌟 Features
- **Easy Conversion**: Convert files directly from the context menu
- **Supported Formats**:
  - Video: MP4, MKV, AVI, MOV, WMV, WEBM, FLV, TS, 3GP
  - Audio: MP3, WAV, OGG, OPUS, AAC, FLAC, WMA, M4A, AC3
  - Image: JPG/JPEG, PNG, BMP, GIF, WEBP, TIFF, ICO, HEIC
- **Additional Tools**:
  - Mute audio/video files
  - Resize images (50% and 75%)
  - **Compress audio/video files to specific sizes** (10MB, 20MB, 50MB, 100MB, 500MB)
- **One-Click Installation**: Automatically sets up right-click context menus
- **Administrator Privileges**: Ensures smooth system integration
- **User-Friendly**: Detailed progress indicator during conversion

### 🧭 Requirements
- Windows Operating System
- FFmpeg installed and accessible in system PATH
- .NET Framework

### 💽 Getting FFmpeg
Format Converter needs a helper program called FFmpeg to convert your files. Here's how to install it:

#### Easiest Method: Install with Winget (Windows Package Manager)

1. **Install FFmpeg using Winget**:
   - Press the Windows key, type "Command Prompt" or "PowerShell" and click on it to open
   - Copy and paste this line, then press Enter:
   ```
   winget install ffmpeg
   ```
   - Wait for it to finish downloading and installing
   
2. **You're Done!** Format Converter will now be able to find FFmpeg automatically.

#### Manual Installation Method
If you prefer to install FFmpeg manually:

1. **Download the FFmpeg Program**:
   - Go to [gyan.dev/ffmpeg/builds](https://www.gyan.dev/ffmpeg/builds/)
   - Click on "ffmpeg-release-essentials.zip" to download

2. **Extract the Files**:
   - Find the downloaded ZIP file in your Downloads folder
   - Right-click on it and select "Extract All..."
   - Choose a location that's easy to remember, like "C:\\ffmpeg"
   - Click "Extract"

3. **Tell Windows Where to Find FFmpeg**:
   - Right-click on the Start button and select "System"
   - Click on "Advanced system settings" on the right
   - At the bottom of the new window, click "Environment Variables"
   - In the "System variables" box, find and click on "Path", then click "Edit"
   - Click "New" and type in the path to the bin folder (e.g., "C:\\ffmpeg\\bin")
   - Click "OK" on all open windows

4. **Check if It Worked**:
   - Press the Windows key, type "cmd" and click on Command Prompt
   - Type `ffmpeg -version` and press Enter
   - If you see information about FFmpeg (not an error), it's working!

### 💾 Installation
1. Download the executable
2. Run the application once to register context menus
3. Right-click any supported file to use "Convert To" option

### 🚀 Usage
#### Context Menu Conversion
1. Right-click on a supported media file
2. Select "Convert To"
3. Choose the desired output format or tool
4. The converted file will be automatically saved and selected in File Explorer

#### Command Line Usage
```
FormatConverter.exe [input_file] [output_format]
```

#### Batch Conversion
```
FormatConverter.exe [input_file1] [input_file2] [...] [output_format]
```

#### Special Operations
```
FormatConverter.exe [input_file] mute
FormatConverter.exe [input_file] resize50
FormatConverter.exe [input_file] resize75
FormatConverter.exe [input_file] compress 10
FormatConverter.exe [input_file] compress 20
FormatConverter.exe [input_file] compress 50
FormatConverter.exe [input_file] compress 100
FormatConverter.exe [input_file] compress 500
```

#### Uninstallation
Run the application with `-unregister` flag to remove context menus:
```
FormatConverter.exe -unregister
```

### 🔄 Conversion Capabilities

| Input Type | Possible Conversions |
|------------|----------------------|
| Video Files |  |
| MP4        | MKV, AVI, MOV, WMV, WEBM, FLV, TS, 3GP, MP3, M4A |
| MKV        | MP4, AVI, MOV, WMV, WEBM, FLV, TS, 3GP, MP3, M4A |
| AVI        | MP4, MKV, MOV, WMV, WEBM, FLV, TS, 3GP, MP3, M4A |
| MOV        | MP4, MKV, AVI, WMV, WEBM, FLV, TS, 3GP, MP3, M4A |
| WMV        | MP4, MKV, AVI, MOV, WEBM, FLV, TS, 3GP, MP3, M4A |
| WEBM       | MP4, MKV, AVI, MOV, WMV, FLV, TS, 3GP, MP3, M4A |
| FLV        | MP4, MKV, AVI, MOV, WMV, WEBM, TS, 3GP, MP3, M4A |
| TS         | MP4, MKV, AVI, MOV, WMV, WEBM, FLV, 3GP, MP3, M4A |
| 3GP        | MP4, MKV, AVI, MOV, WMV, WEBM, FLV, TS, MP3, M4A |
| Audio Files |  |
| MP3        | WAV, OGG, OPUS, AAC, FLAC, WMA, M4A, AC3 |
| WAV        | MP3, OGG, OPUS, AAC, FLAC, WMA, M4A, AC3 |
| OGG        | MP3, WAV, OPUS, AAC, FLAC, WMA, M4A, AC3 |
| OPUS       | MP3, WAV, OGG, AAC, FLAC, WMA, M4A, AC3 |
| AAC        | MP3, WAV, OGG, OPUS, FLAC, WMA, M4A, AC3 |
| FLAC       | MP3, WAV, OGG, OPUS, AAC, WMA, M4A, AC3 |
| WMA        | MP3, WAV, OGG, OPUS, AAC, FLAC, M4A, AC3 |
| M4A        | MP3, WAV, OGG, OPUS, AAC, FLAC, WMA, AC3 |
| AC3        | MP3, WAV, OGG, OPUS, AAC, FLAC, WMA, M4A |
| Image Files |  |
| JPG/JPEG   | PNG, BMP, GIF, WEBP, TIFF, ICO, HEIC |
| PNG        | JPG, BMP, GIF, WEBP, TIFF, ICO, HEIC |
| BMP        | JPG, PNG, GIF, WEBP, TIFF, ICO, HEIC |
| GIF        | JPG, PNG, BMP, WEBP, TIFF, ICO, HEIC |
| WEBP       | JPG, PNG, BMP, GIF, TIFF, ICO, HEIC |
| TIFF       | JPG, PNG, BMP, GIF, WEBP, ICO, HEIC |
| ICO        | JPG, PNG, BMP, GIF, WEBP, TIFF, HEIC |
| HEIC       | JPG, PNG, BMP, GIF, WEBP, TIFF, ICO |

### ⚠️ Notes
- Requires administrator privileges for installation
- Utilizes FFmpeg for robust media conversion
- Audio-to-video conversions create a black background
- Images can be resized to 50% or 75% of original dimensions
- Audio and video files can be muted (creates a new file with no sound)
- **New Feature**: Audio and video files can be compressed to specific file sizes (10MB, 20MB, 50MB, 100MB, 500MB)
- Converted files are saved in the same directory as the original with the new extension
- Batch conversion supports multiple input files with a single output format

### ❓ Troubleshooting
- **"FFmpeg is not recognized" error**: Make sure FFmpeg is correctly installed and added to your system PATH
- **Conversion fails**: Check if FFmpeg supports the specific conversion between your chosen formats
- **Missing context menu**: Try running the application again with administrator privileges
- **Slow conversion**: Some formats (like WEBM) may take longer to process due to compression requirements

