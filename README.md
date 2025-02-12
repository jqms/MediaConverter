# Format Converter

## 🎬 Quick Media Conversion Utility for Windows

### Overview

Format Converter is a lightweight Windows utility that seamlessly adds context menu options for converting media files between various formats. With just a right-click, you can quickly transform your audio and video files.

### 🌟 Features

- **Easy Conversion**: Convert files directly from the context menu
- **Supported Formats**:
  - Video: MP4, MOV
  - Audio: MP3, OGG, OPUS, WAV
- **One-Click Installation**: Automatically sets up right-click context menus
- **Administrator Privileges**: Ensures smooth system integration
- **User-Friendly**: Simple progress indicator during conversion

### 🛠️ Requirements

- Windows Operating System
- FFmpeg installed and accessible in system PATH
- .NET Framework

### 💾 Installation

1. Download the executable
2. Run the application once to register context menus
3. Right-click any supported file to use "Convert To" option

### 🚀 Usage

#### Context Menu Conversion
1. Right-click on a supported media file
2. Select "Convert To"
3. Choose the desired output format
4. The converted file will be automatically saved and selected in File Explorer

#### Command Line Usage
```
FormatConverter.exe [input_file] [output_format]
```

#### Uninstallation
Run the application with `-unregister` flag to remove context menus

### 🔄 Conversion Capabilities

| Input Type | Possible Conversions |
|-----------|----------------------|
| MP4       | MP3, MOV             |
| MOV       | MP3, MP4             |
| MP3       | MP4, OGG, OPUS, WAV  |
| OGG       | MP3, MP4, OPUS, WAV  |
| OPUS      | MP3, MP4, OGG, WAV   |
| WAV       | MP3, MP4, OGG, OPUS  |

### ⚠️ Notes

- Requires administrator privileges for installation
- Utilizes FFmpeg for robust media conversion
- Audio-to-video conversions create a black background
