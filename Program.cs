using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;
using System.Management;

namespace FormatConverter {
    class Program {
        [DllImport("kernel32.dll")]
        private static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const uint CREATE_NO_WINDOW = 0x08000000;

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SHParseDisplayName(string name, IntPtr pbc, out IntPtr pidl, uint sfgaoIn,
            out uint psfgaoOut);

        [DllImport("shell32.dll")]
        public static extern int SHOpenFolderAndSelectItems(IntPtr pidlFolder, uint cidl, [In] IntPtr[] apidl,
            uint dwFlags);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr ILFindChild(IntPtr pidlParent, IntPtr pidlChild);

        [DllImport("ole32.dll")]
        public static extern void CoTaskMemFree(IntPtr pv);

        private static readonly string[] VideoConversions =
            { "mp4", "mkv", "avi", "mov", "wmv", "webm", "flv", "mp3", "m4a" };

        private static readonly string[] AudioConversions =
            { "mp3", "wav", "ogg", "opus", "aac", "flac", "m4a", "wma" };

        private static readonly string[] ImageConversions =
            { "jpg", "jpeg", "png", "bmp", "gif", "webp", "tiff", "ico" };

        private static readonly Dictionary<string, (string type, string[] conversions)> FormatMappings = new()
        {
            { ".mp4", ("video", VideoConversions) },
            { ".mkv", ("video", VideoConversions) },
            { ".avi", ("video", VideoConversions) },
            { ".mov", ("video", VideoConversions) },
            { ".wmv", ("video", VideoConversions) },
            { ".webm", ("video", VideoConversions) },
            { ".flv", ("video", VideoConversions) },
            { ".ts", ("video", VideoConversions) },
            { ".3gp", ("video", VideoConversions) },

            { ".mp3", ("audio", AudioConversions) },
            { ".wav", ("audio", AudioConversions) },
            { ".ogg", ("audio", AudioConversions) },
            { ".opus", ("audio", AudioConversions) },
            { ".aac", ("audio", AudioConversions) },
            { ".flac", ("audio", AudioConversions) },
            { ".wma", ("audio", AudioConversions) },
            { ".m4a", ("audio", AudioConversions) },
            { ".ac3", ("audio", AudioConversions) },

            { ".jpg", ("image", ImageConversions) },
            { ".jpeg", ("image", ImageConversions) },
            { ".png", ("image", ImageConversions) },
            { ".bmp", ("image", ImageConversions) },
            { ".gif", ("image", ImageConversions) },
            { ".webp", ("image", ImageConversions) },
            { ".tiff", ("image", ImageConversions) },
            { ".ico", ("image", ImageConversions) },
            { ".heic", ("image", ImageConversions) }
        };

        private static string detectedGpu = null;

        static string DetectGpuType()
        {
            if (detectedGpu != null)
                return detectedGpu;

            var gpuInfo = new System.Text.StringBuilder();
            gpuInfo.AppendLine("GPU Detection Information:");
            gpuInfo.AppendLine("------------------------");
            
            try 
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "ffmpeg",
                    Arguments = "-hide_banner -hwaccels",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using (var process = Process.Start(startInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    process.WaitForExit();
                    
                    gpuInfo.AppendLine("FFmpeg Hardware Accelerators:");
                    gpuInfo.AppendLine(output);
                    
                    if (output.Contains("cuda") || output.Contains("nvenc"))
                    {
                        detectedGpu = "nvidia";
                        gpuInfo.AppendLine("Detected NVIDIA GPU via FFmpeg");
                    }
                    else if (output.Contains("amf") || output.Contains("d3d11va"))
                    {
                        detectedGpu = "amd";
                        gpuInfo.AppendLine("Detected AMD GPU via FFmpeg");
                    }
                    else if (output.Contains("qsv"))
                    {
                        detectedGpu = "intel";
                        gpuInfo.AppendLine("Detected Intel GPU via FFmpeg");
                    }
                }
                
                if (detectedGpu == null)
                {
                    startInfo = new ProcessStartInfo
                    {
                        FileName = "wmic",
                        Arguments = "path win32_VideoController get name",
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    };

                    using (var process = Process.Start(startInfo))
                    {
                        string output = process.StandardOutput.ReadToEnd().ToLower();
                        process.WaitForExit();
                        
                        gpuInfo.AppendLine("\nWMIC GPU Detection:");
                        gpuInfo.AppendLine(output);
                        
                        if (output.Contains("nvidia"))
                        {
                            detectedGpu = "nvidia";
                            gpuInfo.AppendLine("Detected NVIDIA GPU via WMIC");
                        }
                        else if (output.Contains("amd") || output.Contains("radeon"))
                        {
                            detectedGpu = "amd";
                            gpuInfo.AppendLine("Detected AMD GPU via WMIC");
                        }
                        else if (output.Contains("intel"))
                        {
                            detectedGpu = "intel";
                            gpuInfo.AppendLine("Detected Intel GPU via WMIC");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                gpuInfo.AppendLine($"Error detecting GPU: {ex.Message}");
            }
            
            if (detectedGpu == null)
            {
                detectedGpu = "generic";
                gpuInfo.AppendLine("No specific GPU detected, using generic settings");
            }
            
            gpuInfo.AppendLine($"\nFinal GPU Type: {detectedGpu}");
            
            //MessageBox.Show(gpuInfo.ToString(), "GPU Detection Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            return detectedGpu;
        }

        [STAThread]
        static void Main(string[] args) {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length == 0) {
                if (!IsAdministrator()) {
                    var procInfo = new ProcessStartInfo {
                        UseShellExecute = true,
                        WorkingDirectory = Environment.CurrentDirectory,
                        FileName = Application.ExecutablePath,
                        Verb = "runas"
                    };
                    try {
                        Process.Start(procInfo);
                    }
                    catch {
                    }

                    return;
                }

                if (!IsFfmpegInstalled())
                {
                    if (!InstallFfmpeg())
                    {
                        return;
                    }
                }

                RegisterContextMenus();
                CustomMessageBox.Show(
                    "Right-click on a file to convert it to another format.",
                    "Success", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Information);
                return;
            }

            static bool InstallFfmpeg()
            {
                try
                {
                    var result = CustomMessageBox.Show(
                        "FFmpeg is required but not found on your system.\n\nWould you like to install it now using winget?",
                        "FFmpeg Required",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        var startInfo = new ProcessStartInfo
                        {
                            FileName = "winget",
                            Arguments = "install ffmpeg",
                            UseShellExecute = true,
                            CreateNoWindow = false
                        };

                        using var process = Process.Start(startInfo);
                        process.WaitForExit();

                        if (process.ExitCode == 0)
                        {
                            CustomMessageBox.Show(
                                "FFmpeg was installed successfully. You may need to restart the application.", 
                                "Installation Complete", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Information);
                            return true;
                        }
                        else
                        {
                            CustomMessageBox.Show(
                                "FFmpeg installation failed. Please install FFmpeg manually.", 
                                "Installation Failed", 
                                MessageBoxButtons.OK, 
                                MessageBoxIcon.Error);
                            return false;
                        }
                    }
                    else
                    {
                        CustomMessageBox.Show(
                            "FFmpeg is required for this application to work.", 
                            "FFmpeg Required", 
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Warning);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    CustomMessageBox.Show(
                        $"Error during FFmpeg installation: {ex.Message}\n\nPlease install FFmpeg manually.", 
                        "Error", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Error);
                    return false;
                }
            }

            if (args.Length == 1 && args[0] == "-unregister") {
                if (!IsAdministrator()) {
                    var procInfo = new ProcessStartInfo {
                        UseShellExecute = true,
                        WorkingDirectory = Environment.CurrentDirectory,
                        FileName = Application.ExecutablePath,
                        Verb = "runas"
                    };
                    try {
                        Process.Start(procInfo);
                    }
                    catch {
                    }

                    return;
                }

                UnregisterContextMenus();
                MessageBox.Show("Context menus removed successfully.", "Info");
                return;
            }

            if (args.Length == 2 && args[1] == "mute") {
                string inputFile = args[0];
                string extension = Path.GetExtension(inputFile);
                string outputFile = Path.Combine(
                    Path.GetDirectoryName(inputFile),
                    Path.GetFileNameWithoutExtension(inputFile) + "_muted" + extension);

                try {
                    if (File.Exists(outputFile)) {
                        var result = MessageBox.Show(
                            $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                            "File Exists",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Cancel || result == DialogResult.No)
                            return;

                        try {
                            File.Delete(outputFile);
                        }
                        catch (Exception ex) {
                            MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                            return;
                        }
                    }

                    PerformMuteWithProgress(inputFile, outputFile);
                    OpenFolderAndSelectFile(outputFile);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Error during muting: {ex.Message}", "Error");
                    Environment.Exit(1);
                }

                return;
            }

            if (args.Length >= 2 && args[1] == "compress") {
                string inputFile = args[0];
                string extension = Path.GetExtension(inputFile);

                int targetSizeMB = 10;
                if (args.Length >= 3 && int.TryParse(args[2], out int size)) {
                    targetSizeMB = size;
                }

                string outputFile = Path.Combine(
                    Path.GetDirectoryName(inputFile),
                    Path.GetFileNameWithoutExtension(inputFile) + $"_compressed{targetSizeMB}MB" + extension);

                try {
                    if (File.Exists(outputFile)) {
                        var result = CustomMessageBox.Show(
                            $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                            "File Exists",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Cancel || result == DialogResult.No)
                            return;

                        try {
                            File.Delete(outputFile);
                        }
                        catch (Exception ex) {
                            MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                            return;
                        }
                    }

                    PerformCompressWithProgress(inputFile, outputFile, targetSizeMB);
                    OpenFolderAndSelectFile(outputFile);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Error during compression: {ex.Message}", "Error");
                    Environment.Exit(1);
                }

                return;
            }

            if (args.Length == 2 && (args[1] == "resize50" || args[1] == "resize75")) {
                string inputFile = args[0];
                string extension = Path.GetExtension(inputFile);
                string scale = args[1] == "resize50" ? "50%" : "75%";
                string scaleValue = args[1] == "resize50" ? "0.5" : "0.75";
                string outputFile = Path.Combine(
                    Path.GetDirectoryName(inputFile),
                    Path.GetFileNameWithoutExtension(inputFile) + "_" + scale.Replace("%", "pct") + extension);

                try {
                    if (File.Exists(outputFile)) {
                        var result = MessageBox.Show(
                            $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                            "File Exists",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Cancel || result == DialogResult.No)
                            return;

                        try {
                            File.Delete(outputFile);
                        }
                        catch (Exception ex) {
                            MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                            return;
                        }
                    }

                    PerformResizeWithProgress(inputFile, outputFile, scaleValue);
                    OpenFolderAndSelectFile(outputFile);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Error during resizing: {ex.Message}", "Error");
                    Environment.Exit(1);
                }

                return;
            }

            if (args.Length > 1) {
                if (args.Length > 2) {
                    string outputFormat = args.Last();
                    try {
                        foreach (var inputFile in args.Take(args.Length - 1)) {
                            string outputFile = Path.ChangeExtension(inputFile, outputFormat);
                            if (File.Exists(outputFile)) {
                                var result = MessageBox.Show(
                                    $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                                    "File Exists",
                                    MessageBoxButtons.YesNoCancel,
                                    MessageBoxIcon.Warning);

                                if (result == DialogResult.Cancel)
                                    return;
                                if (result == DialogResult.No)
                                    continue;
                                try {
                                    File.Delete(outputFile);
                                }
                                catch (Exception ex) {
                                    MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                                    continue;
                                }
                            }

                            PerformConversionWithProgress(inputFile, outputFile);
                        }

                        MessageBox.Show("Batch conversion completed successfully.", "Info");
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Error during batch conversion: {ex.Message}", "Error");
                        Environment.Exit(1);
                    }
                }
                else {
                    string inputFile = args[0];
                    string outputFormat = args[1];
                    string outputFile = Path.ChangeExtension(inputFile, outputFormat);
                    try {
                        if (File.Exists(outputFile)) {
                            var result = MessageBox.Show(
                                $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                                "File Exists",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Warning);

                            if (result == DialogResult.Cancel || result == DialogResult.No)
                                return;

                            try {
                                File.Delete(outputFile);
                            }
                            catch (Exception ex) {
                                MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                                return;
                            }
                        }

                        PerformConversionWithProgress(inputFile, outputFile);
                        OpenFolderAndSelectFile(outputFile);
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Error during conversion: {ex.Message}", "Error");
                        Environment.Exit(1);
                    }
                }
            }
        }

        static bool IsFfmpegInstalled()
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "ffmpeg",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using var process = Process.Start(startInfo);
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                return process.ExitCode == 0 && !string.IsNullOrWhiteSpace(output);
            }
            catch
            {
                return false;
            }
        }


        static bool IsAdministrator() {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RegisterContextMenus() {
            string execPath = Application.ExecutablePath;
            foreach (var mapping in FormatMappings) {
                string extension = mapping.Key;
                var (type, conversions) = mapping.Value;

                using var mainKey =
                    Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To");
                mainKey.SetValue("SubCommands", "");
                mainKey.SetValue("MUIVerb", "Convert To");
                using var shellKey = Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To\shell");

                mainKey.SetValue("Icon", $"\"{execPath}\",0");

                using var toolsKey = shellKey.CreateSubKey("!Tools");
                toolsKey.SetValue("MUIVerb", "Tools");
                toolsKey.SetValue("SubCommands", "");

                using var toolsShellKey = toolsKey.CreateSubKey("shell");

                if (type == "audio" || type == "video") {
                    using var muteKey = toolsShellKey.CreateSubKey("Mute");
                    muteKey.SetValue("", "Mute");
                    using var muteCmdKey = muteKey.CreateSubKey("command");
                    muteCmdKey.SetValue("", $"\"{execPath}\" \"%1\" mute");

                    using var compressKey = toolsShellKey.CreateSubKey("Compress");
                    compressKey.SetValue("MUIVerb", "Compress");
                    compressKey.SetValue("SubCommands", "");

                    using var compressShellKey = compressKey.CreateSubKey("shell");

                    using var compress10MB = compressShellKey.CreateSubKey("1_Compress_010MB");
                    compress10MB.SetValue("", "Compress to <10MB");
                    using var compress10MBCmd = compress10MB.CreateSubKey("command");
                    compress10MBCmd.SetValue("", $"\"{execPath}\" \"%1\" compress 10");

                    using var compress20MB = compressShellKey.CreateSubKey("2_Compress_020MB");
                    compress20MB.SetValue("", "Compress to <20MB");
                    using var compress20MBCmd = compress20MB.CreateSubKey("command");
                    compress20MBCmd.SetValue("", $"\"{execPath}\" \"%1\" compress 20");

                    using var compress50MB = compressShellKey.CreateSubKey("3_Compress_050MB");
                    compress50MB.SetValue("", "Compress to <50MB");
                    using var compress50MBCmd = compress50MB.CreateSubKey("command");
                    compress50MBCmd.SetValue("", $"\"{execPath}\" \"%1\" compress 50");

                    using var compress100MB = compressShellKey.CreateSubKey("4_Compress_100MB");
                    compress100MB.SetValue("", "Compress to <100MB");
                    using var compress100MBCmd = compress100MB.CreateSubKey("command");
                    compress100MBCmd.SetValue("", $"\"{execPath}\" \"%1\" compress 100");

                    using var compress500MB = compressShellKey.CreateSubKey("5_Compress_500MB");
                    compress500MB.SetValue("", "Compress to <500MB");
                    using var compress500MBCmd = compress500MB.CreateSubKey("command");
                    compress500MBCmd.SetValue("", $"\"{execPath}\" \"%1\" compress 500");
                }

                if (type == "image") {
                    using var resizeKey = toolsShellKey.CreateSubKey("Resize50");
                    resizeKey.SetValue("", "Resize to 50%");
                    using var resizeCmdKey = resizeKey.CreateSubKey("command");
                    resizeCmdKey.SetValue("", $"\"{execPath}\" \"%1\" resize50");

                    using var resize75Key = toolsShellKey.CreateSubKey("Resize75");
                    resize75Key.SetValue("", "Resize to 75%");
                    using var resize75CmdKey = resize75Key.CreateSubKey("command");
                    resize75CmdKey.SetValue("", $"\"{execPath}\" \"%1\" resize75");
                }

                foreach (var format in conversions) {
                    if (format == extension.TrimStart('.'))
                        continue;
                    using var formatKey = shellKey.CreateSubKey(format.ToUpper());
                    formatKey.SetValue("", format.ToUpper());
                    using var cmdKey = formatKey.CreateSubKey("command");
                    cmdKey.SetValue("", $"\"{execPath}\" \"%1\" {format}");
                }
            }
        }

        static void UnregisterContextMenus() {
            foreach (var mapping in FormatMappings) {
                string extension = mapping.Key;
                try {
                    Registry.ClassesRoot.DeleteSubKeyTree($@"SystemFileAssociations\{extension}\shell\Convert To", false);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Error removing menu for {extension}: {ex.Message}", "Error");
                }
            }

            try {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\MediaConverter", false);
            }
            catch { }
        }

        static void PerformConversionWithProgress(string input, string output) {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing) {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var conversionTask = Task.Run(() => {

                string inputExt = Path.GetExtension(input).ToLower();
                string outputExt = Path.GetExtension(output).ToLower();
                var (inputType, _) = FormatMappings[inputExt];
                string arguments = "";

                if (inputType == "image") {
                    switch (outputExt) {
                        case ".webp":
                            arguments = $"-i \"{input}\" -quality 90 -compression_level 6 \"{output}\"";
                            break;
                        case ".jpg":
                        case ".jpeg":
                            arguments = $"-i \"{input}\" -quality 95 \"{output}\"";
                            break;
                        case ".png":
                            arguments = $"-i \"{input}\" -compression_level 9 \"{output}\"";
                            break;
                        case ".tiff":
                            arguments = $"-i \"{input}\" -compression_algo lzw \"{output}\"";
                            break;
                        case ".gif":
                            if (inputExt == ".gif") {
                                arguments =
                                    $"-i \"{input}\" -lavfi \"fps=15,scale=trunc(iw/2)*2:trunc(ih/2)*2:flags=lanczos\" \"{output}\"";
                            }
                            else {
                                arguments = $"-i \"{input}\" \"{output}\"";
                            }

                            break;
                        case ".bmp":
                            arguments = $"-i \"{input}\" -pix_fmt rgb24 \"{output}\"";
                            break;
                        case ".ico":
                            arguments = $"-i \"{input}\" -vf scale=256:256 \"{output}\"";
                            break;
                        case ".heic":
                            arguments = $"-i \"{input}\" -quality 90 \"{output}\"";
                            break;
                        default:
                            arguments = $"-i \"{input}\" \"{output}\"";
                            break;
                    }
                }
                else if (inputType == "audio") {
                    switch (outputExt) {
                        case ".mp3":
                            arguments = $"-i \"{input}\" -vn -codec:a libmp3lame -q:a 0 \"{output}\"";
                            break;
                        case ".m4a":
                            arguments = $"-i \"{input}\" -vn -c:a aac -b:a 256k \"{output}\"";
                            break;
                        case ".flac":
                            arguments = $"-i \"{input}\" -vn -codec:a flac -compression_level 12 \"{output}\"";
                            break;
                        case ".opus":
                            arguments = $"-i \"{input}\" -vn -c:a libopus -b:a 192k \"{output}\"";
                            break;
                        case ".ogg":
                            arguments = $"-i \"{input}\" -vn -c:a libvorbis -q:a 7 \"{output}\"";
                            break;
                        case ".wav":
                            arguments = $"-i \"{input}\" -vn -c:a pcm_s24le \"{output}\"";
                            break;
                        case ".wma":
                            arguments = $"-i \"{input}\" -vn -c:a wmav2 -b:a 256k \"{output}\"";
                            break;
                        case ".aac":
                            arguments = $"-i \"{input}\" -vn -c:a aac -b:a 256k \"{output}\"";
                            break;
                        case ".ac3":
                            arguments = $"-i \"{input}\" -vn -c:a ac3 -b:a 448k \"{output}\"";
                            break;
                        default:
                            arguments = $"-i \"{input}\" -vn \"{output}\"";
                            break;
                    }
                }
                else if (inputType == "video") {
                    string videoCodec, audioCodec, extraParams = "";

                    switch (outputExt) {
                        case ".mp4":
                            videoCodec = "-c:v libx264 -crf 23 -preset medium";
                            audioCodec = "-c:a aac -b:a 192k";
                            extraParams = "-movflags +faststart";
                            break;
                        case ".mkv":
                            videoCodec = "-c:v libx264 -crf 21 -preset slower";
                            audioCodec = "-c:a libopus -b:a 192k";
                            extraParams = "";
                            break;
                        case ".webm":
                            videoCodec = "-c:v libvpx-vp9 -crf 30 -b:v 0";
                            audioCodec = "-c:a libopus -b:a 128k";
                            extraParams = "-deadline good -cpu-used 2";
                            break;
                        case ".avi":
                            videoCodec = "-c:v mpeg4 -qscale:v 3";
                            audioCodec = "-c:a mp3 -q:a 3";
                            extraParams = "";
                            break;
                        case ".wmv":
                            videoCodec = "-c:v wmv2 -qscale:v 3";
                            audioCodec = "-c:a wmav2 -b:a 256k";
                            extraParams = "";
                            break;
                        case ".flv":
                            videoCodec = "-c:v flv -qscale:v 3";
                            audioCodec = "-c:a mp3 -q:a 3";
                            extraParams = "";
                            break;
                        case ".mov":
                            videoCodec = "-c:v prores_ks -profile:v 3";
                            audioCodec = "-c:a pcm_s24le";
                            extraParams = "";
                            break;
                        case ".ts":
                            videoCodec = "-c:v libx264 -crf 23 -preset medium";
                            audioCodec = "-c:a aac -b:a 192k";
                            extraParams = "-f mpegts";
                            break;
                        case ".3gp":
                            videoCodec = "-c:v libx264 -crf 28 -preset faster -profile:v baseline -level 3.0";
                            audioCodec = "-c:a aac -b:a 128k -ac 2";
                            extraParams = "";
                            break;
                        case ".mp3":
                            videoCodec = "-vn";
                            audioCodec = "-codec:a libmp3lame -q:a 0";
                            extraParams = "";
                            break;
                        case ".m4a":
                            videoCodec = "-vn";
                            audioCodec = "-c:a aac -b:a 256k";
                            extraParams = "";
                            break;
                        default:
                            videoCodec = "-c:v libx264 -crf 23";
                            audioCodec = "-c:a aac -b:a 192k";
                            extraParams = "";
                            break;
                    }

                    arguments = $"-i \"{input}\" {videoCodec} {audioCodec} {extraParams} \"{output}\"";

                    if (inputType == "audio" && outputExt != ".mp3" && outputExt != ".m4a") {
                        arguments =
                            $"-i \"{input}\" -f lavfi -i color=c=black:s=1920x1080 -shortest {videoCodec} {audioCodec} {extraParams} \"{output}\"";
                    }
                }

                var startInfo = new ProcessStartInfo {
                    FileName = "ffmpeg",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    LoadUserProfile = false,
                    ErrorDialog = false,
                    WorkingDirectory = Path.GetDirectoryName(output),
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };

                ffmpegProcess = new Process { StartInfo = startInfo };

                TimeSpan duration = TimeSpan.Zero;
                TimeSpan currentTime = TimeSpan.Zero;

                ffmpegProcess.ErrorDataReceived += (sender, e) => {
                    if (string.IsNullOrEmpty(e.Data))
                        return;

                    string data = e.Data;

                    if (duration == TimeSpan.Zero) {
                        var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                        if (durationMatch.Success) {
                            int hours = int.Parse(durationMatch.Groups[1].Value);
                            int minutes = int.Parse(durationMatch.Groups[2].Value);
                            int seconds = int.Parse(durationMatch.Groups[3].Value);
                            int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                            duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                        }
                    }

                    var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                    if (timeMatch.Success) {
                        int hours = int.Parse(timeMatch.Groups[1].Value);
                        int minutes = int.Parse(timeMatch.Groups[2].Value);
                        int seconds = int.Parse(timeMatch.Groups[3].Value);
                        int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                        currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                        if (duration != TimeSpan.Zero) {
                            int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                            percentage = Math.Min(99, Math.Max(0, percentage));

                            var forms = Application.OpenForms;
                            foreach (Form form in forms) {
                                if (form is ProgressForm progressForm) {
                                    progressForm.BeginInvoke(new Action(() => {
                                        progressForm.UpdateProgress(percentage);
                                    }));
                                    break;
                                }
                            }
                        }
                    }
                };

                progressForm.Invoke(new Action(() => {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled()) {
                    string errorSummary = "FFmpeg failed to complete the operation.";
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}. {errorSummary}");
                }

                if (isReplacing && File.Exists(backupPath)) {
                    try {
                        File.Delete(backupPath);
                    }
                    catch {
                    }
                }
            });

            conversionTask.ContinueWith(t => {
                if (t.IsFaulted) {
                    progressForm.Invoke(new Action(() => {
                        if (!progressForm.WasCancelled()) {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else {
                            progressForm.Close();
                        }
                    }));
                }
                else {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try {
                conversionTask.Wait();
            }
            catch (AggregateException ae) {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void PerformCompressWithProgress(string input, string output, int targetSizeMB = 20) {
            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing) {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var compressTask = Task.Run(() => {
                string inputExt = Path.GetExtension(input).ToLower();
                var (inputType, _) = FormatMappings[inputExt];
                string arguments = "";

                var infoStartInfo = new ProcessStartInfo {
                    FileName = "ffprobe",
                    Arguments = $"-v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 \"{input}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    CreateNoWindow = true
                };

                using var infoProcess = new Process { StartInfo = infoStartInfo };
                infoProcess.Start();
                string durationStr = infoProcess.StandardOutput.ReadToEnd().Trim();
                infoProcess.WaitForExit();

                long targetSizeInBits = (long)(targetSizeMB * 0.95 * 8 * 1024 * 1024);
                double duration = 0;

                if (double.TryParse(durationStr, out duration) && duration > 0) {
                    if (inputType == "video") {
                        int videoBitrate = (int)((targetSizeInBits * 0.85) / duration);
                        int audioBitrate = (int)((targetSizeInBits * 0.15) / duration);

                        audioBitrate = Math.Min(256000, Math.Max(32000, audioBitrate));

                        int crf = 23;
                        
                        if (videoBitrate < 500000) crf = 28;
                        else if (videoBitrate < 1000000) crf = 26;
                        else if (videoBitrate < 2000000) crf = 24;

                        string gpuType = DetectGpuType();
                        
                        switch (gpuType) {
                            case "nvidia":
                                arguments = $"-hwaccel cuda -i \"{input}\" -c:v h264_nvenc -preset p2 " +
                                        $"-b:v {videoBitrate / 1000}k -maxrate {videoBitrate / 1000}k " +
                                        $"-bufsize {videoBitrate / 500}k -c:a aac -b:a {audioBitrate / 1000}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            case "amd":
                                arguments = $"-hwaccel d3d11va -i \"{input}\" -c:v h264_amf -quality balanced " +
                                        $"-b:v {videoBitrate / 1000}k -maxrate {videoBitrate / 1000}k " +
                                        $"-bufsize {videoBitrate / 500}k -c:a aac -b:a {audioBitrate / 1000}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            case "intel":
                                arguments = $"-hwaccel qsv -i \"{input}\" -c:v h264_qsv -preset medium " +
                                        $"-b:v {videoBitrate / 1000}k -maxrate {videoBitrate / 1000}k " +
                                        $"-bufsize {videoBitrate / 500}k -c:a aac -b:a {audioBitrate / 1000}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            default:
                                arguments = $"-i \"{input}\" -c:v libx264 -preset medium -crf {crf} " +
                                        $"-maxrate {videoBitrate / 1000}k -bufsize {videoBitrate / 500}k " +
                                        $"-c:a aac -b:a {audioBitrate / 1000}k -ac 2 -movflags +faststart \"{output}\"";
                                break;
                        }
                    }
                    else if (inputType == "audio") {
                        int audioBitrate = (int)(targetSizeInBits / duration);

                        int minBitrate = 32000;
                        int maxBitrate = 320000;

                        audioBitrate = Math.Min(maxBitrate, Math.Max(minBitrate, audioBitrate));

                        arguments = $"-i \"{input}\" -c:a aac -b:a {audioBitrate / 1000}k -ac 2 \"{output}\"";
                    }
                }
                else {
                    if (inputType == "video") {
                        int bitrate = targetSizeMB * 8000;
                        
                        string gpuType = DetectGpuType();
                        
                        switch (gpuType) {
                            case "nvidia":
                                arguments = $"-hwaccel cuda -i \"{input}\" -c:v h264_nvenc -preset p2 " +
                                        $"-b:v {bitrate / 10}k -c:a aac -b:a {Math.Min(128, targetSizeMB / 2)}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            case "amd":
                                arguments = $"-hwaccel d3d11va -i \"{input}\" -c:v h264_amf -quality balanced " +
                                        $"-b:v {bitrate / 10}k -c:a aac -b:a {Math.Min(128, targetSizeMB / 2)}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            case "intel":
                                arguments = $"-hwaccel qsv -i \"{input}\" -c:v h264_qsv -preset medium " +
                                        $"-b:v {bitrate / 10}k -c:a aac -b:a {Math.Min(128, targetSizeMB / 2)}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                            default:
                                arguments = $"-i \"{input}\" -c:v libx264 -preset medium " +
                                        $"-b:v {bitrate / 10}k -c:a aac -b:a {Math.Min(128, targetSizeMB / 2)}k -ac 2 " +
                                        $"-movflags +faststart \"{output}\"";
                                break;
                        }
                    }
                    else if (inputType == "audio") {
                        int bitrate = Math.Min(320, Math.Max(64, targetSizeMB * 10));

                        arguments = $"-i \"{input}\" -c:a aac -b:a {bitrate}k -ac 2 \"{output}\"";
                    }
                }

                var startInfo = new ProcessStartInfo {
                    FileName = "ffmpeg",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    LoadUserProfile = false,
                    ErrorDialog = false,
                    WorkingDirectory = Path.GetDirectoryName(output),
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };

                ffmpegProcess = new Process { StartInfo = startInfo };

                TimeSpan totalDuration = TimeSpan.Zero;
                TimeSpan currentTime = TimeSpan.Zero;

                ffmpegProcess.ErrorDataReceived += (sender, e) => {
                    if (string.IsNullOrEmpty(e.Data))
                        return;

                    string data = e.Data;

                    if (totalDuration == TimeSpan.Zero) {
                        var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                        if (durationMatch.Success) {
                            int hours = int.Parse(durationMatch.Groups[1].Value);
                            int minutes = int.Parse(durationMatch.Groups[2].Value);
                            int seconds = int.Parse(durationMatch.Groups[3].Value);
                            int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                            totalDuration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                        }
                    }

                    var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                    if (timeMatch.Success) {
                        int hours = int.Parse(timeMatch.Groups[1].Value);
                        int minutes = int.Parse(timeMatch.Groups[2].Value);
                        int seconds = int.Parse(timeMatch.Groups[3].Value);
                        int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                        currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                        if (totalDuration != TimeSpan.Zero) {
                            int percentage = (int)((currentTime.TotalMilliseconds / totalDuration.TotalMilliseconds) * 100);
                            percentage = Math.Min(99, Math.Max(0, percentage));

                            var forms = Application.OpenForms;
                            foreach (Form form in forms) {
                                if (form is ProgressForm progressForm) {
                                    progressForm.BeginInvoke(new Action(() => {
                                        progressForm.UpdateProgress(percentage);
                                    }));
                                    break;
                                }
                            }
                        }
                    }
                };

                progressForm.Invoke(new Action(() => {
                    progressForm.SetStatusText($"Compressing to {targetSizeMB}MB...");
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));


                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled()) {
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}");
                }

                if (File.Exists(output)) {
                    var fileInfo = new FileInfo(output);
                    if (fileInfo.Length > targetSizeMB * 1024 * 1024) {
                        progressForm.BeginInvoke(new Action(() => {
                            progressForm.SetStatusText($"Warning: File still exceeds {targetSizeMB}MB");
                        }));
                    }
                }

                if (isReplacing && File.Exists(backupPath)) {
                    try {
                        File.Delete(backupPath);
                    }
                    catch {
                    }
                }
            });

            compressTask.ContinueWith(t => {
                if (t.IsFaulted) {
                    progressForm.Invoke(new Action(() => {
                        if (!progressForm.WasCancelled()) {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else {
                            progressForm.Close();
                        }
                    }));
                }
                else {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try {
                compressTask.Wait();
            }
            catch (AggregateException ae) {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void PerformMuteWithProgress(string input, string output) {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing) {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var muteTask = Task.Run(() => {

                string inputExt = Path.GetExtension(input).ToLower();
                var (inputType, _) = FormatMappings[inputExt];
                string arguments;

                if (inputType == "video") {
                    arguments = $"-i \"{input}\" -c:v copy -c:a aac -af \"volume=0\" \"{output}\"";
                }
                else if (inputType == "audio") {
                    arguments = $"-i \"{input}\" -af \"volume=0\" \"{output}\"";
                }
                else {
                    throw new Exception("Mute operation only supports audio and video files.");
                }

                var startInfo = new ProcessStartInfo {
                    FileName = "ffmpeg",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    LoadUserProfile = false,
                    ErrorDialog = false,
                    WorkingDirectory = Path.GetDirectoryName(output),
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };

                ffmpegProcess = new Process { StartInfo = startInfo };

                TimeSpan duration = TimeSpan.Zero;
                TimeSpan currentTime = TimeSpan.Zero;

                ffmpegProcess.ErrorDataReceived += (sender, e) => {
                    if (string.IsNullOrEmpty(e.Data))
                        return;

                    string data = e.Data;

                    if (duration == TimeSpan.Zero) {
                        var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                        if (durationMatch.Success) {
                            int hours = int.Parse(durationMatch.Groups[1].Value);
                            int minutes = int.Parse(durationMatch.Groups[2].Value);
                            int seconds = int.Parse(durationMatch.Groups[3].Value);
                            int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                            duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                        }
                    }

                    var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                    if (timeMatch.Success) {
                        int hours = int.Parse(timeMatch.Groups[1].Value);
                        int minutes = int.Parse(timeMatch.Groups[2].Value);
                        int seconds = int.Parse(timeMatch.Groups[3].Value);
                        int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                        currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                        if (duration != TimeSpan.Zero) {
                            int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                            percentage = Math.Min(99, Math.Max(0, percentage));

                            var forms = Application.OpenForms;
                            foreach (Form form in forms) {
                                if (form is ProgressForm progressForm) {
                                    progressForm.BeginInvoke(new Action(() => {
                                        progressForm.UpdateProgress(percentage);
                                    }));
                                    break;
                                }
                            }
                        }
                    }
                };

                progressForm.Invoke(new Action(() => {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled()) {
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}");
                }

                if (isReplacing && File.Exists(backupPath)) {
                    try {
                        File.Delete(backupPath);
                    }
                    catch {
                    }
                }
            });

            muteTask.ContinueWith(t => {
                if (t.IsFaulted) {
                    progressForm.Invoke(new Action(() => {
                        if (!progressForm.WasCancelled()) {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else {
                            progressForm.Close();
                        }
                    }));
                }
                else {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try {
                muteTask.Wait();
            }
            catch (AggregateException ae) {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void PerformResizeWithProgress(string input, string output, string scale) {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing) {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex) {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var resizeTask = Task.Run(() => {

                string inputExt = Path.GetExtension(input).ToLower();
                var (inputType, _) = FormatMappings[inputExt];

                if (inputType != "image") {
                    throw new Exception("Resize operation only supports image files.");
                }

                string arguments = $"-i \"{input}\" -vf \"scale=iw*{scale}:ih*{scale}\" \"{output}\"";

                var startInfo = new ProcessStartInfo {
                    FileName = "ffmpeg",
                    Arguments = arguments,
                    UseShellExecute = false,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden,
                    LoadUserProfile = false,
                    ErrorDialog = false,
                    WorkingDirectory = Path.GetDirectoryName(output),
                    StandardErrorEncoding = System.Text.Encoding.UTF8
                };

                ffmpegProcess = new Process { StartInfo = startInfo };

                ffmpegProcess.ErrorDataReceived += (sender, e) => { };

                progressForm.Invoke(new Action(() => {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled()) {
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}");
                }

                if (isReplacing && File.Exists(backupPath)) {
                    try {
                        File.Delete(backupPath);
                    }
                    catch {
                    }
                }
            });

            resizeTask.ContinueWith(t => {
                if (t.IsFaulted) {
                    progressForm.Invoke(new Action(() => {
                        if (!progressForm.WasCancelled()) {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else {
                            progressForm.Close();
                        }
                    }));
                }
                else {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try {
                resizeTask.Wait();
            }
            catch (AggregateException ae) {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void MuteMedia(string input, string output) {
            string inputExt = Path.GetExtension(input).ToLower();
            var (inputType, _) = FormatMappings[inputExt];
            string arguments;
            if (inputType == "video") {
                arguments = $"-i \"{input}\" -c:v copy -c:a aac -af \"volume=0\" \"{output}\"";
            }
            else if (inputType == "audio") {
                arguments = $"-i \"{input}\" -af \"volume=0\" \"{output}\"";
            }
            else {
                throw new Exception("Mute operation only supports audio and video files.");
            }

            var startInfo = new ProcessStartInfo {
                FileName = "ffmpeg",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                LoadUserProfile = false,
                ErrorDialog = false,
                WorkingDirectory = Path.GetDirectoryName(output),
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };

            using var process = new Process { StartInfo = startInfo };

            TimeSpan duration = TimeSpan.Zero;
            TimeSpan currentTime = TimeSpan.Zero;

            process.ErrorDataReceived += (sender, e) => {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero) {
                    var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success) {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success) {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero) {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms) {
                            if (form is ProgressForm progressForm) {
                                progressForm.BeginInvoke(new Action(() => {
                                    progressForm.UpdateProgress(percentage);
                                }));
                                break;
                            }
                        }
                    }
                }
            };

            process.Start();
            process.BeginErrorReadLine();
            process.WaitForExit();

            if (process.ExitCode != 0) {
                throw new Exception($"FFmpeg exited with code {process.ExitCode}");
            }
        }

        static void ResizeImage(string input, string output, string scale) {
            string inputExt = Path.GetExtension(input).ToLower();
            var (inputType, _) = FormatMappings[inputExt];

            if (inputType != "image") {
                throw new Exception("Resize operation only supports image files.");
            }

            string arguments = $"-i \"{input}\" -vf \"scale=iw*{scale}:ih*{scale}\" \"{output}\"";

            var startInfo = new ProcessStartInfo {
                FileName = "ffmpeg",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                LoadUserProfile = false,
                ErrorDialog = false,
                WorkingDirectory = Path.GetDirectoryName(output),
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };

            using var process = new Process { StartInfo = startInfo };

            TimeSpan duration = TimeSpan.Zero;
            TimeSpan currentTime = TimeSpan.Zero;

            process.ErrorDataReceived += (sender, e) => {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero) {
                    var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success) {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success) {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero) {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms) {
                            if (form is ProgressForm progressForm) {
                                progressForm.BeginInvoke(new Action(() => {
                                    progressForm.UpdateProgress(percentage);
                                }));
                                break;
                            }
                        }
                    }
                }
            };

            process.Start();
            process.BeginErrorReadLine();
            process.WaitForExit();

            if (process.ExitCode != 0) {
                throw new Exception($"FFmpeg exited with code {process.ExitCode}");
            }
        }

        static void ConvertMedia(string input, string output) {
            string inputExt = Path.GetExtension(input).ToLower();
            string outputExt = Path.GetExtension(output).ToLower();
            var (inputType, _) = FormatMappings[inputExt];
            string arguments = "";

            if (inputType == "image") {
                switch (outputExt) {
                    case ".webp":
                        arguments = $"-i \"{input}\" -quality 90 -compression_level 6 \"{output}\"";
                        break;
                    case ".jpg":
                    case ".jpeg":
                        arguments = $"-i \"{input}\" -quality 95 \"{output}\"";
                        break;
                    case ".png":
                        arguments = $"-i \"{input}\" -compression_level 9 \"{output}\"";
                        break;
                    case ".tiff":
                        arguments = $"-i \"{input}\" -compression_algo lzw \"{output}\"";
                        break;
                    case ".gif":
                        if (inputExt == ".gif") {

                            arguments =
                                $"-i \"{input}\" -lavfi \"fps=15,scale=trunc(iw/2)*2:trunc(ih/2)*2:flags=lanczos\" \"{output}\"";
                        }
                        else {
                            arguments = $"-i \"{input}\" \"{output}\"";
                        }

                        break;
                    case ".bmp":
                        arguments = $"-i \"{input}\" -pix_fmt rgb24 \"{output}\"";
                        break;
                    case ".ico":
                        arguments = $"-i \"{input}\" -vf scale=256:256 \"{output}\"";
                        break;
                    case ".heic":
                        arguments = $"-i \"{input}\" -quality 90 \"{output}\"";
                        break;
                    default:
                        arguments = $"-i \"{input}\" \"{output}\"";
                        break;
                }
            }
            else if (inputType == "audio") {
                switch (outputExt) {
                    case ".mp3":
                        arguments = $"-i \"{input}\" -vn -codec:a libmp3lame -q:a 0 \"{output}\"";
                        break;
                    case ".m4a":
                        arguments = $"-i \"{input}\" -vn -c:a aac -b:a 256k \"{output}\"";
                        break;
                    case ".flac":
                        arguments = $"-i \"{input}\" -vn -codec:a flac -compression_level 12 \"{output}\"";
                        break;
                    case ".opus":
                        arguments = $"-i \"{input}\" -vn -c:a libopus -b:a 192k \"{output}\"";
                        break;
                    case ".ogg":
                        arguments = $"-i \"{input}\" -vn -c:a libvorbis -q:a 7 \"{output}\"";
                        break;
                    case ".wav":
                        arguments = $"-i \"{input}\" -vn -c:a pcm_s24le \"{output}\"";
                        break;
                    case ".wma":
                        arguments = $"-i \"{input}\" -vn -c:a wmav2 -b:a 256k \"{output}\"";
                        break;
                    case ".aac":
                        arguments = $"-i \"{input}\" -vn -c:a aac -b:a 256k \"{output}\"";
                        break;
                    case ".ac3":
                        arguments = $"-i \"{input}\" -vn -c:a ac3 -b:a 448k \"{output}\"";
                        break;
                    default:
                        arguments = $"-i \"{input}\" -vn \"{output}\"";
                        break;
                }
            }
            else if (inputType == "video") {
                string videoCodec, audioCodec, extraParams = "";
                string baseCodec = "h264";
                string gpuType = DetectGpuType();
                string hwAccelParams = "";

                switch (outputExt) {
                    case ".mp4":
                        videoCodec = "-c:v libx264 -crf 23 -preset medium";
                        audioCodec = "-c:a aac -b:a 192k";
                        extraParams = "-movflags +faststart";
                        break;
                    case ".mkv":
                        videoCodec = "-c:v libx264 -crf 21 -preset slower";
                        audioCodec = "-c:a libopus -b:a 192k";
                        extraParams = "";
                        break;
                    case ".webm":
                        videoCodec = "-c:v libvpx-vp9 -crf 30 -b:v 0";
                        audioCodec = "-c:a libopus -b:a 128k";
                        extraParams = "-deadline good -cpu-used 2";
                        break;
                    case ".avi":
                        videoCodec = "-c:v mpeg4 -qscale:v 3";
                        audioCodec = "-c:a mp3 -q:a 3";
                        extraParams = "";
                        break;
                    case ".wmv":
                        videoCodec = "-c:v wmv2 -qscale:v 3";
                        audioCodec = "-c:a wmav2 -b:a 256k";
                        extraParams = "";
                        break;
                    case ".flv":
                        videoCodec = "-c:v flv -qscale:v 3";
                        audioCodec = "-c:a mp3 -q:a 3";
                        extraParams = "";
                        break;
                    case ".mov":
                        videoCodec = "-c:v prores_ks -profile:v 3";
                        audioCodec = "-c:a pcm_s24le";
                        extraParams = "";
                        break;
                    case ".ts":
                        videoCodec = "-c:v libx264 -crf 23 -preset medium";
                        audioCodec = "-c:a aac -b:a 192k";
                        extraParams = "-f mpegts";
                        break;
                    case ".3gp":
                        videoCodec = "-c:v libx264 -crf 28 -preset faster -profile:v baseline -level 3.0";
                        audioCodec = "-c:a aac -b:a 128k -ac 2";
                        extraParams = "";
                        break;
                    case ".mp3":
                        videoCodec = "-vn";
                        audioCodec = "-codec:a libmp3lame -q:a 0";
                        extraParams = "";
                        break;
                    case ".m4a":
                        videoCodec = "-vn";
                        audioCodec = "-c:a aac -b:a 256k";
                        extraParams = "";
                        break;
                    default:
                        videoCodec = "-c:v libx264 -crf 23";
                        audioCodec = "-c:a aac -b:a 192k";
                        extraParams = "";
                        break;
                }

                if (!outputExt.Equals(".mp3") && !outputExt.Equals(".m4a") && !videoCodec.Contains("-vn")) {
                    bool isH264 = videoCodec.Contains("libx264");
                    bool isVP9 = videoCodec.Contains("libvpx-vp9");
                    
                    switch (gpuType) {
                        case "nvidia":
                            if (isH264) {
                                string qualityParams = "";
                                if (videoCodec.Contains("-preset")) {
                                    qualityParams = "-preset p2";
                                }
                                videoCodec = videoCodec.Replace("-c:v libx264", "-hwaccel cuda -c:v h264_nvenc")
                                                    .Replace("-crf 23", "")
                                                    .Replace("-preset medium", "")
                                                    .Replace("-preset slower", "")
                                                    .Replace("-preset faster", "") + " " + qualityParams;
                            }
                            break;
                        case "amd":
                            if (isH264) {
                                videoCodec = videoCodec.Replace("-c:v libx264", "-hwaccel d3d11va -c:v h264_amf")
                                                    .Replace("-crf 23", "")
                                                    .Replace("-crf 21", "")
                                                    .Replace("-crf 28", "")
                                                    .Replace("-preset medium", "-quality balanced")
                                                    .Replace("-preset slower", "-quality quality")
                                                    .Replace("-preset faster", "-quality speed");
                            }
                            break;
                        case "intel":
                            if (isH264) {
                                string speedPreset = "-preset medium";
                                if (videoCodec.Contains("-preset slower")) speedPreset = "-preset slow";
                                if (videoCodec.Contains("-preset faster")) speedPreset = "-preset fast";
                                
                                videoCodec = videoCodec.Replace("-c:v libx264", "-hwaccel qsv -c:v h264_qsv")
                                                    .Replace("-crf 23", "")
                                                    .Replace("-crf 21", "")
                                                    .Replace("-crf 28", "")
                                                    .Replace("-preset medium", speedPreset)
                                                    .Replace("-preset slower", speedPreset)
                                                    .Replace("-preset faster", speedPreset);
                            }
                            break;
                    }
                }

                arguments = $"-i \"{input}\" {videoCodec} {audioCodec} {extraParams} \"{output}\"";

                if (inputType == "audio" && outputExt != ".mp3" && outputExt != ".m4a") {
                    arguments =
                        $"-i \"{input}\" -f lavfi -i color=c=black:s=1920x1080 -shortest {videoCodec} {audioCodec} {extraParams} \"{output}\"";
                }
            }

            var startInfo = new ProcessStartInfo {
                FileName = "ffmpeg",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                LoadUserProfile = false,
                ErrorDialog = false,
                WorkingDirectory = Path.GetDirectoryName(output),
                StandardErrorEncoding = System.Text.Encoding.UTF8
            };

            using var process = new Process { StartInfo = startInfo };

            TimeSpan duration = TimeSpan.Zero;
            TimeSpan currentTime = TimeSpan.Zero;

            process.ErrorDataReceived += (sender, e) => {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero) {
                    var durationMatch =
                        System.Text.RegularExpressions.Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success) {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = System.Text.RegularExpressions.Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success) {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero) {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms) {
                            if (form is ProgressForm progressForm) {
                                progressForm.BeginInvoke(new Action(() => {
                                    progressForm.UpdateProgress(percentage);
                                }));
                                break;
                            }
                        }
                    }
                }
            };

            process.Start();
            process.BeginErrorReadLine();
            process.WaitForExit();

            if (process.ExitCode != 0) {
                string errorSummary = "FFmpeg failed to complete the operation.";
                throw new Exception($"FFmpeg exited with code {process.ExitCode}. {errorSummary}");
            }
        }

        static void OpenFolderAndSelectFile(string filePath) {
            try {
                string folder = Path.GetDirectoryName(filePath);

                if (!File.Exists(filePath)) {

                    int attempts = 0;
                    while (!File.Exists(filePath) && attempts < 10) {
                        System.Threading.Thread.Sleep(100);
                        attempts++;
                    }

                    if (!File.Exists(filePath)) {
                        return;
                    }
                }

                bool usedExisting = false;

                try {
                    Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                    dynamic shell = Activator.CreateInstance(shellAppType);
                    foreach (dynamic window in shell.Windows()) {
                        try {
                            string windowPath = new Uri((string)window.LocationURL).LocalPath;
                            if (string.Equals(windowPath.TrimEnd('\\'), folder.TrimEnd('\\'),
                                    StringComparison.OrdinalIgnoreCase)) {
                                window.Document.SelectItem(filePath, 0);
                                usedExisting = true;
                                break;
                            }
                        }
                        catch {
                            continue;
                        }
                    }
                }
                catch {

                }

                if (!usedExisting) {

                    System.Threading.Thread.Sleep(200);

                    try {
                        IntPtr filePidl;
                        uint sfgao;
                        int hr = SHParseDisplayName(filePath, IntPtr.Zero, out filePidl, 0, out sfgao);
                        if (hr != 0) {

                            Process.Start("explorer.exe", folder);
                            return;
                        }

                        IntPtr folderPidl;
                        hr = SHParseDisplayName(folder, IntPtr.Zero, out folderPidl, 0, out sfgao);
                        if (hr != 0) {
                            CoTaskMemFree(filePidl);

                            Process.Start("explorer.exe", folder);
                            return;
                        }

                        IntPtr relativePidl = ILFindChild(folderPidl, filePidl);
                        if (relativePidl == IntPtr.Zero)
                            relativePidl = filePidl;

                        IntPtr[] items = { relativePidl };
                        hr = SHOpenFolderAndSelectItems(folderPidl, (uint)items.Length, items, 0);
                        CoTaskMemFree(filePidl);
                        CoTaskMemFree(folderPidl);

                        if (hr != 0) {

                        }
                    }
                    catch {

                    }
                }
            }
            catch {

                try {
                    string folder = Path.GetDirectoryName(filePath);
                }
                catch {

                }
            }
        }

        class ProgressForm : Form {
            private Panel progressBarContainer;
            private Panel progressBarFill;
            private Label statusLabel;
            private Label percentLabel;
            private System.Windows.Forms.Timer shimmerTimer;
            private System.Windows.Forms.Timer animationTimer;
            private bool isErrorState = false;

            private int targetProgress = 0;
            private double currentProgress = 0;
            private const double ANIMATION_SPEED = 0.15;

            private Process currentProcess;
            private string outputFilePath;
            private string backupFilePath;
            private bool isReplacement = false;

            [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
            private static extern IntPtr CreateRoundRectRgn(
                int nLeftRect,
                int nTopRect,
                int nRightRect,
                int nBottomRect,
                int nWidthEllipse,
                int nHeightEllipse
            );

            [DllImport("user32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

            [DllImport("user32.dll")]
            public static extern bool ReleaseCapture();

            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            public void SetProcessInfo(Process process, string outputPath, bool isReplacingFile = false,
                string originalBackupPath = null) {
                currentProcess = process;
                outputFilePath = outputPath;
                isReplacement = isReplacingFile;
                backupFilePath = originalBackupPath;
            }

            private void FormDragMouseDown(object sender, MouseEventArgs e) {
                if (e.Button == MouseButtons.Left) {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            }

            

            public ProgressForm() {
                try {
                    this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    this.ShowIcon = true;
                }
                catch {

                }
                this.Text = "Converting...";
                this.TopMost = true;
                this.ClientSize = new Size(360, 100);
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.ControlBox = false;
                this.DoubleBuffered = true;

                this.BackColor = Color.Black;

                this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 8, 8));

                percentLabel = new Label {
                    Text = "0%",
                    Width = 336,
                    Height = 20,
                    Location = new Point(12, 12),
                    Font = new Font("Segoe UI", 14F, FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = Color.White,
                    BackColor = Color.Transparent
                };

                progressBarContainer = new Panel {
                    Width = 336,
                    Height = 6,
                    Location = new Point(12, 40),
                    BackColor = Color.FromArgb(26, 26, 26),
                };
                progressBarContainer.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarContainer.Width, progressBarContainer.Height, 3,
                        3));

                progressBarFill = new Panel {
                    Width = 0,
                    Height = 6,
                    Location = new Point(0, 0),
                    BackColor = Color.FromArgb(0, 114, 255)
                };
                progressBarFill.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarFill.Width, progressBarFill.Height, 3, 3));
                progressBarContainer.Controls.Add(progressBarFill);

                statusLabel = new Label {
                    Text = "Initializing...",
                    Width = 336,
                    Height = 20,
                    Location = new Point(12, 55),
                    Font = new Font("Segoe UI", 9F, FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = Color.FromArgb(204, 204, 204),
                    BackColor = Color.Transparent
                };

                this.MouseDown += FormDragMouseDown;
                percentLabel.MouseDown += FormDragMouseDown;
                progressBarContainer.MouseDown += FormDragMouseDown;
                progressBarFill.MouseDown += FormDragMouseDown;
                statusLabel.MouseDown += FormDragMouseDown;

                Label closeButton = new Label {
                    Text = "✕",
                    Size = new Size(24, 24),
                    Location = new Point(Width - 30, 8),
                    Font = new Font("Arial", 10F, FontStyle.Bold),
                    ForeColor = Color.Gray,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = Color.Transparent
                };
                closeButton.Click += (s, e) => {
                    DialogResult result = MessageBox.Show(
                        "Are you sure you want to cancel the operation?",
                        "Cancel",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (result == DialogResult.Yes) {
                        HandleCancellation();
                        this.Close();
                    }
                };
                closeButton.MouseEnter += (s, e) => closeButton.ForeColor = Color.White;
                closeButton.MouseLeave += (s, e) => closeButton.ForeColor = Color.Gray;

                this.Controls.Add(percentLabel);
                this.Controls.Add(progressBarContainer);
                this.Controls.Add(statusLabel);
                this.Controls.Add(closeButton);

                closeButton.BringToFront();

                this.FormClosing += (s, e) => {
                    if (currentProcess != null && !currentProcess.HasExited) {
                        DialogResult result = MessageBox.Show(
                            "Are you sure you want to cancel the operation?",
                            "Cancel",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Yes) {
                            HandleCancellation();
                        }
                        else {
                            e.Cancel = true;
                        }
                    }
                };

                UpdateProgress(0);
                statusLabel.Text = "Starting process...";

                shimmerTimer = new System.Windows.Forms.Timer();
                shimmerTimer.Interval = 50;
                shimmerTimer.Tick += (sender, e) => { progressBarFill.Invalidate(); };
                shimmerTimer.Start();

                animationTimer = new System.Windows.Forms.Timer();
                animationTimer.Interval = 16;
                animationTimer.Tick += (sender, e) => {
                    if (Math.Abs(currentProgress - targetProgress) > 0.1) {
                        currentProgress += (targetProgress - currentProgress) * ANIMATION_SPEED;
                        UpdateVisualProgress((int)Math.Round(currentProgress));
                    }
                };
                animationTimer.Start();

                progressBarFill.Paint += ProgressBarFill_Paint;
            }

            private bool wasCancelled = false;

            public bool WasCancelled() {
                return wasCancelled;
            }

            private void HandleCancellation() {
                try {
                    wasCancelled = true;

                    if (currentProcess != null && !currentProcess.HasExited) {
                        try {
                            currentProcess.Kill(true);
                        }
                        catch {

                            currentProcess.Kill();
                        }

                        currentProcess = null;
                    }

                    if (!string.IsNullOrEmpty(outputFilePath)) {
                        try {

                            if (File.Exists(outputFilePath)) {
                                File.Delete(outputFilePath);
                            }

                            if (isReplacement && !string.IsNullOrEmpty(backupFilePath) && File.Exists(backupFilePath)) {
                                File.Move(backupFilePath, outputFilePath, true);
                            }
                        }
                        catch (Exception ex) {
                            MessageBox.Show($"Warning: Could not clean up files: {ex.Message}", "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex) {
                    MessageBox.Show($"Error during cancellation: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void ProgressBarFill_Paint(object sender, PaintEventArgs e) {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                if (isErrorState) {
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                               progressBarFill.ClientRectangle,
                               Color.FromArgb(255, 100, 100),
                               Color.FromArgb(200, 50, 50),
                               LinearGradientMode.Horizontal)) {
                        e.Graphics.FillRectangle(brush, progressBarFill.ClientRectangle);
                    }
                }
                else {
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                               progressBarFill.ClientRectangle,
                               Color.FromArgb(0, 198, 255),
                               Color.FromArgb(0, 114, 255),
                               LinearGradientMode.Horizontal)) {
                        e.Graphics.FillRectangle(brush, progressBarFill.ClientRectangle);
                    }

                    int shimmerWidth = progressBarFill.Width / 3;
                    int shimmerPosition =
                        (int)(DateTime.Now.Millisecond / 1000.0 * (progressBarFill.Width + shimmerWidth * 2)) -
                        shimmerWidth;

                    using (LinearGradientBrush shimmerBrush = new LinearGradientBrush(
                               new Rectangle(shimmerPosition, 0, shimmerWidth, progressBarFill.Height),
                               Color.FromArgb(0, Color.White),
                               Color.FromArgb(40, Color.White),
                               LinearGradientMode.Horizontal)) {
                        e.Graphics.FillRectangle(shimmerBrush,
                            new Rectangle(shimmerPosition, 0, shimmerWidth, progressBarFill.Height));
                    }
                }
            }

            private void UpdateVisualProgress(int value) {
                if (InvokeRequired) {
                    Invoke(new Action<int>(UpdateVisualProgress), value);
                    return;
                }

                int fillWidth = (int)((progressBarContainer.Width * value) / 100.0);
                progressBarFill.Width = fillWidth;
                progressBarFill.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarFill.Width, progressBarFill.Height, 3, 3));
            }

            public void UpdateProgress(int value) {
                if (InvokeRequired) {
                    Invoke(new Action<int>(UpdateProgress), value);
                    return;
                }

                targetProgress = value;
                percentLabel.Text = $"{value}%";

                if (value >= 100) {
                    statusLabel.Text = "Operation completed successfully!";
                }
                else if (value >= 75) {
                    statusLabel.Text = "Finalizing...";
                }
                else if (value >= 50) {
                    statusLabel.Text = "Processing...";
                }
                else if (value >= 25) {
                    statusLabel.Text = "Converting...";
                }
                else {
                    statusLabel.Text = "Preparing...";
                }

                string gpuType = DetectGpuType();
                statusLabel.Text += $" ({gpuType})";
            }

            public void SetStatusText(string status) {
                if (InvokeRequired) {
                    Invoke(new Action<string>(SetStatusText), status);
                    return;
                }

                statusLabel.Text = status;
            }

            public void SetCompleted() {
                if (InvokeRequired) {
                    Invoke(new Action(SetCompleted));
                    return;
                }

                UpdateProgress(100);
                statusLabel.Text = "Operation completed successfully!";
                System.Threading.Thread.Sleep(500);
                Close();
            }

            public void SetError(string errorMessage) {
                if (InvokeRequired) {
                    Invoke(new Action<string>(SetError), errorMessage);
                    return;
                }

                statusLabel.Text = $"Error: {errorMessage}";
                statusLabel.ForeColor = Color.FromArgb(255, 100, 100);
                isErrorState = true;
                progressBarFill.Invalidate();
            }

            protected override void Dispose(bool disposing) {
                if (disposing) {
                    shimmerTimer?.Dispose();
                    animationTimer?.Dispose();
                }

                base.Dispose(disposing);
            }
        }

        class CustomMessageBox : Form
        {
            private Panel mainPanel;
            private Label messageLabel;
            private Label titleLabel;
            private Button yesButton;
            private Button noButton;
            private Button okButton;
            
            [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
            private static extern IntPtr CreateRoundRectRgn(
                int nLeftRect,
                int nTopRect,
                int nRightRect,
                int nBottomRect,
                int nWidthEllipse,
                int nHeightEllipse
            );

            [DllImport("user32.dll")]
            public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

            [DllImport("user32.dll")]
            public static extern bool ReleaseCapture();

            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            public CustomMessageBox(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
            {
                int height = 120;
                if (message == "FFmpeg is required but not found on your system.\n\nWould you like to install it now using winget?") height = 170;
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.Size = new Size(360, height);
                this.BackColor = Color.Black;
                this.ForeColor = Color.White;
                this.Font = new Font("Segoe UI", 9F);
                this.TopMost = true;
                this.ControlBox = false;
                this.DoubleBuffered = true;
                
                this.MouseDown += FormDragMouseDown;
                
                try {
                    this.Icon = System.Drawing.Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                    this.ShowIcon = true;
                } 
                catch { }
                
                this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, this.Width, this.Height, 8, 8));
                
                InitializeComponents(message, title, buttons, icon);
            }

            private void InitializeComponents(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
            {
                titleLabel = new Label {
                    Text = title,
                    Width = 336,
                    Height = 25,
                    Location = new Point(12, 12),
                    Font = new Font("Segoe UI", 14F, FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = Color.White,
                    BackColor = Color.Transparent
                };

                int height = 40;
                if (title == "FFmpeg Required") height = 60;
                messageLabel = new Label {
                    Text = message,
                    Width = 336,
                    Height = height,
                    Location = new Point(12, 45),
                    Font = new Font("Segoe UI", 9F),
                    ForeColor = Color.FromArgb(204, 204, 204),
                    BackColor = Color.Transparent,
                    TextAlign = ContentAlignment.TopLeft
                };
                
                Label closeButton = new Label {
                    Text = "✕",
                    Size = new Size(24, 24),
                    Location = new Point(Width - 30, 8),
                    Font = new Font("Arial", 10F, FontStyle.Bold),
                    ForeColor = Color.Gray,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = Color.Transparent
                };
                
                closeButton.Click += (s, e) => { 
                    this.DialogResult = DialogResult.Cancel;
                    this.Close();
                };
                
                closeButton.MouseEnter += (s, e) => closeButton.ForeColor = Color.White;
                closeButton.MouseLeave += (s, e) => closeButton.ForeColor = Color.Gray;
                
                titleLabel.MouseDown += FormDragMouseDown;
                messageLabel.MouseDown += FormDragMouseDown;
                
                if (buttons == MessageBoxButtons.YesNo)
                {
                    yesButton = new Button {
                        Text = "Yes",
                        Size = new Size(80, 30),
                        Location = new Point(Width - 175, Height - 45),
                        FlatStyle = FlatStyle.Flat,
                        ForeColor = Color.White,
                        BackColor = Color.FromArgb(0, 114, 255),
                        Font = new Font("Segoe UI", 9F)
                    };
                    yesButton.FlatAppearance.BorderSize = 0;

                    yesButton.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, yesButton.Width, yesButton.Height, 15, 15));
                    
                    noButton = new Button {
                        Text = "No",
                        Size = new Size(80, 30),
                        Location = new Point(Width - 90, Height - 45),
                        FlatStyle = FlatStyle.Flat,
                        ForeColor = Color.White,
                        BackColor = Color.FromArgb(60, 60, 60),
                        Font = new Font("Segoe UI", 9F)
                    };
                    noButton.FlatAppearance.BorderSize = 0;

                    noButton.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, noButton.Width, noButton.Height, 15, 15));
                    
                    yesButton.Click += (s, e) => { this.DialogResult = DialogResult.Yes; this.Close(); };
                    noButton.Click += (s, e) => { this.DialogResult = DialogResult.No; this.Close(); };
                    
                    this.Controls.Add(yesButton);
                    this.Controls.Add(noButton);
                }
                else
                {
                    okButton = new Button {
                        Text = "OK",
                        Size = new Size(80, 30),
                        Location = new Point(Width - 90, Height - 45),
                        FlatStyle = FlatStyle.Flat,
                        ForeColor = Color.White,
                        BackColor = Color.FromArgb(0, 114, 255),
                        Font = new Font("Segoe UI", 9F)
                    };
                    okButton.FlatAppearance.BorderSize = 0;
                    okButton.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, okButton.Width, okButton.Height, 15, 15));
                    
                    okButton.Click += (s, e) => { this.DialogResult = DialogResult.OK; this.Close(); };
                    this.Controls.Add(okButton);
                }

                if (icon != MessageBoxIcon.None)
                {
                    Panel iconPanel = new Panel {
                        Size = new Size(24, 24),
                        Location = new Point(12, 45),
                        BackColor = Color.Transparent
                    };
                    
                    Color iconColor;
                    switch (icon)
                    {
                        case MessageBoxIcon.Error:
                            iconColor = Color.FromArgb(255, 100, 100);
                            break;
                        case MessageBoxIcon.Warning:
                            iconColor = Color.FromArgb(255, 170, 0);
                            break;
                        case MessageBoxIcon.Question:
                            iconColor = Color.FromArgb(0, 114, 255);
                            break;
                        case MessageBoxIcon.Information:
                        default:
                            iconColor = Color.FromArgb(0, 198, 255);
                            break;
                    }
                    
                    iconPanel.Paint += (s, e) => {
                        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        
                        if (icon == MessageBoxIcon.Error)
                        {
                            using (Pen pen = new Pen(iconColor, 3)) {
                                e.Graphics.DrawLine(pen, 4, 4, 20, 20);
                                e.Graphics.DrawLine(pen, 4, 20, 20, 4);
                            }
                        }
                        else if (icon == MessageBoxIcon.Warning)
                        {
                            using (Brush brush = new SolidBrush(iconColor)) {
                                e.Graphics.FillPolygon(brush, new Point[] {
                                    new Point(12, 2),
                                    new Point(22, 20),
                                    new Point(2, 20)
                                });
                            }
                            using (Brush brush = new SolidBrush(Color.Black)) {
                                e.Graphics.FillRectangle(brush, 11, 8, 2, 6);
                                e.Graphics.FillRectangle(brush, 11, 16, 2, 2);
                            }
                        }
                        else if (icon == MessageBoxIcon.Question)
                        {
                            using (Brush brush = new SolidBrush(iconColor)) {
                                e.Graphics.FillEllipse(brush, 2, 2, 20, 20);
                            }
                            using (Brush brush = new SolidBrush(Color.White)) {
                                Font font = new Font("Arial", 14, FontStyle.Bold);
                                e.Graphics.DrawString("?", font, brush, 7, 1);
                            }
                        }
                        else
                        {
                            using (Brush brush = new SolidBrush(iconColor)) {
                                e.Graphics.FillEllipse(brush, 2, 2, 20, 20);
                            }
                            using (Brush brush = new SolidBrush(Color.White)) {
                                Font font = new Font("Arial", 14, FontStyle.Bold);
                                e.Graphics.DrawString("i", font, brush, 9, 2);
                            }
                        }
                    };
                    
                    messageLabel.Location = new Point(45, 45);
                    messageLabel.Width = 303;
                    
                    this.Controls.Add(iconPanel);
                }
                
                this.Controls.Add(titleLabel);
                this.Controls.Add(messageLabel);
                this.Controls.Add(closeButton);
                closeButton.BringToFront();
            }
            
            private void FormDragMouseDown(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            }
            
            public static DialogResult Show(string message, string title, MessageBoxButtons buttons, MessageBoxIcon icon)
            {
                using (var messageBox = new CustomMessageBox(message, title, buttons, icon))
                {
                    return messageBox.ShowDialog();
                }
            }
            
            public static DialogResult Show(string message, string title)
            {
                return Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.None);
            }
        }
    }
}
