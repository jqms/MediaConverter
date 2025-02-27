using Microsoft.Win32;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;

namespace FormatConverter
{
    class Program
    {
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

        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length == 0)
            {
                if (!IsAdministrator())
                {
                    var procInfo = new ProcessStartInfo
                    {
                        UseShellExecute = true,
                        WorkingDirectory = Environment.CurrentDirectory,
                        FileName = Application.ExecutablePath,
                        Verb = "runas"
                    };
                    try
                    {
                        Process.Start(procInfo);
                    }
                    catch
                    {
                    }

                    return;
                }

                RegisterContextMenus();
                MessageBox.Show("Done! Right-click a file and select 'Convert To' to convert it.", "Info");
                return;
            }

            if (args.Length == 1 && args[0] == "-unregister")
            {
                if (!IsAdministrator())
                {
                    var procInfo = new ProcessStartInfo
                    {
                        UseShellExecute = true,
                        WorkingDirectory = Environment.CurrentDirectory,
                        FileName = Application.ExecutablePath,
                        Verb = "runas"
                    };
                    try
                    {
                        Process.Start(procInfo);
                    }
                    catch
                    {
                    }

                    return;
                }

                UnregisterContextMenus();
                MessageBox.Show("Context menus removed successfully.", "Info");
                return;
            }

            if (args.Length == 2 && args[1] == "mute")
            {
                string inputFile = args[0];
                string extension = Path.GetExtension(inputFile);
                string outputFile = Path.Combine(
                    Path.GetDirectoryName(inputFile),
                    Path.GetFileNameWithoutExtension(inputFile) + "_muted" + extension);

                try
                {
                    if (File.Exists(outputFile))
                    {
                        var result = MessageBox.Show(
                            $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                            "File Exists",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Cancel || result == DialogResult.No)
                            return;

                        try
                        {
                            File.Delete(outputFile);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                            return;
                        }
                    }

                    PerformMuteWithProgress(inputFile, outputFile);
                    OpenFolderAndSelectFile(outputFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error during muting: {ex.Message}", "Error");
                    Environment.Exit(1);
                }

                return;
            }

            if (args.Length == 2 && (args[1] == "resize50" || args[1] == "resize75"))
            {
                string inputFile = args[0];
                string extension = Path.GetExtension(inputFile);
                string scale = args[1] == "resize50" ? "50%" : "75%";
                string scaleValue = args[1] == "resize50" ? "0.5" : "0.75";
                string outputFile = Path.Combine(
                    Path.GetDirectoryName(inputFile),
                    Path.GetFileNameWithoutExtension(inputFile) + "_" + scale.Replace("%", "pct") + extension);

                try
                {
                    if (File.Exists(outputFile))
                    {
                        var result = MessageBox.Show(
                            $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                            "File Exists",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Warning);

                        if (result == DialogResult.Cancel || result == DialogResult.No)
                            return;

                        try
                        {
                            File.Delete(outputFile);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                            return;
                        }
                    }

                    PerformResizeWithProgress(inputFile, outputFile, scaleValue);
                    OpenFolderAndSelectFile(outputFile);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error during resizing: {ex.Message}", "Error");
                    Environment.Exit(1);
                }

                return;
            }

            if (args.Length > 1)
            {
                if (args.Length > 2)
                {
                    string outputFormat = args.Last();
                    try
                    {
                        foreach (var inputFile in args.Take(args.Length - 1))
                        {
                            string outputFile = Path.ChangeExtension(inputFile, outputFormat);
                            if (File.Exists(outputFile))
                            {
                                var result = MessageBox.Show(
                                    $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                                    "File Exists",
                                    MessageBoxButtons.YesNoCancel,
                                    MessageBoxIcon.Warning);

                                if (result == DialogResult.Cancel)
                                    return;
                                if (result == DialogResult.No)
                                    continue;
                                try
                                {
                                    File.Delete(outputFile);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                                    continue;
                                }
                            }

                            PerformConversionWithProgress(inputFile, outputFile);
                        }

                        MessageBox.Show("Batch conversion completed successfully.", "Info");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error during batch conversion: {ex.Message}", "Error");
                        Environment.Exit(1);
                    }
                }
                else
                {
                    string inputFile = args[0];
                    string outputFormat = args[1];
                    string outputFile = Path.ChangeExtension(inputFile, outputFormat);
                    try
                    {
                        if (File.Exists(outputFile))
                        {
                            var result = MessageBox.Show(
                                $"File already exists:\n{outputFile}\n\nDo you want to replace it?",
                                "File Exists",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Warning);

                            if (result == DialogResult.Cancel || result == DialogResult.No)
                                return;

                            try
                            {
                                File.Delete(outputFile);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Cannot replace existing file: {ex.Message}", "Error");
                                return;
                            }
                        }

                        PerformConversionWithProgress(inputFile, outputFile);
                        OpenFolderAndSelectFile(outputFile);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error during conversion: {ex.Message}", "Error");
                        Environment.Exit(1);
                    }
                }
            }
        }

        static bool IsAdministrator()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        static void RegisterContextMenus()
        {
            string execPath = Application.ExecutablePath;
            foreach (var mapping in FormatMappings)
            {
                string extension = mapping.Key;
                var (type, conversions) = mapping.Value;

                using var mainKey =
                    Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To");
                mainKey.SetValue("SubCommands", "");
                mainKey.SetValue("MUIVerb", "Convert To");
                using var shellKey =
                    Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To\shell");

                using var toolsKey = shellKey.CreateSubKey("!Tools");
                toolsKey.SetValue("MUIVerb", "Tools");
                toolsKey.SetValue("SubCommands", "");

                using var toolsShellKey = toolsKey.CreateSubKey("shell");

                if (type == "audio" || type == "video")
                {
                    using var muteKey = toolsShellKey.CreateSubKey("Mute");
                    muteKey.SetValue("", "Mute");
                    using var muteCmdKey = muteKey.CreateSubKey("command");
                    muteCmdKey.SetValue("", $"\"{execPath}\" \"%1\" mute");
                }

                if (type == "image")
                {
                    using var resizeKey = toolsShellKey.CreateSubKey("Resize50");
                    resizeKey.SetValue("", "Resize to 50%");
                    using var resizeCmdKey = resizeKey.CreateSubKey("command");
                    resizeCmdKey.SetValue("", $"\"{execPath}\" \"%1\" resize50");

                    using var resize75Key = toolsShellKey.CreateSubKey("Resize75");
                    resize75Key.SetValue("", "Resize to 75%");
                    using var resize75CmdKey = resize75Key.CreateSubKey("command");
                    resize75CmdKey.SetValue("", $"\"{execPath}\" \"%1\" resize75");
                }

                foreach (var format in conversions)
                {
                    if (format == extension.TrimStart('.'))
                        continue;
                    using var formatKey = shellKey.CreateSubKey(format.ToUpper());
                    formatKey.SetValue("", format.ToUpper());
                    using var cmdKey = formatKey.CreateSubKey("command");
                    cmdKey.SetValue("", $"\"{execPath}\" \"%1\" {format}");
                }
            }
        }

        static void UnregisterContextMenus()
        {
            foreach (var mapping in FormatMappings)
            {
                string extension = mapping.Key;
                try
                {
                    Registry.ClassesRoot.DeleteSubKeyTree($@"SystemFileAssociations\{extension}\shell\Convert To",
                        false);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error removing menu for {extension}: {ex.Message}", "Error");
                }
            }
        }

        static void PerformConversionWithProgress(string input, string output)
        {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing)
            {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try
                {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var conversionTask = Task.Run(() =>
            {

                string inputExt = Path.GetExtension(input).ToLower();
                string outputExt = Path.GetExtension(output).ToLower();
                var (inputType, _) = FormatMappings[inputExt];
                string arguments = "";

                if (inputType == "image")
                {
                    switch (outputExt)
                    {
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
                            if (inputExt == ".gif")
                            {
                                arguments =
                                    $"-i \"{input}\" -lavfi \"fps=15,scale=trunc(iw/2)*2:trunc(ih/2)*2:flags=lanczos\" \"{output}\"";
                            }
                            else
                            {
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
                else if (inputType == "audio")
                {
                    switch (outputExt)
                    {
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
                else if (inputType == "video")
                {
                    string videoCodec, audioCodec, extraParams = "";

                    switch (outputExt)
                    {
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

                    if (inputType == "audio" && outputExt != ".mp3" && outputExt != ".m4a")
                    {
                        arguments =
                            $"-i \"{input}\" -f lavfi -i color=c=black:s=1920x1080 -shortest {videoCodec} {audioCodec} {extraParams} \"{output}\"";
                    }
                }

                var startInfo = new ProcessStartInfo
                {
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

                ffmpegProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data))
                        return;

                    string data = e.Data;

                    if (duration == TimeSpan.Zero)
                    {
                        var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                        if (durationMatch.Success)
                        {
                            int hours = int.Parse(durationMatch.Groups[1].Value);
                            int minutes = int.Parse(durationMatch.Groups[2].Value);
                            int seconds = int.Parse(durationMatch.Groups[3].Value);
                            int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                            duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                        }
                    }

                    var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                    if (timeMatch.Success)
                    {
                        int hours = int.Parse(timeMatch.Groups[1].Value);
                        int minutes = int.Parse(timeMatch.Groups[2].Value);
                        int seconds = int.Parse(timeMatch.Groups[3].Value);
                        int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                        currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                        if (duration != TimeSpan.Zero)
                        {
                            int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                            percentage = Math.Min(99, Math.Max(0, percentage));

                            var forms = Application.OpenForms;
                            foreach (Form form in forms)
                            {
                                if (form is ProgressForm progressForm)
                                {
                                    progressForm.BeginInvoke(new Action(() =>
                                    {
                                        progressForm.UpdateProgress(percentage);
                                    }));
                                    break;
                                }
                            }
                        }
                    }
                };

                progressForm.Invoke(new Action(() =>
                {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled())
                {
                    string errorSummary = "FFmpeg failed to complete the operation.";
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}. {errorSummary}");
                }

                if (isReplacing && File.Exists(backupPath))
                {
                    try
                    {
                        File.Delete(backupPath);
                    }
                    catch
                    {
                    }
                }
            });

            conversionTask.ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    progressForm.Invoke(new Action(() =>
                    {
                        if (!progressForm.WasCancelled())
                        {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else
                        {
                            progressForm.Close();
                        }
                    }));
                }
                else
                {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try
            {
                conversionTask.Wait();
            }
            catch (AggregateException ae)
            {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void PerformMuteWithProgress(string input, string output)
        {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing)
            {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try
                {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var muteTask = Task.Run(() =>
            {

                string inputExt = Path.GetExtension(input).ToLower();
                var (inputType, _) = FormatMappings[inputExt];
                string arguments;

                if (inputType == "video")
                {
                    arguments = $"-i \"{input}\" -c:v copy -c:a aac -af \"volume=0\" \"{output}\"";
                }
                else if (inputType == "audio")
                {
                    arguments = $"-i \"{input}\" -af \"volume=0\" \"{output}\"";
                }
                else
                {
                    throw new Exception("Mute operation only supports audio and video files.");
                }

                var startInfo = new ProcessStartInfo
                {
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

                ffmpegProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data))
                        return;

                    string data = e.Data;

                    if (duration == TimeSpan.Zero)
                    {
                        var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                        if (durationMatch.Success)
                        {
                            int hours = int.Parse(durationMatch.Groups[1].Value);
                            int minutes = int.Parse(durationMatch.Groups[2].Value);
                            int seconds = int.Parse(durationMatch.Groups[3].Value);
                            int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                            duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                        }
                    }

                    var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                    if (timeMatch.Success)
                    {
                        int hours = int.Parse(timeMatch.Groups[1].Value);
                        int minutes = int.Parse(timeMatch.Groups[2].Value);
                        int seconds = int.Parse(timeMatch.Groups[3].Value);
                        int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                        currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                        if (duration != TimeSpan.Zero)
                        {
                            int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                            percentage = Math.Min(99, Math.Max(0, percentage));

                            var forms = Application.OpenForms;
                            foreach (Form form in forms)
                            {
                                if (form is ProgressForm progressForm)
                                {
                                    progressForm.BeginInvoke(new Action(() =>
                                    {
                                        progressForm.UpdateProgress(percentage);
                                    }));
                                    break;
                                }
                            }
                        }
                    }
                };

                progressForm.Invoke(new Action(() =>
                {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled())
                {
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}");
                }

                if (isReplacing && File.Exists(backupPath))
                {
                    try
                    {
                        File.Delete(backupPath);
                    }
                    catch
                    {
                    }
                }
            });

            muteTask.ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    progressForm.Invoke(new Action(() =>
                    {
                        if (!progressForm.WasCancelled())
                        {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else
                        {
                            progressForm.Close();
                        }
                    }));
                }
                else
                {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try
            {
                muteTask.Wait();
            }
            catch (AggregateException ae)
            {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void PerformResizeWithProgress(string input, string output, string scale)
        {

            string backupPath = null;
            bool isReplacing = File.Exists(output);

            if (isReplacing)
            {
                backupPath = Path.Combine(
                    Path.GetDirectoryName(output),
                    Path.GetFileNameWithoutExtension(output) + ".backup" + Path.GetExtension(output)
                );

                try
                {
                    File.Copy(output, backupPath, true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Warning: Could not create backup: {ex.Message}", "Backup Warning");
                    backupPath = null;
                }
            }

            using var progressForm = new ProgressForm();
            Process ffmpegProcess = null;

            var resizeTask = Task.Run(() =>
            {

                string inputExt = Path.GetExtension(input).ToLower();
                var (inputType, _) = FormatMappings[inputExt];

                if (inputType != "image")
                {
                    throw new Exception("Resize operation only supports image files.");
                }

                string arguments = $"-i \"{input}\" -vf \"scale=iw*{scale}:ih*{scale}\" \"{output}\"";

                var startInfo = new ProcessStartInfo
                {
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

                progressForm.Invoke(new Action(() =>
                {
                    progressForm.SetProcessInfo(ffmpegProcess, output, isReplacing, backupPath);
                }));

                ffmpegProcess.Start();
                ffmpegProcess.BeginErrorReadLine();
                ffmpegProcess.WaitForExit();

                if (ffmpegProcess.ExitCode != 0 && !progressForm.WasCancelled())
                {
                    throw new Exception($"FFmpeg exited with code {ffmpegProcess.ExitCode}");
                }

                if (isReplacing && File.Exists(backupPath))
                {
                    try
                    {
                        File.Delete(backupPath);
                    }
                    catch
                    {
                    }
                }
            });

            resizeTask.ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    progressForm.Invoke(new Action(() =>
                    {
                        if (!progressForm.WasCancelled())
                        {
                            progressForm.SetError(t.Exception?.InnerException?.Message ?? "Unknown error");
                        }
                        else
                        {
                            progressForm.Close();
                        }
                    }));
                }
                else
                {
                    progressForm.Invoke(new Action(() => { progressForm.SetCompleted(); }));
                }
            });

            Application.Run(progressForm);

            try
            {
                resizeTask.Wait();
            }
            catch (AggregateException ae)
            {
                if (ae.InnerException != null)
                    throw ae.InnerException;
                throw;
            }
        }

        static void MuteMedia(string input, string output)
        {
            string inputExt = Path.GetExtension(input).ToLower();
            var (inputType, _) = FormatMappings[inputExt];
            string arguments;
            if (inputType == "video")
            {
                arguments = $"-i \"{input}\" -c:v copy -c:a aac -af \"volume=0\" \"{output}\"";
            }
            else if (inputType == "audio")
            {
                arguments = $"-i \"{input}\" -af \"volume=0\" \"{output}\"";
            }
            else
            {
                throw new Exception("Mute operation only supports audio and video files.");
            }

            var startInfo = new ProcessStartInfo
            {
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

            process.ErrorDataReceived += (sender, e) =>
            {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero)
                {
                    var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success)
                    {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success)
                {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero)
                    {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms)
                        {
                            if (form is ProgressForm progressForm)
                            {
                                progressForm.BeginInvoke(new Action(() =>
                                {
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

            if (process.ExitCode != 0)
            {
                throw new Exception($"FFmpeg exited with code {process.ExitCode}");
            }
        }

        static void ResizeImage(string input, string output, string scale)
        {
            string inputExt = Path.GetExtension(input).ToLower();
            var (inputType, _) = FormatMappings[inputExt];

            if (inputType != "image")
            {
                throw new Exception("Resize operation only supports image files.");
            }

            string arguments = $"-i \"{input}\" -vf \"scale=iw*{scale}:ih*{scale}\" \"{output}\"";

            var startInfo = new ProcessStartInfo
            {
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

            process.ErrorDataReceived += (sender, e) =>
            {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero)
                {
                    var durationMatch = Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success)
                    {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success)
                {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero)
                    {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms)
                        {
                            if (form is ProgressForm progressForm)
                            {
                                progressForm.BeginInvoke(new Action(() =>
                                {
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

            if (process.ExitCode != 0)
            {
                throw new Exception($"FFmpeg exited with code {process.ExitCode}");
            }
        }

        static void ConvertMedia(string input, string output)
        {
            string inputExt = Path.GetExtension(input).ToLower();
            string outputExt = Path.GetExtension(output).ToLower();
            var (inputType, _) = FormatMappings[inputExt];
            string arguments = "";

            if (inputType == "image")
            {
                switch (outputExt)
                {
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
                        if (inputExt == ".gif")
                        {

                            arguments =
                                $"-i \"{input}\" -lavfi \"fps=15,scale=trunc(iw/2)*2:trunc(ih/2)*2:flags=lanczos\" \"{output}\"";
                        }
                        else
                        {
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
            else if (inputType == "audio")
            {
                switch (outputExt)
                {
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
            else if (inputType == "video")
            {
                string videoCodec, audioCodec, extraParams = "";

                switch (outputExt)
                {
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

                if (inputType == "audio" && outputExt != ".mp3" && outputExt != ".m4a")
                {
                    arguments =
                        $"-i \"{input}\" -f lavfi -i color=c=black:s=1920x1080 -shortest {videoCodec} {audioCodec} {extraParams} \"{output}\"";
                }
            }

            var startInfo = new ProcessStartInfo
            {
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

            process.ErrorDataReceived += (sender, e) =>
            {
                if (string.IsNullOrEmpty(e.Data))
                    return;

                string data = e.Data;

                if (duration == TimeSpan.Zero)
                {
                    var durationMatch =
                        System.Text.RegularExpressions.Regex.Match(data, @"Duration: (\d+):(\d+):(\d+)\.(\d+)");
                    if (durationMatch.Success)
                    {
                        int hours = int.Parse(durationMatch.Groups[1].Value);
                        int minutes = int.Parse(durationMatch.Groups[2].Value);
                        int seconds = int.Parse(durationMatch.Groups[3].Value);
                        int milliseconds = int.Parse(durationMatch.Groups[4].Value) * 10;
                        duration = new TimeSpan(0, hours, minutes, seconds, milliseconds);
                    }
                }

                var timeMatch = System.Text.RegularExpressions.Regex.Match(data, @"time=(\d+):(\d+):(\d+)\.(\d+)");
                if (timeMatch.Success)
                {
                    int hours = int.Parse(timeMatch.Groups[1].Value);
                    int minutes = int.Parse(timeMatch.Groups[2].Value);
                    int seconds = int.Parse(timeMatch.Groups[3].Value);
                    int milliseconds = int.Parse(timeMatch.Groups[4].Value) * 10;
                    currentTime = new TimeSpan(0, hours, minutes, seconds, milliseconds);

                    if (duration != TimeSpan.Zero)
                    {
                        int percentage = (int)((currentTime.TotalMilliseconds / duration.TotalMilliseconds) * 100);
                        percentage = Math.Min(99, Math.Max(0, percentage));

                        var forms = Application.OpenForms;
                        foreach (Form form in forms)
                        {
                            if (form is ProgressForm progressForm)
                            {
                                progressForm.BeginInvoke(new Action(() =>
                                {
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

            if (process.ExitCode != 0)
            {
                string errorSummary = "FFmpeg failed to complete the operation.";
                throw new Exception($"FFmpeg exited with code {process.ExitCode}. {errorSummary}");
            }
        }

        static void OpenFolderAndSelectFile(string filePath)
        {
            try
            {
                string folder = Path.GetDirectoryName(filePath);

                if (!File.Exists(filePath))
                {

                    int attempts = 0;
                    while (!File.Exists(filePath) && attempts < 10)
                    {
                        System.Threading.Thread.Sleep(100);
                        attempts++;
                    }

                    if (!File.Exists(filePath))
                    {
                        return;
                    }
                }

                bool usedExisting = false;

                try
                {
                    Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                    dynamic shell = Activator.CreateInstance(shellAppType);
                    foreach (dynamic window in shell.Windows())
                    {
                        try
                        {
                            string windowPath = new Uri((string)window.LocationURL).LocalPath;
                            if (string.Equals(windowPath.TrimEnd('\\'), folder.TrimEnd('\\'),
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                window.Document.SelectItem(filePath, 0);
                                usedExisting = true;
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
                catch
                {

                }

                if (!usedExisting)
                {

                    System.Threading.Thread.Sleep(200);

                    try
                    {
                        IntPtr filePidl;
                        uint sfgao;
                        int hr = SHParseDisplayName(filePath, IntPtr.Zero, out filePidl, 0, out sfgao);
                        if (hr != 0)
                        {

                            Process.Start("explorer.exe", folder);
                            return;
                        }

                        IntPtr folderPidl;
                        hr = SHParseDisplayName(folder, IntPtr.Zero, out folderPidl, 0, out sfgao);
                        if (hr != 0)
                        {
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

                        if (hr != 0)
                        {

                        }
                    }
                    catch
                    {

                    }
                }
            }
            catch
            {

                try
                {
                    string folder = Path.GetDirectoryName(filePath);
                }
                catch
                {

                }
            }
        }

        class ProgressForm : Form
        {
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
                string originalBackupPath = null)
            {
                currentProcess = process;
                outputFilePath = outputPath;
                isReplacement = isReplacingFile;
                backupFilePath = originalBackupPath;
            }

            private void FormDragMouseDown(object sender, MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            }

            public ProgressForm()
            {
                this.Text = "Converting...";
                this.TopMost = true;
                this.ClientSize = new Size(360, 100);
                this.FormBorderStyle = FormBorderStyle.None;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.ControlBox = false;

                this.BackColor = Color.Black;

                this.Region = Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 8, 8));

                percentLabel = new Label
                {
                    Text = "0%",
                    Width = 336,
                    Height = 20,
                    Location = new Point(12, 12),
                    Font = new Font("Segoe UI", 14F, FontStyle.Regular),
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = Color.White,
                    BackColor = Color.Transparent
                };

                progressBarContainer = new Panel
                {
                    Width = 336,
                    Height = 6,
                    Location = new Point(12, 40),
                    BackColor = Color.FromArgb(26, 26, 26),
                };
                progressBarContainer.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarContainer.Width, progressBarContainer.Height, 3,
                        3));

                progressBarFill = new Panel
                {
                    Width = 0,
                    Height = 6,
                    Location = new Point(0, 0),
                    BackColor = Color.FromArgb(0, 114, 255)
                };
                progressBarFill.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarFill.Width, progressBarFill.Height, 3, 3));
                progressBarContainer.Controls.Add(progressBarFill);

                statusLabel = new Label
                {
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

                Label closeButton = new Label
                {
                    Text = "✕",
                    Size = new Size(24, 24),
                    Location = new Point(Width - 30, 8),
                    Font = new Font("Arial", 10F, FontStyle.Bold),
                    ForeColor = Color.Gray,
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = Color.Transparent
                };
                closeButton.Click += (s, e) =>
                {
                    DialogResult result = MessageBox.Show(
                        "Are you sure you want to cancel the operation?",
                        "Cancel",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
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

                this.FormClosing += (s, e) =>
                {
                    if (currentProcess != null && !currentProcess.HasExited)
                    {
                        DialogResult result = MessageBox.Show(
                            "Are you sure you want to cancel the operation?",
                            "Cancel",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            HandleCancellation();
                        }
                        else
                        {
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
                animationTimer.Tick += (sender, e) =>
                {
                    if (Math.Abs(currentProgress - targetProgress) > 0.1)
                    {
                        currentProgress += (targetProgress - currentProgress) * ANIMATION_SPEED;
                        UpdateVisualProgress((int)Math.Round(currentProgress));
                    }
                };
                animationTimer.Start();

                progressBarFill.Paint += ProgressBarFill_Paint;
            }

            private bool wasCancelled = false;

            public bool WasCancelled()
            {
                return wasCancelled;
            }

            private void HandleCancellation()
            {
                try
                {
                    wasCancelled = true;

                    if (currentProcess != null && !currentProcess.HasExited)
                    {
                        try
                        {
                            currentProcess.Kill(true);
                        }
                        catch
                        {

                            currentProcess.Kill();
                        }

                        currentProcess = null;
                    }

                    if (!string.IsNullOrEmpty(outputFilePath))
                    {
                        try
                        {

                            if (File.Exists(outputFilePath))
                            {
                                File.Delete(outputFilePath);
                            }

                            if (isReplacement && !string.IsNullOrEmpty(backupFilePath) && File.Exists(backupFilePath))
                            {
                                File.Move(backupFilePath, outputFilePath, true);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Warning: Could not clean up files: {ex.Message}", "Warning",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error during cancellation: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void ProgressBarFill_Paint(object sender, PaintEventArgs e)
            {
                if (isErrorState)
                {
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                               progressBarFill.ClientRectangle,
                               Color.FromArgb(255, 100, 100),
                               Color.FromArgb(200, 50, 50),
                               LinearGradientMode.Horizontal))
                    {
                        e.Graphics.FillRectangle(brush, progressBarFill.ClientRectangle);
                    }
                }
                else
                {
                    using (LinearGradientBrush brush = new LinearGradientBrush(
                               progressBarFill.ClientRectangle,
                               Color.FromArgb(0, 198, 255),
                               Color.FromArgb(0, 114, 255),
                               LinearGradientMode.Horizontal))
                    {
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
                               LinearGradientMode.Horizontal))
                    {
                        e.Graphics.FillRectangle(shimmerBrush,
                            new Rectangle(shimmerPosition, 0, shimmerWidth, progressBarFill.Height));
                    }
                }
            }

            private void UpdateVisualProgress(int value)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<int>(UpdateVisualProgress), value);
                    return;
                }

                int fillWidth = (int)((progressBarContainer.Width * value) / 100.0);
                progressBarFill.Width = fillWidth;
                progressBarFill.Region =
                    Region.FromHrgn(CreateRoundRectRgn(0, 0, progressBarFill.Width, progressBarFill.Height, 3, 3));
            }

            public void UpdateProgress(int value)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<int>(UpdateProgress), value);
                    return;
                }

                targetProgress = value;
                percentLabel.Text = $"{value}%";

                if (value >= 100)
                {
                    statusLabel.Text = "Operation completed successfully!";
                }
                else if (value >= 75)
                {
                    statusLabel.Text = "Finalizing...";
                }
                else if (value >= 50)
                {
                    statusLabel.Text = "Processing...";
                }
                else if (value >= 25)
                {
                    statusLabel.Text = "Converting...";
                }
                else
                {
                    statusLabel.Text = "Preparing...";
                }
            }

            public void SetStatusText(string status)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string>(SetStatusText), status);
                    return;
                }

                statusLabel.Text = status;
            }

            public void SetCompleted()
            {
                if (InvokeRequired)
                {
                    Invoke(new Action(SetCompleted));
                    return;
                }

                UpdateProgress(100);
                statusLabel.Text = "Operation completed successfully!";
                System.Threading.Thread.Sleep(500);
                Close();
            }

            public void SetError(string errorMessage)
            {
                if (InvokeRequired)
                {
                    Invoke(new Action<string>(SetError), errorMessage);
                    return;
                }

                statusLabel.Text = $"Error: {errorMessage}";
                statusLabel.ForeColor = Color.FromArgb(255, 100, 100);
                isErrorState = true;
                progressBarFill.Invalidate();
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    shimmerTimer?.Dispose();
                    animationTimer?.Dispose();
                }

                base.Dispose(disposing);
            }
        }
    }
}
