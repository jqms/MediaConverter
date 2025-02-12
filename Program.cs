using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Drawing;
using System.Security.Principal;

namespace FormatConverter {
    class Program {
        [DllImport("kernel32.dll")]
        private static extern bool FreeConsole();

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern int SHParseDisplayName(string name, IntPtr pbc, out IntPtr pidl, uint sfgaoIn, out uint psfgaoOut);

        [DllImport("shell32.dll")]
        public static extern int SHOpenFolderAndSelectItems(IntPtr pidlFolder, uint cidl, [In] IntPtr[] apidl, uint dwFlags);

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr ILFindChild(IntPtr pidlParent, IntPtr pidlChild);

        [DllImport("ole32.dll")]
        public static extern void CoTaskMemFree(IntPtr pv);

        private static readonly string[] AudioConversions = { "mp4", "mp3", "ogg", "opus", "wav" };

        private static readonly Dictionary<string, (string type, string[] conversions)> FormatMappings = new()
        {
            { ".mp4", ("video", new[] { "mp3", "mov" }) },
            { ".mov", ("video", new[] { "mp3", "mp4" }) },
            { ".mp3", ("audio", AudioConversions) },
            { ".ogg", ("audio", AudioConversions) },
            { ".opus", ("audio", AudioConversions) },
            { ".wav", ("audio", AudioConversions) }
        };

        [STAThread]
        static void Main(string[] args) {
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
                catch { }
                return;
            }
            FreeConsole();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length > 0 && args[0] == "-unregister") {
                UnregisterContextMenus();
                MessageBox.Show("Context menus removed successfully.", "Info");
                return;
            }
            if (args.Length != 2) {
                RegisterContextMenus();
                MessageBox.Show("Done! Right-click a file and select 'Convert To' to convert it.", "Info");
                return;
            }
            RegisterContextMenus();
            string inputFile = args[0];
            string outputFormat = args[1];
            string outputFile = Path.ChangeExtension(inputFile, outputFormat);
            try {
                PerformConversionWithProgress(inputFile, outputFile);
                OpenFolderAndSelectFile(outputFile);
            }
            catch (Exception ex) {
                MessageBox.Show($"Error during conversion: {ex.Message}", "Error");
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
                using var mainKey = Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To");
                mainKey.SetValue("SubCommands", "");
                mainKey.SetValue("MUIVerb", "Convert To");
                using var shellKey = Registry.ClassesRoot.CreateSubKey($@"SystemFileAssociations\{extension}\shell\Convert To\shell");
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
        }

        static void PerformConversionWithProgress(string input, string output) {
            using var progressForm = new ProgressForm();
            var conversionTask = Task.Run(() => {
                ConvertMedia(input, output);
            });
            conversionTask.ContinueWith(t => {
                progressForm.Invoke(new Action(() => {
                    progressForm.Close();
                }));
            });
            Application.Run(progressForm);
            conversionTask.Wait();
            if (conversionTask.IsFaulted && conversionTask.Exception != null)
                throw conversionTask.Exception.InnerException;
        }

        static void ConvertMedia(string input, string output) {
            string inputExt = Path.GetExtension(input).ToLower();
            string outputExt = Path.GetExtension(output).ToLower();
            var (inputType, _) = FormatMappings[inputExt];
            string arguments;
            switch (outputExt) {
                case ".mp3":
                    if (inputType == "video")
                        arguments = $"-i \"{input}\" -vn -acodec libmp3lame -q:a 2 \"{output}\"";
                    else
                        arguments = $"-i \"{input}\" -acodec libmp3lame -q:a 2 \"{output}\"";
                    break;
                case ".mp4":
                    if (inputType == "audio")
                        arguments = $"-i \"{input}\" -f lavfi -i color=c=black:s=1920x1080 -shortest -c:v libx264 -c:a aac \"{output}\"";
                    else
                        arguments = $"-i \"{input}\" -c:v libx264 -c:a aac \"{output}\"";
                    break;
                case ".opus":
                    arguments = $"-i \"{input}\" -c:a libopus -b:a 128k \"{output}\"";
                    break;
                case ".ogg":
                    arguments = $"-i \"{input}\" -c:a libvorbis -q:a 5 \"{output}\"";
                    break;
                case ".wav":
                    arguments = $"-i \"{input}\" -acodec pcm_s16le \"{output}\"";
                    break;
                default:
                    arguments = $"-i \"{input}\" \"{output}\"";
                    break;
            }
            var startInfo = new ProcessStartInfo {
                FileName = "ffmpeg",
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true
            };
            using var process = Process.Start(startInfo);
            if (process == null)
                throw new Exception("Failed to start ffmpeg process.");
            string errorOutput = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0) {
                throw new Exception($"FFmpeg exited with code {process.ExitCode}. Error:\n{errorOutput}");
            }
        }

        static void OpenFolderAndSelectFile(string filePath) {
            string folder = Path.GetDirectoryName(filePath);
            bool usedExisting = false;
            try {
                Type shellAppType = Type.GetTypeFromProgID("Shell.Application");
                dynamic shell = Activator.CreateInstance(shellAppType);
                foreach (dynamic window in shell.Windows()) {
                    try {
                        string windowPath = new Uri((string)window.LocationURL).LocalPath;
                        if (string.Equals(windowPath.TrimEnd('\\'), folder.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)) {
                            window.Document.SelectItem(filePath, 0);
                            usedExisting = true;
                            break;
                        }
                    }
                    catch { continue; }
                }
            }
            catch { }
            if (!usedExisting) {
                IntPtr filePidl;
                uint sfgao;
                int hr = SHParseDisplayName(filePath, IntPtr.Zero, out filePidl, 0, out sfgao);
                if (hr != 0)
                    throw new Exception("Failed to parse file path.");
                IntPtr folderPidl;
                hr = SHParseDisplayName(folder, IntPtr.Zero, out folderPidl, 0, out sfgao);
                if (hr != 0) {
                    CoTaskMemFree(filePidl);
                    throw new Exception("Failed to parse folder path.");
                }
                IntPtr relativePidl = ILFindChild(folderPidl, filePidl);
                if (relativePidl == IntPtr.Zero)
                    relativePidl = filePidl;
                IntPtr[] items = new IntPtr[] { relativePidl };
                hr = SHOpenFolderAndSelectItems(folderPidl, (uint)items.Length, items, 0);
                CoTaskMemFree(filePidl);
                CoTaskMemFree(folderPidl);
                if (hr != 0)
                    throw new Exception("Failed to open folder and select file.");
            }
        }
    }

    class ProgressForm : Form {
        private ProgressBar progressBar;
        public ProgressForm() {
            this.TopMost = true;
            this.ClientSize = new Size(210, 40);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ControlBox = false;
            progressBar = new ProgressBar {
                Style = ProgressBarStyle.Marquee,
                MarqueeAnimationSpeed = 30,
                Width = 200,
                Height = 15,
                Location = new Point(5, 12)
            };
            this.Controls.Add(progressBar);
        }
    }
}