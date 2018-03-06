using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Cache;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
using Newtonsoft.Json;

namespace Viramate {
    static partial class Program {
        [DllImport(
            "kernel32.dll", EntryPoint = "GetStdHandle", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern IntPtr GetStdHandle (int nStdHandle);

        [DllImport(
            "kernel32.dll", EntryPoint = "AllocConsole", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int AllocConsole ();

        [DllImport(
            "kernel32.dll", EntryPoint = "AttachConsole", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int AttachConsole (int processId);

        [DllImport(
            "kernel32.dll", EntryPoint = "SetConsoleTitle", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int SetConsoleTitle (string title);

        [DllImport(
            "kernel32.dll", EntryPoint = "TerminateProcess", 
            SetLastError = true, CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall
        )]
        public static extern int TerminateProcess (int processId, int exitCode);

        private const int ATTACH_PARENT_PROCESS = -1;
        private const int STD_INPUT_HANDLE = -10;
        private const int STD_OUTPUT_HANDLE = -11;
        private const int STD_ERROR_HANDLE = -12;
        private const int MY_CODE_PAGE = 437;

        public static Assembly MyAssembly;
        static bool IsRunningInsideCmd = false;

        public const string ExtensionSourceUrl = "http://luminance.org/vm/ext.zip";
        public const string InstallerSourceUrl = "http://luminance.org/vm/installer.zip";

        static void Main (string[] args) {
            MyAssembly = Assembly.GetExecutingAssembly();

            try {
                if ((args.Length <= 1) || !args.Any(a => a.StartsWith("chrome-extension://"))) {
                    InitConsole();
                    InstallExtension().Wait();
                } else
                    MessagingHostMainLoop();
            } catch (Exception exc) {
                Console.Error.WriteLine("Uncaught: {0}", exc);
                Environment.ExitCode = 1;

                // HACK: Assume the crash might be an installer bug, so try to install an update.
                AutoUpdateInstaller().Wait();
            }
        }

        static void InitConsole () {
            if (Debugger.IsAttached) {
                // No work necessary I think?
            } else {
                if (AttachConsole(ATTACH_PARENT_PROCESS) == 0) {
                    AllocConsole();
                    SetConsoleTitle("Viramate Installer");
                } else {
                    IsRunningInsideCmd = true;
                }

                IntPtr
                    stdinHandle = GetStdHandle(STD_INPUT_HANDLE),
                    stdoutHandle = GetStdHandle(STD_OUTPUT_HANDLE), 
                    stderrHandle = GetStdHandle(STD_ERROR_HANDLE);
                var stdinStream = new FileStream(new SafeFileHandle(stdinHandle, true), FileAccess.Read);
                var stdoutStream = new FileStream(new SafeFileHandle(stdoutHandle, true), FileAccess.Write);
                var stderrStream = new FileStream(new SafeFileHandle(stderrHandle, true), FileAccess.Write);
                var enc = new UTF8Encoding(false, false);
                var stdin = new StreamReader(stdinStream, enc);
                var stdout = new StreamWriter(stdoutStream, enc);
                var stderr = new StreamWriter(stderrStream, enc);
                stdout.AutoFlush = stderr.AutoFlush = true;
                Console.SetIn(stdin);
                Console.SetOut(stdout);
                Console.SetError(stderr);
            }
        }

        static string ExtensionId {
            get {
                /*
                var identity = WindowsIdentity.GetCurrent();
                var binLength = identity.User.BinaryLength;
                var bin = new byte[binLength];
                identity.User.GetBinaryForm(bin, 0);
                Array.Reverse(bin);

                var result = new char[32];
                Array.Copy("viramate".ToCharArray(), result, 8);
                for (int i = 0; i < binLength; i++) {
                    int j = i + 8;
                    if (j >= result.Length)
                        break;
                    result[j] = (char)((int)'a' + (bin[i] % 26));
                }
                return new string(result);
                */

                // FIXME: Generating a new extension ID requires manufacturing a public signing key... ugh
                return "fgpokpknehglcioijejfeebigdnbnokj";
            }
        }

        static bool IsRunningDirectlyFromBuild {
            get {
                return (Debugger.IsAttached || ExecutablePath.EndsWith("ViramateInstaller\\bin\\Viramate.exe"));
            }
        }

        static bool InstallFromDisk {
            get {
                var defaultDebug = IsRunningDirectlyFromBuild;
                var args = Environment.GetCommandLineArgs();
                if (args.Length <= 1)
                    return defaultDebug;

                return (defaultDebug || args.Contains("--disk")) && !args.Contains("--network");
            }
        }

        static string InstallPath {
            get {
                var folder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                return Path.Combine(folder, "Viramate");
            }
        }

        static string ExecutablePath {
            get {
                var cb = MyAssembly.CodeBase;
                var uri = new Uri(cb);
                return uri.LocalPath;
            }
        }

        static string ExecutableDirectory {
            get {
                return Path.GetDirectoryName(ExecutablePath);
            }
        }

        static void MessagingHostMainLoop () {
            var stdin = Console.OpenStandardInput();
            var stdout = Console.OpenStandardOutput();
            var logFilePath = Path.Combine(InstallPath, "installer.log");

            using (var log = new StreamWriter(logFilePath, true, Encoding.UTF8)) {
                Console.SetOut(log);
                Console.SetError(log);
                log.WriteLine($"{DateTime.UtcNow.ToLongTimeString()} > Installer started as native messaging host. Command line: {Environment.CommandLine}");
                WriteMessage(log, stdout, new { type = "serverStarting" });
                log.Flush();
                log.AutoFlush = true;

                try {
                    var wss = new WebSocketServer();
                    var t = wss.Run();
                    WriteMessage(log, stdout, new { type = "serverStarted", url = wss.Url });
                    t.Wait();
                } catch (Exception exc) {
                    log.WriteLine(exc);
                } finally {
                    log.WriteLine($"Exiting.");
                }

                log.Flush();
            }

            stdout.Flush();
            stdout.Close();
            stdin.Close();
        }

        public static string ReadManifestVersion (string zipFilePath) {
            var filename = Path.Combine(InstallPath, "manifest.json");
            try {
                string json;
                if (zipFilePath == null)
                    json = File.ReadAllText(filename, Encoding.UTF8);
                else {
                    using (var zf = new ZipArchive(File.OpenRead(zipFilePath), ZipArchiveMode.Read, false))
                    using (var fileStream = zf.Entries.FirstOrDefault(e => e.FullName.EndsWith("manifest.json")).Open())
                    using (var sr = new StreamReader(fileStream, Encoding.UTF8)) {
                        json = sr.ReadToEnd();
                    }
                }
                return JsonConvert.DeserializeObject<ManifestFragment>(json).version;
            } catch (Exception exc) {
                Console.Error.WriteLine(exc);
                return null;
            }
        }

        static T ReadMessage<T> (Stream stream)
            where T : class 
        {
            var inBuf = new byte[4096000];
            if (stream.Read(inBuf, 0, inBuf.Length) == 0)
                return null;

            var lengthBytes = (int)BitConverter.ToUInt32(inBuf, 0);
            var json = Encoding.UTF8.GetString(inBuf, 4, lengthBytes);
            var result = JsonConvert.DeserializeObject<T>(json);
            return result;
        }

        static void WriteMessage<T> (StreamWriter log, Stream stream, T message)
            where T : class
        {
            var json = JsonConvert.SerializeObject(message);
            var messageByteLength = Encoding.UTF8.GetByteCount(json);
            var messageBuf = new byte[messageByteLength + 4];
            Array.Copy(BitConverter.GetBytes((uint)messageByteLength), messageBuf, 4);
            Encoding.UTF8.GetBytes(json, 0, json.Length, messageBuf, 4);
            stream.Write(messageBuf, 0, messageBuf.Length);
            stream.Flush();
            log.WriteLine($"Wrote {messageByteLength} byte(s) of JSON: {json}");
        }

        class Msg {
            public string type;
        }

        class ManifestFragment {
            public string name;
            public string version;
        }
    }
}
