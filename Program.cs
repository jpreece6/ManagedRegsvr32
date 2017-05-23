using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ManagedRegsvr32
{
    class Program
    {

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr LoadLibraryEx(string lpFileName, IntPtr hReservedNull, LoadLibraryFlags dwFlags);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
        static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        [DllImport("Kernel32", SetLastError = true)]
        static extern Int32 GetLastError();

        [DllImport("Ole32", SetLastError = true)]
        static extern Int32 OleInitialize(IntPtr hReservedNull);

        [DllImport("Ole32", SetLastError = true)]
        static extern void OleUninitialize();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [System.Flags]
        enum LoadLibraryFlags : uint
        {
            DONT_RESOLVE_DLL_REFERENCES = 0x00000001,
            LOAD_IGNORE_CODE_AUTHZ_LEVEL = 0x00000010,
            LOAD_LIBRARY_AS_DATAFILE = 0x00000002,
            LOAD_LIBRARY_AS_DATAFILE_EXCLUSIVE = 0x00000040,
            LOAD_LIBRARY_AS_IMAGE_RESOURCE = 0x00000020,
            LOAD_LIBRARY_SEARCH_APPLICATION_DIR = 0x00000200,
            LOAD_LIBRARY_SEARCH_DEFAULT_DIRS = 0x00001000,
            LOAD_LIBRARY_SEARCH_DLL_LOAD_DIR = 0x00000100,
            LOAD_LIBRARY_SEARCH_SYSTEM32 = 0x00000800,
            LOAD_LIBRARY_SEARCH_USER_DIRS = 0x00000400,
            LOAD_WITH_ALTERED_SEARCH_PATH = 0x00000008
        }

        private delegate IntPtr DllRegisterServer();
        private delegate IntPtr DllUnRegisterServer();

        private static bool _unregister;
        private static Options _options;

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        public enum ExitCode
        {
            Success = 0,
            Load_Failed = 1,
            DllRegisterServerNotFound = 2,
            DllUnRegisterServerNotFound = 3,
            FailedToCallRegisterMethod = 4,
            FailedToCallUnRegisterMethod = 5,
            InvalidArgument = 6,
            OleError = 7,
            ModulePlatform = 8
        }

        [STAThread]
        static void Main(string[] args)
        {
            var hwnd = GetConsoleWindow();
            ShowWindow(hwnd, SW_HIDE);

            _options = new Options();
            var valid = CommandLine.Parser.Default.ParseArgumentsStrict(args, _options);

            if (valid == false)
                Environment.Exit((int)ExitCode.InvalidArgument);

            if (File.Exists(_options.DllPath) == false)
            {
                ShowResult(ExitCode.Load_Failed);
                Environment.Exit((int)ExitCode.Load_Failed);
            }

            if (_options.ShowConsole)
            {
                ShowWindow(hwnd, SW_SHOW);
            }

            _unregister = _options.UnRegister;

            ExitCode exitResult = ExitCode.Load_Failed;
            if (_unregister)
            {
                exitResult = UnRegister(_options.DllPath);
            }
            else
            {
                exitResult = Register(_options.DllPath);
            }

            
            ShowResult(exitResult);

            Environment.Exit((int)exitResult);
        }

        private static void ShowResult(ExitCode code)
        {
            var msg = "";
            MessageBoxIcon icon = MessageBoxIcon.Error;
            switch (code)
            {
                case ExitCode.Load_Failed:
                    msg = "Failed to load Dll please check the path and ensure the Dll is a valid COM object.";
                    break;
                case ExitCode.DllRegisterServerNotFound:
                    msg = "Failed to find DllRegisterServer entry point please ensure the Dll is a valid COM object.";
                    break;
                case ExitCode.DllUnRegisterServerNotFound:
                    msg = "Failed to find DllRegisterServer entry point please ensure the Dll is a valid COM object.";
                    break;
                case ExitCode.FailedToCallRegisterMethod:
                    msg = "DllRegisterServer was found but a call to it could not be made.";
                    break;
                case ExitCode.FailedToCallUnRegisterMethod:
                    msg = "DllUnRegisterServer was found but a call to it could not be made.";
                    break;
                case ExitCode.InvalidArgument:
                    msg = "One or more invalid argument supplied please check the arguments.";
                    break;
                case ExitCode.OleError:
                    msg = "Oleinitalise command failed. You PC may be low on memory close open programs and try again.";
                    break;
                case ExitCode.ModulePlatform:
                    msg = "Check if the module is compatible with x86 or x64 version of ManagedRegsvr32";
                    break;
                case ExitCode.Success:
                    if (_unregister == false)
                    {
                        msg = "Dll was successfully registered!";
                    }
                    else
                    {
                        msg = "Dll was succesfully unregistered!";
                    }
                    icon = MessageBoxIcon.Information;
                    break;
                default:
                    msg = "How the hell did you break it?";
                    break;
            }

            Console.WriteLine(msg);

            if (_options.Silent == false)
                MessageBox.Show(msg, "ManagedRegsvr32", MessageBoxButtons.OK, icon);
        }

        public static ExitCode Register(string path)
        {
            var OleResult = OleInitialize(IntPtr.Zero);

            if (OleResult != 0)
                return ExitCode.OleError;

            IntPtr modulePtr = LoadLibraryEx(path, IntPtr.Zero, LoadLibraryFlags.LOAD_WITH_ALTERED_SEARCH_PATH);

            if (modulePtr == IntPtr.Zero)
            {
                OleUninitialize();

                return CheckLastError(ExitCode.Load_Failed);
            }

            IntPtr DllRegisterServerPtr = GetProcAddress(modulePtr, "DllRegisterServer");

            if (DllRegisterServerPtr == IntPtr.Zero)
            {
                OleUninitialize();
                FreeLibrary(modulePtr);
                return ExitCode.DllRegisterServerNotFound;
            }

            DllRegisterServer DllRegisterServerFunc = (DllRegisterServer) Marshal.GetDelegateForFunctionPointer(DllRegisterServerPtr, typeof(DllRegisterServer));

            var result = DllRegisterServerFunc();
            if (result == IntPtr.Zero)
            {
                OleUninitialize();
                FreeLibrary(modulePtr);
                return ExitCode.Success;
            }

            OleUninitialize();
            FreeLibrary(modulePtr);
            return ExitCode.FailedToCallRegisterMethod;
        }

        public static ExitCode UnRegister(string path)
        {
            var oleResult = OleInitialize(IntPtr.Zero);

            if (oleResult != 0)
                return ExitCode.OleError;

            IntPtr modulePtr = LoadLibraryEx(path, IntPtr.Zero, LoadLibraryFlags.LOAD_WITH_ALTERED_SEARCH_PATH);

            if (modulePtr == IntPtr.Zero)
                return ExitCode.Load_Failed;

            IntPtr DllUnregisterServerPtr = GetProcAddress(modulePtr, "DllUnregisterServer");

            if (DllUnregisterServerPtr == IntPtr.Zero)
            {
                OleUninitialize();
                FreeLibrary(modulePtr);
                return ExitCode.DllUnRegisterServerNotFound;
            }

            DllUnRegisterServer DllUnregisterServerFunc = (DllUnRegisterServer)Marshal.GetDelegateForFunctionPointer(DllUnregisterServerPtr, typeof(DllUnRegisterServer));

            var result = DllUnregisterServerFunc();
            if (result == IntPtr.Zero)
            {
                OleUninitialize();
                FreeLibrary(modulePtr);
                return ExitCode.Success;
            }

            OleUninitialize();
            FreeLibrary(modulePtr);
            return ExitCode.FailedToCallUnRegisterMethod;
        }

        private static ExitCode CheckLastError(ExitCode currentError = ExitCode.Success)
        {
            var result = GetLastError();
            switch(result)
            {
                case 193:
                    return ExitCode.ModulePlatform;
            }

            return currentError;
        }
    }
}
