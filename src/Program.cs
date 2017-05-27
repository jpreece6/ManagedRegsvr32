using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Linq;

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
        private static bool _silent;

        public enum ExitCode
        {
            Success = 0,
            Load_Failed = 1,
            DllRegisterServerNotFound = 2,
            DllUnRegisterServerNotFound = 3,
            FailedToCallRegisterMethod = 4,
            FailedToCallUnRegisterMethod = 5,
            InvalidArguments = 6,
            OleError = 7,
            ModulePlatform = 8,
        }

        [STAThread]
        static void Main(string[] args)
        {
            // Bomb out here if the invoker failed to provide arguments
            if (args.Length == 0)
            {
                ShowHelp();
                Environment.Exit((int) ExitCode.InvalidArguments);
            }

            // Initialise OLE
            if (OleInitialize(IntPtr.Zero) != 0)
                Environment.Exit((int) ExitCode.OleError);

            _unregister = false;
            _silent = false;

            // Process each argument available
            ExitCode exitResult = ExitCode.Load_Failed;
            foreach (var arg in args)
            {
                // Support - and / prefixes
                if (arg.Contains('-') || arg.Contains('/'))
                {
                    switch(arg)
                    {
                        case "-s":
                            _silent = true;
                            break;
                        case "-u":
                            _unregister = true;
                            break;
                    }
                    
                    continue;
                }


                var modulePath = arg;

                if (File.Exists(modulePath) == false)
                {
                    ShowResult(ExitCode.Load_Failed);
                    OleUninitialize();
                    Environment.Exit((int)ExitCode.Load_Failed);
                }

                if (_unregister)
                {
                    exitResult = UnRegister(modulePath);
                }
                else
                {
                    exitResult = Register(modulePath);
                }

                ShowResult(exitResult);

                if (exitResult != ExitCode.Success)
                {
                    OleUninitialize();
                    Environment.Exit((int)exitResult);
                }
            }

            // Clean up and report error code
            OleUninitialize();
            Environment.Exit((int)exitResult);
        }

        /// <summary>
        /// Displays a detailed description of the error code. Does not show if -s has been used
        /// </summary>
        /// <param name="code">Code to detail</param>
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
                case ExitCode.InvalidArguments:
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

            if (_silent == false)
            {

                Console.WriteLine(msg);
                MessageBox.Show(msg, "ManagedRegsvr32", MessageBoxButtons.OK, icon);
            }
        }

        private static ExitCode CallDllUnRegisterServer(IntPtr modulePtr)
        {
            IntPtr DllUnRegisterServerPtr = GetProcAddress(modulePtr, "DllUnregisterServer");

            if (DllUnRegisterServerPtr == IntPtr.Zero)
                return ExitCode.DllUnRegisterServerNotFound;

            DllUnRegisterServer DllUnRegisterServerFunc = (DllUnRegisterServer)Marshal.GetDelegateForFunctionPointer(DllUnRegisterServerPtr, typeof(DllUnRegisterServer));

            return (DllUnRegisterServerFunc() == IntPtr.Zero) ? ExitCode.Success : ExitCode.DllUnRegisterServerNotFound;
        }

        private static ExitCode CallDllRegisterServer(IntPtr modulePtr)
        {
            IntPtr DllRegisterServerPtr = GetProcAddress(modulePtr, "DllRegisterServer");

            if (DllRegisterServerPtr == IntPtr.Zero)
                return ExitCode.DllRegisterServerNotFound;

            DllRegisterServer DllRegisterServerFunc = (DllRegisterServer)Marshal.GetDelegateForFunctionPointer(DllRegisterServerPtr, typeof(DllRegisterServer));

            return (DllRegisterServerFunc() == IntPtr.Zero) ? ExitCode.Success : ExitCode.DllRegisterServerNotFound;
        }

        private static ExitCode ProcessLibrary(string path, Func<IntPtr, ExitCode> callingAction)
        {

            // Load our module
            IntPtr modulePtr = LoadLibraryEx(path, IntPtr.Zero, LoadLibraryFlags.LOAD_WITH_ALTERED_SEARCH_PATH);

            if (modulePtr == IntPtr.Zero)
                return CheckLastError(ExitCode.Load_Failed);

            // Call our desired method
            ExitCode callResult = callingAction(modulePtr);

            // Cleanup code and return our result
            FreeLibrary(modulePtr);
            return callResult;
        }

        /// <summary>
        /// Entry point to call DllRegister of the specifed module
        /// </summary>
        /// <param name="path">Module path</param>
        /// <returns>Exit code of the current process</returns>
        public static ExitCode Register(string path)
        {
            return ProcessLibrary(path, CallDllRegisterServer);
        }


        /// <summary>
        /// Entry point to call DllUnregister of the specifed module
        /// </summary>
        /// <param name="path">Module Path</param>
        /// <returns>Exit code of the current process</returns>
        public static ExitCode UnRegister(string path)
        {
            return ProcessLibrary(path, CallDllUnRegisterServer);
        }

        /// <summary>
        /// Return last error as a exit code
        /// </summary>
        /// <param name="currentError">Current exit code</param>
        /// <returns>Last error as exit code</returns>
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

        /// <summary>
        /// Prints usage details to StdOut
        /// </summary>
        private static void ShowHelp()
        {
            string banner = typeof(Program).Assembly.GetName().Name + " v" + typeof(Program).Assembly.GetName().Version;
            string msg = "\nUsage = [-?] [-s] [-u] ModulePath";
            Console.WriteLine(banner);
            Console.WriteLine(msg);
        }
    }
}
