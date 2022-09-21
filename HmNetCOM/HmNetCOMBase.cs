/*
 * Copyright (C) 2021-2022 Akitsugu Komiyama
 * under the MIT License
 **/

using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace HmNetCOM
{
    internal partial class Hm
    {
        public interface IComDetachMethod
        {
            void OnReleaseObject(int reason=0);
        }

        public interface IComSupportX64
        {
#if (NET || NETCOREAPP3_1)

            bool X64MACRO() { return true; }
#else
            bool X64MACRO();
#endif
        }

        static Hm()
        {
            SetVersion();
            BindHidemaruExternFunctions();
        }

        private static void SetVersion()
        {
            if (Version == 0)
            {
                string hidemaru_fullpath = GetHidemaruExeFullPath();
                System.Diagnostics.FileVersionInfo vi = System.Diagnostics.FileVersionInfo.GetVersionInfo(hidemaru_fullpath);
                Version = 100 * vi.FileMajorPart + 10 * vi.FileMinorPart + 1 * vi.FileBuildPart + 0.01 * vi.FilePrivatePart;
            }
        }
        /// <summary>
        /// 秀丸バージョンの取得
        /// </summary>
        /// <returns>秀丸バージョン</returns>
        public static double Version { get; private set; } = 0;

        private const int filePathMaxLength = 512;

        private static string GetHidemaruExeFullPath()
        {
            var sb = new StringBuilder(filePathMaxLength);
            GetModuleFileName(IntPtr.Zero, sb, filePathMaxLength);
            string hidemaru_fullpath = sb.ToString();
            return hidemaru_fullpath;
        }

        /// <summary>
        /// 呼ばれたプロセスの現在の秀丸エディタのウィンドウハンドルを返します。
        /// </summary>
        /// <returns>現在の秀丸エディタのウィンドウハンドル</returns>
        public static IntPtr WindowHandle
        {
            get
            {
                return pGetCurrentWindowHandle();
            }
        }

        private static T HmClamp<T>(T val, T min, T max) where T : IComparable<T>
        {
            if (val.CompareTo(min) < 0) return min;
            else if (val.CompareTo(max) > 0) return max;
            else return val;
        }

        private static bool LongToInt(long number, out int intvar)
        {
            int ret_number = 0;
            while (true)
            {
                if (number > Int32.MaxValue)
                {
                    number = number - 4294967296;
                    number = number - Int32.MinValue;
                    number = number % 4294967296;
                    number = number + Int32.MinValue;
                }
                else
                {
                    break;
                }
            }
            while (true)
            {
                if (number < Int32.MinValue)
                {
                    number = number + 4294967296;
                    number = number + Int32.MinValue;
                    number = number % 4294967296;
                    number = number - Int32.MinValue;
                }
                else
                {
                    break;
                }
            }

            bool success = false;
            if (Int32.MinValue <= number && number <= Int32.MaxValue)
            {
                ret_number = (int)number;
                success = true;
            }

            intvar = ret_number;
            return success;
        }

        private static bool IsDoubleNumeric(object value)
        {
            return value is double || value is float;
        }
    }
}

namespace HmNetCOM
{

    internal partial class Hm
    {
        // 秀丸本体から出ている関数群
        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetCurrentWindowHandle();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetTotalTextUnicode();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetLineTextUnicode(int nLineNo);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetSelectedTextUnicode();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TGetCursorPosUnicode(out int pnLineNo, out int pnColumn);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TGetCursorPosUnicodeFromMousePos(IntPtr lpPoint, out int pnLineNo, out int pnColumn);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TEvalMacro([MarshalAs(UnmanagedType.LPWStr)] String pwsz);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TCheckQueueStatus();

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TAnalyzeEncoding([MarshalAs(UnmanagedType.LPWStr)] String pwszFileName, IntPtr lParam1, IntPtr lParam2);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TLoadFileUnicode([MarshalAs(UnmanagedType.LPWStr)] String pwszFileName, int nEncode, ref int pcwchOut, IntPtr lParam1, IntPtr lParam2);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate IntPtr TGetStaticVariable([MarshalAs(UnmanagedType.LPWStr)] String pwszSymbolName, int sharedMemoryFlag);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TSetStaticVariable([MarshalAs(UnmanagedType.LPWStr)] String pwszSymbolName, [MarshalAs(UnmanagedType.LPWStr)] String pwszValue, int sharedMemoryFlag);

        [UnmanagedFunctionPointer(CallingConvention.Winapi)]
        private delegate int TGetInputStates();

        // 秀丸本体から出ている関数群
        private static TGetCurrentWindowHandle pGetCurrentWindowHandle;
        private static TGetTotalTextUnicode pGetTotalTextUnicode;
        private static TGetLineTextUnicode pGetLineTextUnicode;
        private static TGetSelectedTextUnicode pGetSelectedTextUnicode;
        private static TGetCursorPosUnicode pGetCursorPosUnicode;
        private static TGetCursorPosUnicodeFromMousePos pGetCursorPosUnicodeFromMousePos;
        private static TEvalMacro pEvalMacro;
        private static TCheckQueueStatus pCheckQueueStatus;
        private static TAnalyzeEncoding pAnalyzeEncoding;
        private static TLoadFileUnicode pLoadFileUnicode;
        private static TGetStaticVariable pGetStaticVariable;
        private static TSetStaticVariable pSetStaticVariable;
        private static TGetInputStates pGetInputStates;

        // 秀丸本体のexeを指すモジュールハンドル
        private static UnManagedDll hmExeHandle;

        private static void BindHidemaruExternFunctions()
        {
            // 初めての代入のみ
            if (hmExeHandle == null)
            {
                try
                {
                    hmExeHandle = new UnManagedDll(GetHidemaruExeFullPath());

                    pGetTotalTextUnicode = hmExeHandle.GetProcDelegate<TGetTotalTextUnicode>("Hidemaru_GetTotalTextUnicode");
                    pGetLineTextUnicode = hmExeHandle.GetProcDelegate<TGetLineTextUnicode>("Hidemaru_GetLineTextUnicode");
                    pGetSelectedTextUnicode = hmExeHandle.GetProcDelegate<TGetSelectedTextUnicode>("Hidemaru_GetSelectedTextUnicode");
                    pGetCursorPosUnicode = hmExeHandle.GetProcDelegate<TGetCursorPosUnicode>("Hidemaru_GetCursorPosUnicode");
                    pEvalMacro = hmExeHandle.GetProcDelegate<TEvalMacro>("Hidemaru_EvalMacro");
                    pCheckQueueStatus = hmExeHandle.GetProcDelegate<TCheckQueueStatus>("Hidemaru_CheckQueueStatus");

                    pGetCursorPosUnicodeFromMousePos = hmExeHandle.GetProcDelegate<TGetCursorPosUnicodeFromMousePos>("Hidemaru_GetCursorPosUnicodeFromMousePos");
                    pGetCurrentWindowHandle = hmExeHandle.GetProcDelegate<TGetCurrentWindowHandle>("Hidemaru_GetCurrentWindowHandle");

                    if (Version >= 890)
                    {
                        pAnalyzeEncoding = hmExeHandle.GetProcDelegate<TAnalyzeEncoding>("Hidemaru_AnalyzeEncoding");
                        pLoadFileUnicode = hmExeHandle.GetProcDelegate<TLoadFileUnicode>("Hidemaru_LoadFileUnicode");
                    }
                    if (Version >= 915)
                    {
                        pGetStaticVariable = hmExeHandle.GetProcDelegate<TGetStaticVariable>("Hidemaru_GetStaticVariable");
                        pSetStaticVariable = hmExeHandle.GetProcDelegate<TSetStaticVariable>("Hidemaru_SetStaticVariable");
                    }
                    if (Version >= 919)
                    {
                        pGetInputStates = hmExeHandle.GetProcDelegate<TGetInputStates>("Hidemaru_GetInputStates");
                    }

                }
                catch (Exception e)
                {
                    System.Diagnostics.Trace.WriteLine(e.Message);
                }

            }
        }
    }
}

namespace HmNetCOM
{
    namespace HmNativeMethods {
        internal partial class NativeMethods
        {
            [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
            protected extern static uint GetModuleFileName(IntPtr hModule, StringBuilder lpFilename, int nSize);

            [DllImport("kernel32.dll", SetLastError = true)]
            protected extern static IntPtr GlobalLock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]

            [return: MarshalAs(UnmanagedType.Bool)]
            protected extern static bool GlobalUnlock(IntPtr hMem);

            [DllImport("kernel32.dll", SetLastError = true)]
            protected extern static IntPtr GlobalFree(IntPtr hMem);

            [StructLayout(LayoutKind.Sequential)]
            protected struct POINT
            {
                public int X;
                public int Y;
            }
            [DllImport("user32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            protected extern static bool GetCursorPos(out POINT lpPoint);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true, CharSet = CharSet.Unicode)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, StringBuilder lParam);

            [DllImport("user32.dll", EntryPoint = "SendMessage", SetLastError = true, CharSet = CharSet.Unicode)]
            protected static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, StringBuilder wParam, StringBuilder lParam);
        }
    }

    internal partial class Hm : HmNativeMethods.NativeMethods
    {
    }
}

namespace HmNetCOM
{
    namespace HmNativeMethods {
        internal partial class NativeMethods
        {
            [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
            protected static extern IntPtr LoadLibrary(string lpFileName);

            [DllImport("kernel32", CharSet = CharSet.Ansi, BestFitMapping=false, ExactSpelling=true, SetLastError=true)]
            protected static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

            [DllImport("kernel32", SetLastError = true)]
            protected static extern bool FreeLibrary(IntPtr hModule);
        }
    }

    internal partial class Hm
    {
        // アンマネージドライブラリの遅延での読み込み。C++のLoadLibraryと同じことをするため
        // これをする理由は、この実行dllとHideamru.exeが異なるディレクトリに存在する可能性があるため、
        // C#風のDllImportは成立しないからだ。
        internal sealed class UnManagedDll : HmNativeMethods.NativeMethods, IDisposable
        {

            IntPtr moduleHandle;
            public UnManagedDll(string lpFileName)
            {
                moduleHandle = LoadLibrary(lpFileName);
            }

            // コード分析などの際の警告抑制のためにデストラクタをつけておく
            ~UnManagedDll()
            {
                // C#はメインドメインのdllは(このコードが存在するdll)はプロセスが終わらない限り解放されないので、
                // ここではネイティブのdllも事前には解放しない方がよい。(プロセス終了による解放に委ねる）
                // デストラクタでは何もしない。
                // コード分析でも警告がでないように、コード分析では実行されないことがわからない形で
                // 決して実行されないコードにしておく
                if (moduleHandle == (IntPtr)(-1)) { this.Dispose(); };
            }

            public IntPtr ModuleHandle
            {
                get
                {
                    return moduleHandle;
                }
            }

            public T GetProcDelegate<T>(string method) where T : class
            {
                IntPtr methodHandle = GetProcAddress(moduleHandle, method);
                T r = Marshal.GetDelegateForFunctionPointer(methodHandle, typeof(T)) as T;
                return r;
            }

            public void Dispose()
            {
                FreeLibrary(moduleHandle);
            }
        }

    }
}

