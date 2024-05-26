/*
 * Copyright (C) 2021-2024 Akitsugu Komiyama
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
        public static partial class Macro
        {
            /// <summary>
            /// マクロの debuginfo と同じ関数だが、マクロ実行中以外でも扱える。
            /// 必要かどうかは微妙なところだが、マクロの debuginfo(0)～debuginfo(2) など、表示設定の反映の恩恵を受けられる点が異なる
            /// マクロの debuginfo とは異なり、マクロの実行が終えたからといって、自動的に debuginfo(0)にリセットされたりはしない。
            /// よって非表示にするためには、明示的に debuginfo(0)を、マクロ側もしくはプログラム側から明示的に実行する必要がある。
            /// </summary>
            /// <returns>成功したときは0以外、失敗したときは0を返す。</returns>
            public static int DebugInfo(params Object[] messages)
            {
                if (Version < 898)
                {
                    throw new MissingMethodException("Hidemaru_Macro_DebugInfo_Exception");
                }
                if (pDebugInfo == null)
                {
                    throw new MissingMethodException("Hidemaru_Macro_DebugInfo_Exception");
                }
                List<String> list = new List<String>();
                foreach (var exp in messages)
                {
                    var mixedString = exp.ToString();
                    string unifiedString = mixedString.Replace("\r\n", "\n").Replace("\n", "\r\n");
                    list.Add(unifiedString);
                }

                String joind = String.Join(" ", list);

                return pDebugInfo(joind);
            }

            /// <summary>
            /// マクロを実行中か否かを判定する
            /// </summary>
            /// <returns>実行中ならtrue, そうでなければfalse</returns>
            public static bool IsExecuting
            {
                get
                {
                    const int WM_USER = 0x400;
                    const int WM_ISMACROEXECUTING = WM_USER + 167;

                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        IntPtr cwch = SendMessage(hWndHidemaru, WM_ISMACROEXECUTING, IntPtr.Zero, IntPtr.Zero);
                        return (cwch != IntPtr.Zero);
                    }

                    return false;
                }
            }

            /// <summary>
            /// マクロの静的な変数
            /// </summary>
            internal static TStaticVar StaticVar = new TStaticVar();

            /// <summary>
            /// マクロの静的な変数
            /// </summary>
            internal partial class TStaticVar { 

                /// <summary>
                /// マクロの静的な変数の値(文字列)の読み書き
                /// </summary>
                /// <param name = "name">変数名</param>
                /// <param name = "value">書き込みの場合、代入する値</param>
                /// <param name = "sharedflag">共有フラグ</param>
                /// <returns>対象の静的変数名(name)に格納されている文字列</returns>
                public string this[string name, int sharedflag] {
                    get { return GetStaticVariable(name, sharedflag); }
                    set { SetStaticVariable(name, value, sharedflag); }
                }

                /// <summary>
                /// マクロの静的な変数の値(文字列)を取得する
                /// </summary>
                /// <param name = "name">変数名</param>
                /// <param name = "sharedflag">共有フラグ</param>
                /// <returns>対象の静的変数名(name)に格納されている文字列</returns>
                public string Get(string name, int sharedflag)
                {
                    return GetStaticVariable(name, sharedflag);
                }

                /// <summary>
                /// マクロの静的な変数へと値(文字列)を設定する
                /// </summary>
                /// <param name = "name">変数名</param>
                /// <param name = "value">設定する値(文字列)</param>
                /// <param name = "sharedflag">共有フラグ</param>
                /// <returns>取得に成功すれば真、失敗すれば偽が返る</returns>
                public bool Set(string name, string value, int sharedflag)
                {
                    var ret = SetStaticVariable(name, value, sharedflag);
                    if (ret != 0)
                    {
                        return true;
                    }
                    return false;
                }

                private static int SetStaticVariable(String symbolname, String value, int sharedMemoryFlag)
                {
                    try
                    {
                        if (Version < 915)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_SetGlobalVariable_Exception");
                        }
                        if (pSetStaticVariable == null)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_SetGlobalVariable_Exception");
                        }

                        return pSetStaticVariable(symbolname, value, sharedMemoryFlag);
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Trace.WriteLine(e.Message);
                        throw;
                    }
                }

                private static string GetStaticVariable(String symbolname, int sharedMemoryFlag)
                {
                    try
                    {
                        if (Version < 915)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }
                        if (pGetStaticVariable == null)
                        {
                            throw new MissingMethodException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }

                        string staticText = "";

                        IntPtr hGlobal = pGetStaticVariable(symbolname, sharedMemoryFlag);
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_Macro_GetStaticVariable_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            staticText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);

                        return staticText;
                    }
                    catch (Exception e)
                    {
                        System.Diagnostics.Trace.WriteLine(e.Message);
                        throw;
                    }
                }
            }

            /// <summary>
            /// マクロをプログラム内から実行した際の返り値のインターフェイス
            /// </summary>
            /// <returns>(Result, Message, Error)</returns>
            public interface IResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
            }

            private class TResult : IResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }

                public TResult(int Result, String Message, Exception Error)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                }
            }

            /// <summary>
            /// 現在のマクロ実行中に、プログラム中で、マクロを文字列で実行。
            /// マクロ実行中のみ実行可能なメソッド。
            /// </summary>
            /// <returns>(Result, Message, Error)</returns>
            public static IResult Eval(String expression)
            {
                TResult result;
                if (!IsExecuting)
                {
                    Exception e = new InvalidOperationException("Hidemaru_Macro_IsNotExecuting_Exception");
                    result = new TResult(-1, "", e);
                    return result;
                }
                int success = 0;
                try
                {
                    success = pEvalMacro(expression);
                }
                catch (Exception)
                {
                    throw;
                }
                if (success == 0)
                {
                    Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                    result = new TResult(0, "", e);
                    return result;
                }
                else
                {
                    result = new TResult(success, "", null);
                    return result;
                }

            }

            public static partial class Exec
            {
                /// <summary>
                /// マクロを実行していない時に、プログラム中で、マクロファイルを与えて新たなマクロを実行。
                /// マクロを実行していない時のみ実行可能なメソッド。
                /// </summary>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult File(string filepath)
                {
                    TResult result;
                    if (IsExecuting)
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_IsExecuting_Exception");
                        result = new TResult(-1, "", e);
                        return result;
                    }
                    if (!System.IO.File.Exists(filepath))
                    {
                        Exception e = new FileNotFoundException(filepath);
                        result = new TResult(-1, "", e);
                        return result;
                    }

                    const int WM_USER = 0x400;
                    const int WM_REMOTE_EXECMACRO_FILE = WM_USER + 271;
                    IntPtr hWndHidemaru = WindowHandle;

                    StringBuilder sbFileName = new StringBuilder(filepath);
                    StringBuilder sbRet = new StringBuilder("\x0f0f", 0x0f0f + 1); // 最初の値は帰り値のバッファー
                    IntPtr cwch = SendMessage(hWndHidemaru, WM_REMOTE_EXECMACRO_FILE, sbRet, sbFileName);
                    if (cwch != IntPtr.Zero)
                    {
                        result = new TResult(1, sbRet.ToString(), null);
                    }
                    else
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                        result = new TResult(0, sbRet.ToString(), e);
                    }
                    return result;
                }

                /// <summary>
                /// マクロを実行していない時に、プログラム中で、文字列で新たなマクロを実行。
                /// マクロを実行していない時のみ実行可能なメソッド。
                /// </summary>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult Eval(string expression)
                {
                    TResult result;
                    if (IsExecuting)
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_IsExecuting_Exception");
                        result = new TResult(-1, "", e);
                        return result;
                    }

                    const int WM_USER = 0x400;
                    const int WM_REMOTE_EXECMACRO_MEMORY = WM_USER + 272;
                    IntPtr hWndHidemaru = WindowHandle;

                    StringBuilder sbExpression = new StringBuilder(expression);
                    StringBuilder sbRet = new StringBuilder("\x0f0f", 0x0f0f + 1); // 最初の値は帰り値のバッファー
                    IntPtr cwch = SendMessage(hWndHidemaru, WM_REMOTE_EXECMACRO_MEMORY, sbRet, sbExpression);
                    if (cwch != IntPtr.Zero)
                    {
                        result = new TResult(1, sbRet.ToString(), null);
                    }
                    else
                    {
                        Exception e = new InvalidOperationException("Hidemaru_Macro_Eval_Exception");
                        result = new TResult(0, sbRet.ToString(), e);
                    }
                    return result;
                }
            }
        }
    }
}

namespace HmNetCOM
{

    internal static class HmMacroExtentensions
    {
        public static void Deconstruct(this Hm.Macro.IResult result, out int Result, out Exception Error, out String Message)
        {
            Result = result.Result;
            Error = result.Error;
            Message = result.Message;
        }

        public static void Deconstruct(this Hm.Macro.IFunctionResult result, out object Result, out List<Object> Args, out Exception Error, out String Message)
        {
            Result = result.Result;
            Args = result.Args;
            Error = result.Error;
            Message = result.Message;
        }

        public static void Deconstruct(this Hm.Macro.IStatementResult result, out int Result, out List<Object> Args, out Exception Error, out String Message)
        {
            Result = result.Result;
            Args = result.Args;
            Error = result.Error;
            Message = result.Message;
        }
    }
}
