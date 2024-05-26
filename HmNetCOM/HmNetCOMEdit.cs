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
        public static partial class Edit
        {
            /// <summary>
            /// キー入力があるなどの理由で処理を中断するべきかを返す。
            /// </summary>
            /// <returns>中断するべきならtrue、そうでなければfalse</returns>
            public static bool QueueStatus
            {
                get { return pCheckQueueStatus() != 0; }
            }

            /// <summary>
            /// 現在アクティブな編集領域のテキスト全体を返す。
            /// </summary>
            /// <returns>編集領域のテキスト全体</returns>
            public static string TotalText
            {
                get
                {
                    string totalText = "";
                    try
                    {
                        IntPtr hGlobal = pGetTotalTextUnicode();
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetTotalTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            totalText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return totalText;
                }
                set
                {
                    // 935.β6以降は、settotaltext() が実装された。
                    if (Version >= 935.06)
                    {
                        SetTotalText2(value);
                    }
                    else
                    {
                        SetTotalText(value);
                    }
                }
            }
            static partial void SetTotalText(string text);
            static partial void SetTotalText2(string text);


            /// <summary>
            /// 現在、単純選択している場合、その選択中のテキスト内容を返す。
            /// </summary>
            /// <returns>選択中のテキスト内容</returns>
            public static string SelectedText
            {
                get
                {
                    string selectedText = "";
                    try
                    {
                        IntPtr hGlobal = pGetSelectedTextUnicode();
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetSelectedTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            selectedText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return selectedText;
                }
                set
                {
                    SetSelectedText(value);
                }
            }
            static partial void SetSelectedText(string text);

            /// <summary>
            /// 現在、カーソルがある行(エディタ的)のテキスト内容を返す。
            /// </summary>
            /// <returns>カーソルがある行のテキスト内容</returns>
            public static string LineText
            {
                get
                {
                    string lineText = "";

                    ICursorPos pos = CursorPos;
                    if (pos.LineNo < 0 || pos.Column < 0)
                    {
                        return lineText;
                    }

                    try
                    {
                        IntPtr hGlobal = pGetLineTextUnicode(pos.LineNo);
                        if (hGlobal == IntPtr.Zero)
                        {
                            new InvalidOperationException("Hidemaru_GetLineTextUnicode_Exception");
                        }

                        var pwsz = GlobalLock(hGlobal);
                        if (pwsz != IntPtr.Zero)
                        {
                            lineText = Marshal.PtrToStringUni(pwsz);
                            GlobalUnlock(hGlobal);
                        }
                        GlobalFree(hGlobal);
                    }
                    catch (Exception)
                    {
                        throw;
                    }

                    return lineText;
                }
                set
                {
                    SetLineText(value);
                }
            }
            static partial void SetLineText(string text);

            /// <summary>
            /// CursorPos の返り値のインターフェイス
            /// </summary>
            /// <returns>(LineNo, Column)</returns>
            public interface ICursorPos
            {
                int LineNo { get; }
                int Column { get; }
            }

            private struct TCursurPos : ICursorPos
            {
                public int Column { get; set; }
                public int LineNo { get; set; }
            }

            /// <summary>
            /// MousePos の返り値のインターフェイス
            /// </summary>
            /// <returns>(LineNo, Column, X, Y)</returns>
            public interface IMousePos
            {
                int LineNo { get; }
                int Column { get; }
                int X { get; }
                int Y { get; }
            }

            private struct TMousePos : IMousePos
            {
                public int LineNo { get; set; }
                public int Column { get; set; }
                public int X { get; set; }
                public int Y { get; set; }
            }

            /// <summary>
            /// ユニコードのエディタ的な換算でのカーソルの位置を返す
            /// </summary>
            /// <returns>(LineNo, Column)</returns>
            public static ICursorPos CursorPos
            {
                get
                {
                    int lineno = -1;
                    int column = -1;
                    int success = pGetCursorPosUnicode(out lineno, out column);
                    if (success != 0)
                    {
                        TCursurPos pos = new TCursurPos();
                        pos.LineNo = lineno;
                        pos.Column = column;
                        return pos;
                    }
                    else
                    {
                        TCursurPos pos = new TCursurPos();
                        pos.LineNo = -1;
                        pos.Column = -1;
                        return pos;
                    }

                }
            }

            /// <summary>
            /// ユニコードのエディタ的な換算でのマウスの位置に対応するカーソルの位置を返す
            /// </summary>
            /// <returns>(LineNo, Column, X, Y)</returns>
            public static IMousePos MousePos
            {
                get
                {
                    POINT lpPoint;
                    bool success_1 = GetCursorPos(out lpPoint);

                    TMousePos pos = new TMousePos
                    {
                        LineNo = -1,
                        Column = -1,
                        X = -1,
                        Y = -1,
                    };

                    if (!success_1)
                    {
                        return pos;
                    }

                    int column = -1;
                    int lineno = -1;
                    int success_2 = pGetCursorPosUnicodeFromMousePos(IntPtr.Zero, out lineno, out column);
                    if (success_2 == 0)
                    {
                        return pos;
                    }

                    pos.LineNo = lineno;
                    pos.Column = column;
                    pos.X = lpPoint.X;
                    pos.Y = lpPoint.Y;
                    return pos;
                }
            }

            /// <summary>
            /// 現在開いているファイル名のフルパスを返す、無題テキストであれば、nullを返す。
            /// </summary>
            /// <returns>ファイル名のフルパス、もしくは null</returns>

            public static string FilePath
            {
                get
                {
                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        const int WM_USER = 0x400;
                        const int WM_HIDEMARUINFO = WM_USER + 181;
                        const int HIDEMARUINFO_GETFILEFULLPATH = 4;

                        StringBuilder sb = new StringBuilder(filePathMaxLength); // まぁこんくらいでさすがに十分なんじゃないの...
                        IntPtr cwch = SendMessage(hWndHidemaru, WM_HIDEMARUINFO, new IntPtr(HIDEMARUINFO_GETFILEFULLPATH), sb);
                        String filename = sb.ToString();
                        if (String.IsNullOrEmpty(filename))
                        {
                            return null;
                        }
                        else
                        {
                            return filename;
                        }
                    }
                    return null;
                }
            }

            /// <summary>
            /// 現在開いている編集エリアで、文字列の編集や何らかの具体的操作を行ったかチェックする。マクロ変数のupdatecount相当
            /// </summary>
            /// <returns>一回の操作でも数カウント上がる。32bitの値を超えると一周する。初期値は1以上。</returns>

            public static int UpdateCount
            {
                get
                {
                    if (Version < 912.98)
                    {
                        throw new MissingMethodException("Hidemaru_Edit_UpdateCount_Exception");
                    }
                    IntPtr hWndHidemaru = WindowHandle;
                    if (hWndHidemaru != IntPtr.Zero)
                    {
                        const int WM_USER = 0x400;
                        const int WM_HIDEMARUINFO = WM_USER + 181;
                        const int HIDEMARUINFO_GETUPDATECOUNT = 7;

                        IntPtr updatecount = SendMessage(hWndHidemaru, WM_HIDEMARUINFO, (IntPtr)HIDEMARUINFO_GETUPDATECOUNT, IntPtr.Zero);
                        return (int)updatecount;
                    }
                    return -1;
                }
            }

            /// <summary>
            /// <para>各種の入力ができるかどうかを判断するための状態を表します。（V9.19以降）</para>
            /// <para>以下の値の論理和です。</para>
            /// <para>0x00000002 ウィンドウ移動/サイズ変更中</para>
            /// <para>0x00000004 メニュー操作中</para>
            /// <para>0x00000008 システムメニュー操作中</para>
            /// <para>0x00000010 ポップアップメニュー操作中</para>
            /// <para>0x00000100 IME入力中</para>
            /// <para>0x00000200 何らかのダイアログ表示中</para>
            /// <para>0x00000400 ウィンドウがDisable状態</para>
            /// <para>0x00000800 非アクティブなタブまたは非表示のウィンドウ</para>
            /// <para>0x00001000 検索ダイアログの疑似モードレス状態</para>
            /// <para>0x00002000 なめらかスクロール中</para>
            /// <para>0x00004000 中ボタンによるオートスクロール中</para>
            /// <para>0x00008000 キーやマウスの操作直後(100ms 以内)</para>
            /// <para>0x00010000 何かマウスのボタンを押している</para>
            /// <para>0x00020000 マウスキャプチャ状態(ドラッグ状態)</para>
            /// <para>0x00040000 Hidemaru_CheckQueueStatus相当</para>
            /// </summary>
            /// <returns>一回の操作でも数カウント上がる。32bitの値を超えると一周する。初期値は1以上。</returns>
            public static int InputStates
            {
                get
                {
                    if (Version < 919.11)
                    {
                        throw new MissingMethodException("Hidemaru_Edit_InputStates");
                    }
                    if (pGetInputStates == null)
                    {
                        throw new MissingMethodException("Hidemaru_Edit_InputStates");
                    }

                    return pGetInputStates();
                }
            }

        }
    }
}


namespace HmNetCOM
{
    internal partial class Hm
    {
        public static partial class Edit
        {
            static partial void SetTotalText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                string cmd = $@"
                begingroupundo;
                selectall;
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                endgroupundo;
                ";
                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                } else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

            static partial void SetTotalText2(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                string cmd = $@"
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                settotaltext member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                ";
                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                } else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

            static partial void SetSelectedText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                string cmd = $@"
                if (selecting) {{
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                }}
                ";

                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                }
                else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

            static partial void SetLineText(string text)
            {
                string myDllFullPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string myTargetDllFullPath = HmMacroCOMVar.GetMyTargetDllFullPath(myDllFullPath);
                string myTargetClass = HmMacroCOMVar.GetMyTargetClass(myDllFullPath);
                HmMacroCOMVar.SetMacroVar(text);
                var pos = Edit.CursorPos;
                string cmd = $@"
                begingroupundo;
                selectline;
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                insert member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                moveto2 {pos.Column}, {pos.LineNo};
                endgroupundo;
                ";

                Macro.IResult result = null;
                if (Macro.IsExecuting)
                {
                    result = Hm.Macro.Eval(cmd);
                }
                else
                {
                    result = Hm.Macro.Exec.Eval(cmd);
                }

                HmMacroCOMVar.ClearVar();
                if (result.Error != null)
                {
                    throw result.Error;
                }
            }

        }
    }
}


namespace HmNetCOM
{

    internal static class HmEditExtentensions
    {
        public static void Deconstruct(this Hm.Edit.ICursorPos pos, out int LineNo, out int Column)
        {
            LineNo = pos.LineNo;
            Column = pos.Column;
        }

        public static void Deconstruct(this Hm.Edit.IMousePos pos, out int LineNo, out int Column, out int X, out int Y)
        {
            LineNo = pos.LineNo;
            Column = pos.Column;
            X = pos.X;
            Y = pos.Y;
        }
    }
}
