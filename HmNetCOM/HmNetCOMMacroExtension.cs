﻿/*
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
    // このインターフェイスは秀丸マクロのjsmode(WebView2)でCOMを呼び出す際に必要
    interface IHmMacroCOMVar
    {
        object DllToMacro();
        int MacroToDll(object variable);
        int MethodToDll(String dllfullpath, String typefullname, String methodname, String message_param);
    }

    public partial class HmMacroCOMVar {
        private const string HmMacroCOMVarInterface = "HmMacroCOMVarGUID";
    }
}

namespace HmNetCOM
{
    // 秀丸のCOMから呼び出して、マクロ⇔COMといったように、マクロとプログラムで変数値を互いに伝搬する
    [ComVisible(true)]
#if (NET || NETCOREAPP3_1)
#else
    [ClassInterface(ClassInterfaceType.None)]
#endif
    [Guid(HmMacroCOMVarInterface)]
    public partial class HmMacroCOMVar : IHmMacroCOMVar, Hm.IComSupportX64
    {
        private static object marcroVar = null;
        public object DllToMacro()
        {
            return marcroVar;
        }
        public int MacroToDll(object variable)
        {
            marcroVar = variable;
            return 1;
        }
        public int MethodToDll(String dllfullpath, String typefullname, String methodname, String message_param)
        {
            marcroVar = message_param;

            try
            {
                MethodToDllHelper(dllfullpath, typefullname, methodname, message_param);
                return 1;
            }
            catch(Exception e)
            {
                System.Diagnostics.Trace.WriteLine(e);
            }
            return 0;
        }

        private void TraceMethodInfo(String assm_path, String class_name, String method_name)
        {
            System.Diagnostics.Trace.WriteLine("アセンブリパス   :" + assm_path);
            System.Diagnostics.Trace.WriteLine("名前空間.クラス名:" + class_name);
            System.Diagnostics.Trace.WriteLine("メソッド名       :" + method_name);
        }
        private static void TraceExceptionInfo(Exception e)
        {
            System.Diagnostics.Trace.WriteLine(e.GetType());
            System.Diagnostics.Trace.WriteLine(e.Message);
            System.Diagnostics.Trace.WriteLine(e.StackTrace);
        }
        private Object MethodToDllHelper(String assm_path, String class_name, String method_name, String message_param)
        {
            Exception method_ex = null;
            try
            {
                Assembly assm = null;
                Type t = null;

                if (assm_path.Length > 0)
                {
                    assm = Assembly.LoadFile(assm_path);
                    if (assm == null)
                    {
                        System.Diagnostics.Trace.WriteLine("ロード出来ない");
                    }
                    else
                    {
                        // System::Diagnostics::Trace::WriteLine(assm->FullName);
                    }

                    foreach (Type t2 in assm.GetExportedTypes())
                    {
                        if (t2.ToString() == class_name)
                        {
                            t = assm.GetType(class_name);
                        }
                    }
                }
                else
                {
                    t = Type.GetType(class_name);
                }
                if (t == null)
                {
                    System.Diagnostics.Trace.WriteLine("MissingMethodException(クラスもしくはメソッドを見つけることが出来ない):");
                    TraceMethodInfo(assm_path, class_name, method_name);
                    return null;
                }

                // メソッドの定義タイプを探る。
                MethodInfo m;
                try
                {
                    m = t.GetMethod(method_name);
                }
                catch (Exception ex)
                {
                    // 基本コースだと一致してない系の可能性やオーバーロードなど未解決エラーを控えておく
                    // t->GetMethod(...)は論理的には不要だが、「エラー情報のときにわかりやすい情報を.NETに自動で出力してもらう」ためにダミーで呼び出しておく
                    method_ex = ex;

                    // オーバーロードなら1つに解決できるように型情報も含めてmは上書き
                    List<Type> args_types = new List<Type>();
                    args_types.Add(Type.GetType(message_param));
                    m = t.GetMethod(method_name, args_types.ToArray());
                }

                Object o = null;
                try
                {
                    // オーバーロードなら1つに解決できるように型情報も含めてmは上書き
                    List<Object> args_values = new List<Object>();
                    args_values.Add(message_param);
                    o = m.Invoke(null, args_values.ToArray());
                }
                catch (Exception)
                {
                    System.Diagnostics.Trace.WriteLine("指定のメソッドの実行時、例外が発生しました。");
                    throw;
                }
                return o;
            }
            catch (Exception e)
            {
                System.Diagnostics.Trace.WriteLine("指定のアセンブリやメソッドを特定する前に、例外が発生しました。");
                TraceMethodInfo(assm_path, class_name, method_name);
                if (method_ex != null)
                {
                    TraceExceptionInfo(method_ex);
                }
                TraceExceptionInfo(e);
            }

            return null;

        }
        public bool X64MACRO() {
            return true;
        }
    }

    public partial class HmMacroCOMVar
    {
        static HmMacroCOMVar()
        {
            var h = new HmMacroCOMVar();
            myGuidLabel = h.GetType().GUID.ToString();
            myClassFullName = h.GetType().FullName;
        }

        internal static void SetMacroVar(object obj)
        {
            marcroVar = obj;
        }
        internal static object GetMacroVar()
        {
            return marcroVar;
        }
        private static string myGuidLabel = "";
        private static string myClassFullName = "";

        internal static string GetMyTargetDllFullPath(string thisDllFullPath)
        {
            string myTargetClass = myClassFullName;
            string thisComHostFullPath = System.IO.Path.ChangeExtension(thisDllFullPath, "comhost.dll");
            if (System.IO.File.Exists(thisComHostFullPath))
            {
                return thisComHostFullPath;
            }

            return thisDllFullPath;
        }

        internal static string GetMyTargetClass(string thisDllFullPath)
        {
            string myTargetClass = myClassFullName;
            string thisComHostFullPath = System.IO.Path.ChangeExtension(thisDllFullPath, "comhost.dll");
            if (System.IO.File.Exists(thisComHostFullPath))
            {
                myTargetClass = "{" + myGuidLabel + "}";
            }

            return myTargetClass;
        }

        internal static object GetVar(string var_name)
        {
            string myDllFullPath = GetExecutingAssemblyLocation();
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            var result = Hm.Macro.Eval($@"
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                #_COM_NET_PINVOKE_MACRO_VAR_RESULT = member(#_COM_NET_PINVOKE_MACRO_VAR, ""MacroToDll"", {var_name});
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
                #_COM_NET_PINVOKE_MACRO_VAR_RESULT = 0;
            ");
            if (result.Error != null)
            {
                throw result.Error;
            }
            return HmMacroCOMVar.marcroVar;
        }

        internal static int SetVar(string var_name, object obj)
        {
            string myDllFullPath = GetExecutingAssemblyLocation();
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            HmMacroCOMVar.marcroVar = obj;
            var result = Hm.Macro.Eval($@"
                #_COM_NET_PINVOKE_MACRO_VAR = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                {var_name} = member(#_COM_NET_PINVOKE_MACRO_VAR, ""DllToMacro"" );
                releaseobject(#_COM_NET_PINVOKE_MACRO_VAR);
            ");
            if (result.Error != null)
            {
                throw result.Error;
            }
            return 1;
        }

        internal static void ClearVar()
        {
            HmMacroCOMVar.marcroVar = null;
        }

        internal static Hm.Macro.IResult BornMacroScopeMethod(String scopename, String dllfullpath, String typefullname, String methodname)
        {

            string myDllFullPath = GetExecutingAssemblyLocation();
            string myTargetDllFullPath = GetMyTargetDllFullPath(myDllFullPath);
            string myTargetClass = GetMyTargetClass(myDllFullPath);
            ClearVar();
            var result = Hm.Macro.Exec.Eval($@"
                #_COM_NET_PINVOKE_METHOD_CALL = createobject(@""{myTargetDllFullPath}"", @""{myTargetClass}"" );
                #_COM_NET_PINVOKE_METHOD_CALL_RESULT = member(#_COM_NET_PINVOKE_METHOD_CALL, ""MethodToDll"", @""{dllfullpath}"", @""{typefullname}"", @""{methodname}"",  R""MACRO_OF_SCOPENAME({scopename})MACRO_OF_SCOPENAME"");
                releaseobject(#_COM_NET_PINVOKE_METHOD_CALL);
                #_COM_NET_PINVOKE_METHOD_CALL_RESULT = 0;
            ");
            return result;
        }
    }


    internal partial class Hm
    {
        public static partial class Macro
        {

            public static partial class Exec
            {
                /// <summary>
                /// 指定のC#のstaticメソッドを「新たなマクロ実行空間」として呼び出す
                /// </summary>
                /// <param name = "message_parameter">文字列パラメータ</param>
                /// <param name = "delegate_method">呼び出したいC#メソッド「public methodname(string message_parameter)の型」に従うメソッドであること</param>
                /// <returns>(Result, Message, Error)</returns>
                public static IResult Method(string message_parameter, Delegate delegate_method)
                {
                    string parameter = message_parameter;
                    // 渡されたメソッドが自分自身のdllと異なるのはダメ
                    if (delegate_method.Method.DeclaringType.Assembly.Location != GetExecutingAssemblyLocation();
                    {
                        string message_no_dll_myself = "The Delegate method must in " + GetExecutingAssemblyLocation();
                        var result_no_dll_myself = new TResult(0, "", new MissingMethodException(message_no_dll_myself));
                        System.Diagnostics.Trace.WriteLine(result_no_dll_myself);
                        return result_no_dll_myself;
                    }
                    else if (delegate_method.Method.IsStatic && delegate_method.Method.IsPublic)
                    {
                        var ret = HmMacroCOMVar.BornMacroScopeMethod(parameter, delegate_method.Method.DeclaringType.Assembly.Location, delegate_method.Method.DeclaringType.FullName, delegate_method.Method.Name);
                        if (ret.Result > 0) {
                            var result = new TResult(ret.Result, message_parameter, ret.Error);
                            return result;
                        } else {
                            var result = new TResult(ret.Result, ret.Message, ret.Error);
                            return result;
                        }
                    }
                    else if (!delegate_method.Method.IsStatic)
                    {

                        string message_no_static = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + " is not 'STATIC' in " + delegate_method.Method.DeclaringType.Assembly.Location;
                        var result_no_static = new TResult(0, "", new MissingMethodException(message_no_static));
                        System.Diagnostics.Trace.WriteLine(message_no_static);
                        return result_no_static;
                    }
                    else if (!delegate_method.Method.IsPublic)
                    {
                        string message_no_public = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + " is not 'PUBLIC' in " + delegate_method.Method.DeclaringType.Assembly.Location;
                        var result_no_public = new TResult(0, "", new MissingMethodException(message_no_public));
                        System.Diagnostics.Trace.WriteLine(message_no_public);
                        return result_no_public;
                    }
                    string message_missing = delegate_method.Method.DeclaringType.FullName + "." + delegate_method.Method.Name + "is 'MISSING' access in " + delegate_method.Method.DeclaringType.Assembly.Location;
                    var result_missing = new TResult(0, "", new MissingMethodException(delegate_method.Method.Name + " is missing access"));
                    System.Diagnostics.Trace.WriteLine(result_missing);
                    return result_missing;
                }
            }
        }


        public static partial class Macro
        {
            // マクロでの問い合わせ結果系
            public interface IStatementResult
            {
                int Result { get; }
                String Message { get; }
                Exception Error { get; }
                List<Object> Args { get; }
            }


            private class TStatementResult : IStatementResult
            {
                public int Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public List<Object> Args { get; set; }

                public TStatementResult(int Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<object>(Args); // コピー渡し
                }
            }

            private static int statement_base_random = 0;
            /// <summary>
            /// 秀丸マクロの関数のような「命令文」を実行
            /// </summary>
            /// <param name = "statement_name">（関数のような）命令文名</param>
            /// <param name = "args">命令文の引数</param>
            /// <returns>(Result, Args, Message, Error)</returns>
            internal static IStatementResult Statement(string statement_name, params object[] args)
            {
                string funcname = statement_name;
                if (statement_base_random == 0)
                {
                    statement_base_random = new System.Random().Next(Int16.MaxValue) + 1;

                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, statement_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「文」で実行する形
                string expression = $"{funcname} {args_string};\n";

                // 実行する
                IResult ret = Macro.Eval(expression);

                int macro_result = ret.Result;
                if (ret.Error == null)
                {
                    try
                    {
                        Object tmp_var = Macro.Var["result"]; // この中のGetMethodで例外が発生する可能性あり

                        if (IntPtr.Size == 4)
                        {
                            macro_result = (Int32)tmp_var + 0; // 確実に複製を
                        }
                        else
                        {
                            Int64 macro_result64 = (Int64)tmp_var + 0; // 確実に複製を
                            Int32 macro_result32 = (Int32)HmClamp<Int64>(macro_result64, Int32.MinValue, Int32.MaxValue);
                            macro_result = (Int32)macro_result32;
                        }
                    } catch(Exception) {
                    }
                }

                // 成否も含めて結果を入れる。
                IStatementResult result = new TStatementResult(macro_result, ret.Message, ret.Error, new List<Object>());

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = 0;
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = "";
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            Macro.Var[l.Key + "[" + iix + "]"] = "";
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }

                return result;
            }

            // マクロでの問い合わせ結果系
            public interface IFunctionResult
            {
                object Result { get; }
                String Message { get; }
                Exception Error { get; }
                List<Object> Args { get; }
            }

            private class TFunctionResult : IFunctionResult
            {
                public object Result { get; set; }
                public string Message { get; set; }
                public Exception Error { get; set; }
                public List<Object> Args { get; set; }

                public TFunctionResult(object Result, String Message, Exception Error, List<Object> Args)
                {
                    this.Result = Result;
                    this.Message = Message;
                    this.Error = Error;
                    this.Args = new List<object>(Args); // コピー渡し
                }
            }

            private static int funciton_base_random = 0;
            /// <summary>
            /// 秀丸マクロの「関数」を実行
            /// </summary>
            /// <param name = "func_name">関数名</param>
            /// <param name = "args">関数の引数</param>
            /// <returns>(Result, Args, Message, Error)</returns>
            public static IFunctionResult Function(string func_name, params object[] args)
            {
                return _AsFunction<Object>(func_name, args);
            }

            /// <summary>
            /// 秀丸マクロの「関数」を実行。関数だけだと返り値が不明な場合にこの<T>付きを使用する。
            /// </summary>
            /// <param name = "func_name">関数名</param>
            /// <param name = "args">関数の引数</param>
            /// <typeparam name="T">String | int | long | IntPtr | double。関数単体だけ確定されない返り値の型を「文字列タイプ」か「整数タイプ」かに振り分け直す。</typeparam>
            /// <returns>(Result, Args, Message, Error)</returns>
            public static IFunctionResult Function<T>(string func_name, params object[] args)
            {
                return _AsFunction<T>(func_name, args);
            }

            public static IFunctionResult _AsFunction<T>(string func_name, params object[] args)
            {
                string funcname = func_name;
                if (funciton_base_random == 0)
                {
                    funciton_base_random = new System.Random().Next(Int16.MaxValue) + 1;

                }

                List<KeyValuePair<string, object>> arg_list = SetMacroVarAndMakeMacroKeyArray(args, funciton_base_random);

                // keyをリスト化
                var arg_keys = new List<String>();
                foreach (var l in arg_list)
                {
                    arg_keys.Add(l.Key);
                }

                // それを「,」で繋げる
                string args_string = String.Join(", ", arg_keys);
                // それを指定の「関数」で実行する形
                string expression = "";

                string result_temp = "";
                Macro.IResult eval_result = new TResult(-1, "", null);
                if (typeof(T)==typeof(int) || typeof(T)==typeof(long) || typeof(T)==typeof(IntPtr) || typeof(T) == typeof(double))
                {
                    expression = $"{funcname}({args_string})";
                    result_temp = "##_tmp_dll_expression_ret";
                    string eval_expresson = result_temp + " = " + expression + ";\n";
                    eval_result = Eval(eval_expresson);
                    expression = result_temp;
                } else if (typeof(T)==typeof(String)) {
                    expression = $"{funcname}({args_string})";
                    result_temp = "$$_tmp_dll_expression_ret";
                    string eval_expresson = result_temp + " = " + expression + ";\n";
                    eval_result = Eval(eval_expresson);
                    expression = result_temp;
                } else {
                    expression = $"{funcname}({args_string})";
                }
                //----------------------------------------------------------------
                TFunctionResult result = new TFunctionResult(null, "", null, new List<Object>());
                result.Args = new List<object>();

                Object ret = null;
                try
                {
                    ret = Macro.Var[expression]; // この中のGetMethodで例外が発生する可能性あり

                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            result.Result = (Int32)ret + 0; // 確実に複製を
                            result.Message = "";
                            result.Error = null;
                        }
                        else
                        {
                            result.Result = (Int64)ret + 0; // 確実に複製を
                            result.Message = "";
                            result.Error = null;
                        }
                    }
                    else
                    {
                        result.Result = (String)ret + ""; // 確実に複製を
                        result.Message = "";
                        result.Error = null;
                    }

                }
                catch (Exception e)
                {
                    result.Result = null;
                    result.Message = "";
                    result.Error = e;
                }

                if (result_temp.StartsWith("#")) {
                    Macro.Var[result_temp] = 0;
                    if (eval_result?.Error != null) {
                        result.Result = null;
                        result.Message = "";
                        result.Error = eval_result.Error;
                    }
                } else if (result_temp.StartsWith("$")) {
                    Macro.Var[result_temp] = "";
                    if (eval_result?.Error != null) {
                        result.Result = null;
                        result.Message = "";
                        result.Error = eval_result.Error;
                    }
                }

                // 使ったので削除
                for (int ix = 0; ix < arg_list.Count; ix++)
                {
                    var l = arg_list[ix];
                    if (l.Value is Int32 || l.Value is Int64)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = 0;
                    }
                    else if (l.Value is string)
                    {
                        result.Args.Add(Macro.Var[l.Key]);
                        Macro.Var[l.Key] = "";
                    }

                    else if (l.Value.GetType() == new List<int>().GetType() || l.Value.GetType() == new List<long>().GetType() || l.Value.GetType() == new List<IntPtr>().GetType())
                    {
                        result.Args.Add(l.Value);
                        if (l.Value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)l.Value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)l.Value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                        else if (l.Value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)l.Value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[l.Key + "[" + iix + "]"] = 0;
                            }
                        }
                    }
                    else if (l.Value.GetType() == new List<String>().GetType())
                    {
                        result.Args.Add(l.Value);
                        List<String> ptr_list = (List<String>)l.Value;
                        for (int iix = 0; iix < ptr_list.Count; iix++)
                        {
                            Macro.Var[l.Key + "[" + iix + "]"] = "";
                        }
                    }
                    else
                    {
                        result.Args.Add(l.Value);
                    }
                }

                return result;
            }

            private static List<KeyValuePair<string, object>> SetMacroVarAndMakeMacroKeyArray(object[] args, int base_random)
            {
                var arg_list = new List<KeyValuePair<String, Object>>();
                int cur_random = new Random().Next(Int16.MaxValue) + 1;
                foreach (var value in args)
                {
                    bool success = false;
                    cur_random++;
                    object normalized_arg = null;
                    // Boolean型であれば、True:1 Flase:0にマッピングする
                    if (value is bool)
                    {
                        success = true;
                        if ((bool)value == true)
                        {
                            normalized_arg = 1;
                        }
                        else
                        {
                            normalized_arg = 0;
                        }
                    }

                    if (value is string || value is StringBuilder)
                    {
                        success = true;
                        normalized_arg = value.ToString();
                    }

                    // 配列の場合を追加
                    if (!success)
                    {
                        if (value.GetType() == new List<int>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<int>((List<int>)value);
                        }
                        if (value.GetType() == new List<long>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<long>((List<long>)value);
                        }
                        if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<IntPtr>((List<IntPtr>)value);
                        }
                    }

                    if (!success)
                    {
                        if (value.GetType() == new List<string>().GetType())
                        {
                            success = true;
                            normalized_arg = new List<String>((List<String>)value);
                        }
                    }
                    // 以上配列の場合を追加

                    if (!success)
                    {
                        // 32bit
                        if (IntPtr.Size == 4)
                        {
                            // まずは整数でトライ
                            Int32 itmp = 0;
                            try
                            {
                                // intでもIntPtrでもないならば...
                                if (value.GetType() != typeof(int).GetType() && value.GetType() != typeof(IntPtr).GetType())
                                {
                                    int itmp_cycle_bit = 0;
                                    long ltmp = 0;
                                    bool suc = Int64.TryParse(value.ToString(), out ltmp);
                                    if (suc)
                                    {
                                        success = LongToInt((long)ltmp, out itmp_cycle_bit);
                                        itmp = itmp_cycle_bit;
                                    }
                                }
                            }
                            catch (Exception)
                            {

                            }
                            if (!success)
                            {
                                success = Int32.TryParse(value.ToString(), out itmp);
                            }
                            if (success == true)
                            {
                                itmp = HmClamp<Int32>(itmp, Int32.MinValue, Int32.MaxValue);
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    normalized_arg = (Int32)(dtmp);
                                }

                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int64>(itmp, Int64.MinValue, Int64.MaxValue);
                                normalized_arg = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    normalized_arg = (Int64)(dtmp);
                                }
                                else
                                {
                                    normalized_arg = 0;
                                }
                            }
                        }
                    }


                    // 成功しなかった
                    if (!success)
                    {
                        normalized_arg = value.ToString();
                    }

                    if (normalized_arg is Int32 || normalized_arg is Int64)
                    {
                        string key = "#AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        Macro.Var[key] = normalized_arg;
                    }
                    else if (normalized_arg is string)
                    {
                        string key = "$AsMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        Macro.Var[key] = normalized_arg;
                    }
                    else if (value.GetType() == new List<int>().GetType() || value.GetType() == new List<long>().GetType() || value.GetType() == new List<IntPtr>().GetType())
                    {
                        string key = "$AsIntArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        if (value.GetType() == new List<int>().GetType())
                        {
                            List<int> int_list = (List<int>)value;
                            for (int iix = 0; iix < int_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = int_list[iix];
                            }
                        }
                        else if (value.GetType() == new List<long>().GetType())
                        {
                            List<long> long_list = (List<long>)value;
                            for (int iix = 0; iix < long_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = long_list[iix];
                            }
                        }
                        else if (value.GetType() == new List<IntPtr>().GetType())
                        {
                            List<IntPtr> ptr_list = (List<IntPtr>)value;
                            for (int iix = 0; iix < ptr_list.Count; iix++)
                            {
                                Macro.Var[key + "[" + iix + "]"] = ptr_list[iix];
                            }
                        }
                    }
                    else if (value.GetType() == new List<string>().GetType())
                    {
                        string key = "$AsStrArrayOfMacroArs_" + base_random.ToString() + '_' + cur_random.ToString();
                        arg_list.Add(new KeyValuePair<string, object>(key, normalized_arg));
                        List<String> str_list = (List<String>)value;
                        for (int iix = 0; iix < str_list.Count; iix++)
                        {
                            Macro.Var[key + "[" + iix + "]"] = str_list[iix];
                        }
                    }
                }
                return arg_list;
            }


            internal static TMacroVar Var = new TMacroVar();
            internal sealed class TMacroVar
            {
                /// <summary>
                /// 対象の「秀丸マクロ変数名」への読み書き
                /// </summary>
                /// <param name = "var_name">変数のシンボル名</param>
                /// <param name = "value">書き込みの場合、代入する値</param>
                /// <returns>読み取りの場合は、対象の変数の値</returns>
                public Object this[String var_name]
                {
                    get
                    {
                        return GetMethod(var_name);
                    }
                    set
                    {
                        value = SetMethod(var_name, value);
                    }
                }

                private static object SetMethod(string var_name, object value)
                {
                    if (var_name.StartsWith("#"))
                    {
                        Object result = new Object();

                        // Boolean型であれば、True:1 Flase:0にマッピングする
                        if (value is bool)
                        {
                            if ((Boolean)value == true)
                            {
                                value = 1;
                            }
                            else
                            {
                                value = 0;
                            }
                        }

                        // 32bit
                        if (IntPtr.Size == 4)
                        {

                            // まずは整数でトライ
                            Int32 itmp = 0;
                            bool success = false;
                            try
                            {
                                // intでもIntPtrでもないならば...
                                if (value.GetType() != typeof(int).GetType() && value.GetType() != typeof(IntPtr).GetType())
                                {
                                    int itmp_cycle_bit = 0;
                                    long ltmp = 0;
                                    bool suc = Int64.TryParse(value.ToString(), out ltmp);
                                    if (suc)
                                    {
                                        success = LongToInt((long)ltmp, out itmp_cycle_bit);
                                        itmp = itmp_cycle_bit;
                                    }
                                }
                            }
                            catch(Exception)
                            {

                            }
                            if (!success)
                            {
                                success = Int32.TryParse(value.ToString(), out itmp);
                            }

                            if (success == true)
                            {
                                itmp = HmClamp<Int32>(itmp, Int32.MinValue, Int32.MaxValue);
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int32.MinValue, Int32.MaxValue);
                                    result = (Int32)(dtmp);
                                }

                                else
                                {
                                    result = 0;
                                }
                            }
                        }

                        // 64bit
                        else
                        {
                            // まずは整数でトライ
                            Int64 itmp = 0;
                            bool success = Int64.TryParse(value.ToString(), out itmp);

                            if (success == true)
                            {
                                itmp = HmClamp<Int64>(itmp, Int64.MinValue, Int64.MaxValue);
                                result = itmp;
                            }

                            else
                            {
                                // 次に少数でトライ
                                Double dtmp = 0;
                                if (IsDoubleNumeric(value))
                                {
                                    dtmp = (double)value;
                                    success = true;
                                }
                                else
                                {
                                    success = double.TryParse(value.ToString(), out dtmp);
                                }
                                if (success)
                                {
                                    dtmp = HmClamp<double>(dtmp, Int64.MinValue, Int64.MaxValue);
                                    result = (Int64)(dtmp);
                                }
                                else
                                {
                                    result = 0;
                                }
                            }
                        }
                        HmMacroCOMVar.SetVar(var_name, value);
                        HmMacroCOMVar.ClearVar();
                    }

                    else // if (var_name.StartsWith("$")
                    {

                        String result = value.ToString();
                        HmMacroCOMVar.SetVar(var_name, value);
                        HmMacroCOMVar.ClearVar();
                    }

                    return value;
                }

                private static object GetMethod(string var_name)
                {
                    HmMacroCOMVar.ClearVar();
                    Object ret = HmMacroCOMVar.GetVar(var_name);
                    if (ret.GetType().Name != "String")
                    {
                        if (IntPtr.Size == 4)
                        {
                            try {
                                return (Int32)ret + 0;
                            } catch(Exception) {
                            }
                            return (Int32)(dynamic)ret + 0; // 確実に複製を
                        }
                        else
                        {
                            try {
                                return (Int64)ret + 0;
                            } catch(Exception) {
                            }
                            return (Int64)(dynamic)ret + 0; // 確実に複製を
                        }
                    }
                    else
                    {
                        return (String)ret + ""; // 確実に複製を
                    }
                }
            }
        }
    }
}

