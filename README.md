# Hm.NetCOM

[![Latest release](https://img.shields.io/github/v/release/komiyamma/hidemaru_dotnet_com?label=Hm.NetCOM)](https://github.com/komiyamma/hidemaru_dotnet_com/releases)
[![MIT License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)
![Hidemaru 8.98](https://img.shields.io/badge/Hidemaru-v8.98-6479ff.svg)
![.NET Framework 4.8](https://img.shields.io/badge/.NET_Framework-4.7_－_4.8-6479ff.svg)
![.NET 5.0 － 9.0](https://img.shields.io/badge/.NET-5.0_－_9.0-6479ff.svg)
![.NET Core 3.1](https://img.shields.io/badge/.NET_Core-3.1-6479ff.svg)
![C＃](https://img.shields.io/badge/C＃-v7.3_over-6479ff.svg)


Hm.NetCOMとは、秀丸マクロの「COM (createobject)」関連機能を用い、
- .NET 5.0 - .NET 9.0
- .NET Core 3.1
- .NET Framework 4.x

を「C#経由で秀丸マクロ」で利用するためのライブラリです。  


C#上から、「秀丸マクロ変数」の読み書きや、「秀丸マクロ関数」をC#のメソッド形態で実行することが可能です。  

秀丸マクロの「COMで呼び出すdll」の方法にて利用することを前提としています。  

https://秀丸マクロ.net/?page=nobu_tool_hm_dotnet_pinvoke

## 主な機能 (Features)

`Hm.NetCOM`ライブラリを使用することで、C#から秀丸エディタの様々な機能を直感的に操作できます。

- **テキスト操作**: 編集中のテキスト全体の取得・設定、選択範囲のテキストの取得・設定、特定の行のテキストの取得・設定が可能です。
- **カーソル・マウス位置**: カーソルの位置（行・桁）や、マウスカーソル下の文字位置を取得できます。
- **マクロの実行**: C#から秀丸マクロのコマンドを文字列として実行したり、マクロファイルを直接実行したりできます。
- **マクロ変数へのアクセス**: C#から秀丸マクロの数値変数・文字列変数への値の読み書きが可能です。
- **ファイル操作**: 秀丸エディタの文字コード自動判別機能を利用して、ファイルを読み込むことができます。
- **アウトプット枠**: アウトプット枠への文字列の出力、内容のクリアなどが可能です。
- **ファイルマネージャ枠**: ファイルマネージャ枠のモード設定やプロジェクトの操作が可能です。

## 使い方 (Usage)

### 1. C#でクラスライブラリを作成

まず、.NETのクラスライブラリプロジェクトを作成し、`HmNetCOM.dll`を参照に追加します。
そして、COMとして呼び出したい公開メソッドを実装します。

```csharp
// MyHidemaruLibrary.cs
using HmNetCOM;

public class MyHidemaruAddin
{
    // 秀丸マクロから呼び出すメソッド
    public static void ShowTotalText()
    {
        try
        {
            // 現在アクティブな秀丸エディタのテキスト全体を取得
            string text = Hm.Edit.TotalText;

            // 取得したテキストをアウトプット枠に表示
            Hm.OutputPane.Output("秀丸のテキスト内容は以下の通りです:\\n" + text);
        }
        catch (System.Exception e)
        {
            // エラーが発生した場合はアウトプット枠に表示
            Hm.OutputPane.Output("エラー: " + e.Message);
        }
    }
}
```

### 2. 秀丸マクロから呼び出す

作成したC#のライブラリ（dll）を、秀丸マクロの `createobject` を使って呼び出します。

```csharp
// Hidemaru Macro
// 作成したDLLのパスと、完全修飾クラス名を指定
#obj = createobject("C:\\path\\to\\your\\MyHidemaruLibrary.dll", "MyHidemaruAddin");

if (isobject(#obj)) {
    // C#で定義したメソッドを呼び出す
    #obj.ShowTotalText();

    // オブジェクトを解放
    releaseobject(#obj);
} else {
    message "オブジェクトの作成に失敗しました。";
}
```

## 主なクラス (Main Classes)

`HmNetCOM`の機能は、主に静的クラス `Hm` の中のネストされた静的クラスを通じて提供されます。

- `Hm.Edit`: テキストの取得/設定、カーソル位置など、現在編集中のエディタに関連する機能を提供します。
- `Hm.Macro`: マクロの実行、マクロ変数の操作など、秀丸マクロの機能と連携するための機能を提供します。
- `Hm.File`: 秀丸エディタのエンコーディング解決機能を使ってファイルを読み込む機能などを提供します。
- `Hm.OutputPane`: アウトプット枠への出力や操作に関連する機能を提供します。
- `Hm.ExplorerPane`: ファイルマネージャ枠に関連する機能を提供します。
- `Hm.Macro.Flags`: `searchoption`など、マクロのコマンドで使用する定数を提供します。
