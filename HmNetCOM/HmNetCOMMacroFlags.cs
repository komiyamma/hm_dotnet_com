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
        public static partial class Macro
        {
            public static partial class Flags {

                public partial class Encode {
                    //OPENFILE等のENCODE相当
                    public const int Sjis = 0x01;
                    public const int Utf16 = 0x02;
                    public const int Euc = 0x03;
                    public const int Jis = 0x04;
                    public const int Utf7 = 0x05;
                    public const int Utf8 = 0x06;
                    public const int Utf16_be = 0x07;
                    public const int Euro = 0x08;
                    public const int Gb2312 = 0x09;
                    public const int Big5 = 0x0a;
                    public const int Euckr = 0x0b;
                    public const int Johab = 0x0c;
                    public const int Easteuro = 0x0d;
                    public const int Baltic = 0x0e;
                    public const int Greek = 0x0f;
                    public const int Russian = 0x10;
                    public const int Symbol = 0x11;
                    public const int Turkish = 0x12;
                    public const int Hebrew = 0x13;
                    public const int Arabic = 0x14;
                    public const int Thai = 0x15;
                    public const int Vietnamese = 0x16;
                    public const int Mac = 0x17;
                    public const int Oem = 0x18;
                    public const int Default = 0x19;
                    public const int Utf32 = 0x1b;
                    public const int Utf32_be = 0x1c;
                    public const int Binary = 0x1a;
                    public const int LF = 0x40;
                    public const int CR = 0x80;

                    //SAVEASの他のオプションの数値指定
                    public const int Bom = 0x0600;
                    public const int NoBom = 0x0400;
                    public const int Selection = 0x2000;

                    //OPENFILEの他のオプションの数値指定
                    public const int NoAddHist = 0x0100;
                    public const int WS = 0x0800;
                    public const int WB = 0x1000;
                }

                public static partial class SearchOption {
                    //searchoption(検索関係)
                    public const int Word =                 0x00000001;
                    public const int Casesense =            0x00000002;
                    public const int NoCasesense =          0x00000000;
                    public const int Regular =              0x00000010;
                    public const int NoRegular =            0x00000000;
                    public const int Fuzzy =                0x00000020;
                    public const int Hilight =              0x00003800;
                    public const int NoHilight =            0x00002000;
                    public const int LinkNext =             0x00000080;
                    public const int Loop =                 0x01000000;

                    //searchoption(マスク関係)
                    public const int MaskComment =          0x00020000;
                    public const int MaskIfdef =            0x00040000;
                    public const int MaskNormal =           0x00010000;
                    public const int MaskScript =           0x00080000;
                    public const int MaskString =           0x00100000;
                    public const int MaskTag =              0x00200000;
                    public const int MaskOnly =             0x00400000;
                    public const int FEnableMaskFlags =     0x00800000;

                    //searchoption(置換関係)
                    public const int FEnableReplace =       0x00000004;
                    public const int Ask =                  0x00000008;
                    public const int NoClose =              0x02000000;

                    //searchoption(grep関係)
                    public const int SubDir =               0x00000100;
                    public const int Icon =                 0x00000200;
                    public const int Filelist =             0x00000040;
                    public const int FullPath =             0x00000400;
                    public const int OutputSingle =         0x10000000;
                    public const int OutputSameTab =        0x20000000;

                    //searchoption(grepして置換関係)
                    public const int BackUp =               0x04000000;
                    public const int Preview =              0x08000000;
                    
                    // searchoption2を使うよ、というフラグ。なんと、int32_maxを超えているので、特殊な処理が必要。
                    static long FEnableSearchOption2
                    {
                        get
                        {
                            if (IntPtr.Size == 4) { return -0x80000000; } else { return 0x80000000; }
                        }
                    }
                }

                public static partial class SearchOption2 {
                    //searchoption2
                    public const int UnMatch =              0x00000001;
                    public const int InColorMarker =        0x00000002;
                    public const int FGrepFormColumn =      0x00000008;
                    public const int FGrepFormHitOnly =     0x00000010;
                    public const int FGrepFormSortDate =    0x00000020;
                }
            }
        }
    }
}