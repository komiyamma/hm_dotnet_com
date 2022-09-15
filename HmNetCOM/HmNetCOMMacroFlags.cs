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

                public static partial class SearchOption {
                    //searchoption(検索関係)
                    const int Word =                 0x00000001;
                    const int Casesense =            0x00000002;
                    const int NoCasesense =          0x00000000;
                    const int Regular =              0x00000010;
                    const int NoRegular =            0x00000000;
                    const int Fuzzy =                0x00000020;
                    const int Hilight =              0x00003800;
                    const int NoHilight =            0x00002000;
                    const int LinkNext =             0x00000080;
                    const int Loop =                 0x01000000;

                    //searchoption(マスク関係)
                    const int MaskComment =          0x00020000;
                    const int MaskIfdef =            0x00040000;
                    const int MaskNormal =           0x00010000;
                    const int MaskScript =           0x00080000;
                    const int MaskString =           0x00100000;
                    const int MaskTag =              0x00200000;
                    const int MaskOnly =             0x00400000;
                    const int FEnableMaskFlags =     0x00800000;

                    //searchoption(置換関係)
                    const int FEnableReplace =       0x00000004;
                    const int Ask =                  0x00000008;
                    const int NoClose =              0x02000000;

                    //searchoption(grep関係)
                    const int SubDir =               0x00000100;
                    const int Icon =                 0x00000200;
                    const int Filelist =             0x00000040;
                    const int FullPath =             0x00000400;
                    const int OutputSingle =         0x10000000;
                    const int OutputSameTab =        0x20000000;

                    //searchoption(grepして置換関係)
                    const int BackUp =               0x04000000;
                    const int Preview =              0x08000000;
/*
                    int FEnableSearchOption2 {
                        get {
                            if (IntPtr.Size == 4) { return -0x80000000; } else { return 0x80000000; }
                        }
                    };
*/
                }


                public static partial class SearchOption2 {
                    //searchoption2
                    const int UnMatch =              0x00000001;
                    const int InColorMarker =        0x00000002;
                    const int FGrepFormColumn =      0x00000008;
                    const int FGrepFormHitOnly =     0x00000010;
                    const int FGrepFormSortDate =    0x00000020;
                }
            }
        }
    }
}