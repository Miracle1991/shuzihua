using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
namespace shuzihua
{
    public class FileComparer : IComparer<string>
    {
        [System.Runtime.InteropServices.DllImport("Shlwapi.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public static extern int StrCmpLogicalW(string psz1, string psz2);
        public int Compare(string psz1, string psz2)
        {
            return StrCmpLogicalW(psz1, psz2);
        }
    }
}
