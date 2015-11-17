using System;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoTypeEX
{
    public static class INI
    {
        //#==================================================================== API
        [DllImport("kernel32.dll")] private static extern int GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, StringBuilder lpReturnedString, int nSize, string lpFileName);
        [DllImport("kernel32.dll")] private static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);

        //#==================================================================== FUNCTIONS
        public static string GetValue(string section, string key, string filePath)
        {
            int chrs = 256; // buffer length
            StringBuilder buffer = new StringBuilder(chrs); // buffer string

            // if key exists, return it
            if (GetPrivateProfileString(section, key, string.Empty, buffer, chrs, filePath) != 0)
                return buffer.ToString();
            else
                return string.Empty;
        }
        public static void SetValue(string section, string key, string value, string filePath)
        {
            WritePrivateProfileString(section, key, value, filePath);
        }
    }
}
