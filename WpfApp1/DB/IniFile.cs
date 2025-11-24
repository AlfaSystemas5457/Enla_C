using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Enla_C.DB
{
    public class IniFile
    {
        private readonly string path;

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(
            string section,
            string key,
            string defaultValue,
            StringBuilder retVal,
            int size,
            string filePath);

        public IniFile(string iniPath)
        {
            path = iniPath;
        }

        public void Write(string section, string key, string value)
        {
            WritePrivateProfileString(section, key, value, path);
        }

        public string Read(string section, string key, string defaultValue = "")
        {
            var sb = new StringBuilder(255);
            GetPrivateProfileString(section, key, defaultValue, sb, sb.Capacity, path);
            return sb.ToString();
        }

    }
}
