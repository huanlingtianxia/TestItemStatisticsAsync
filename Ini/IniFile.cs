using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;


namespace TestItemStatisticsAcync.Ini
{
    internal class IniFile
    {
        private string _filePath;

        public IniFile(string filePath)
        {
            _filePath = filePath;
        }

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        private static extern bool WritePrivateProfileString(string section, string key, string value, string filePath);


        // 读取INI文件中的字符串  
        public string Read(string section, string key, string defaultValue = "")
        {
            StringBuilder retVal = new StringBuilder(255);
            GetPrivateProfileString(section, key, defaultValue, retVal, retVal.Capacity, _filePath);
            return retVal.ToString();
        }

        // 写入INI文件中的字符串，保留备注  
        public void Write(string section, string key, object value)
        {
            string currentValue = Read(section, key);
            // 检查当前值，如果存在则保留注释  
            string newValue = (currentValue.Contains("#") || currentValue.Contains(";") || string.IsNullOrWhiteSpace(currentValue)) ? currentValue : $" {value}";

            WritePrivateProfileString(section, key, newValue, _filePath);
        }

        // 获取section全部键值对  
        public Dictionary<string, string> GetAllKeys(string section)
        {
            var result = new Dictionary<string, string>();
            var buffer = new StringBuilder(65536);
            GetPrivateProfileString(section, null, null, buffer, buffer.Capacity, _filePath);
            string[] keys = buffer.ToString().Split(new char[1] { '\0' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var key in keys)
            {
                result[key] = Read(section, key);
            }

            return result;
        }
    }
}
