using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestItemStatisticsAcync.Ini
{
    internal interface IIniReaderWriter
    {
        Dictionary<string, string> ReadIniFile(string filePath, Encoding encoding = null);//// 读取 INI 文件并返回一个包含键值对的字典，使用指定编码（如 UTF-8）
        public T GetValue<T>(Dictionary<string, string> data, string key);//// 读取指定键并将值转换为指定类型
        public void WriteIniFile(string filePath, Dictionary<string, object> data, Encoding encoding = null);// 将数据写入 ini 文件，保留原有数据并更新相应节的键值对, 覆盖式写入，不能加注释
        public void WriteIniFile(string filePath, Dictionary<string, object> data, Encoding encoding = null, bool ReserveOriginalValue = true);// 将字典数据写入 INI 文件
    }
}
