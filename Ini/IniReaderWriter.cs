using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestItemStatisticsAcync.Ini
{
    internal class IniReaderWriter: IIniReaderWriter
    {
        #region reader ini
        // 读取 INI 文件并返回一个包含键值对的字典，使用指定编码（如 UTF-8）
        public Dictionary<string, string> ReadIniFile(string filePath, Encoding encoding = null)
        {
            encoding ??= Encoding.UTF8; // 默认为 UTF-8 编码

            var data = new Dictionary<string, string>();
            string[] lines = File.ReadAllLines(filePath, encoding);  // 使用指定编码读取文件

            string currentSection = null;

            if (!File.Exists(filePath)) return data;// 文件不存在，则返回空data

            foreach (var line in lines)
            {
                // 跳过空行和注释
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith(";") || line.StartsWith("#")) continue;

                // 如果是节（[Section]），则设置当前节
                if (line.StartsWith("["))
                {
                    currentSection = line.TrimStart('[').TrimEnd(']');
                }
                else
                {
                    // 否则是键值对
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2)
                    {
                        string key = parts[0].Trim();
                        string value = parts[1].Trim();
                        data[$"{currentSection}.{key}"] = value;
                    }
                }
            }

            return data;
        }

        // 读取指定键并将值转换为指定类型
        public T GetValue<T>(Dictionary<string, string> data, string key)
        {
            if (data.ContainsKey(key))
            {
                string value = data[key];

                // 使用 Convert.ChangeType 进行类型转换
                return (T)Convert.ChangeType(value, typeof(T));
            }

            throw new KeyNotFoundException($"Key '{key}' not found in INI file.");
        }
        #endregion

        // 将数据写入 ini 文件，保留原有数据并更新相应节的键值对, 覆盖式写入，不能加注释
        public void WriteIniFile(string filePath, Dictionary<string, object> data, Encoding encoding = null)
        {
            encoding ??= Encoding.UTF8; // 默认为 UTF-8 编码
            StringBuilder sb = new StringBuilder();
            string currentSection = null;

            // 遍历字典并生成新的内容
            foreach (var kv in data)
            {
                var parts = kv.Key.ToString().Split(new[] { '.' }, 2);
                if (parts.Length == 2)
                {
                    var section = parts[0];
                    var key = parts[1];

                    // 如果节变化，则写新的节
                    if (section != currentSection)
                    {
                        if (currentSection != null)
                        {
                            sb.AppendLine();
                        }
                        sb.AppendLine($"[{section}]");
                        currentSection = section;
                    }

                    // 将值转换为字符串并追加
                    sb.AppendLine($"{key} = {kv.Value.ToString()}");
                }
            }

            // 将修改后的内容写回文件，保留原有数据
            File.WriteAllText(filePath, sb.ToString(), encoding);
        }
        
        // 将数据写入 ini 文件，保留原有数据并更新相应节的键值对，往后面重写，但能保留不在data中的值
        public void WriteIniFile(string filePath, Dictionary<string, object> data, Encoding encoding = null, bool ReserveOriginalValue = true)
        {
            encoding ??= Encoding.UTF8; // 默认编码为 UTF-8

            // 用于保存文件的最终内容
            StringBuilder sb = new StringBuilder();

            // 读取现有的 INI 内容并解析成字典
            Dictionary<string, Dictionary<string, string>> existingData = new Dictionary<string, Dictionary<string, string>>();
            List<string> lines = new List<string>();

            if (ReserveOriginalValue && File.Exists(filePath))
            {
                lines.AddRange(File.ReadAllLines(filePath, encoding));

                string currentSection = null;

                // 解析现有的节和键值对
                foreach (var line in lines)
                {
                    var trimmedLine = line.Trim();

                    if (string.IsNullOrWhiteSpace(trimmedLine))
                    {
                        sb.AppendLine();
                        continue;
                    }

                    // 保留注释行
                    if (trimmedLine.StartsWith("#"))
                    {
                        sb.AppendLine(trimmedLine);
                        continue;
                    }

                    // 判断是否是节
                    if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                    {
                        currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);

                        if (!existingData.ContainsKey(currentSection))
                        {
                            existingData[currentSection] = new Dictionary<string, string>();
                        }
                        sb.AppendLine(trimmedLine); // 保留节名
                    }
                    else if (currentSection != null && trimmedLine.Contains("="))
                    {
                        // 解析节中的键值对
                        var keyValue = trimmedLine.Split(new[] { '=' }, 2);
                        if (keyValue.Length == 2)
                        {
                            existingData[currentSection][keyValue[0].Trim()] = keyValue[1].Trim();
                        }
                        sb.AppendLine(trimmedLine); // 保留键值对
                    }
                }
            }

            // 使用字典按节组织数据
            var sectionDict = new Dictionary<string, Dictionary<string, object>>();

            foreach (var kv in data)
            {
                var parts = kv.Key.ToString().Split(new[] { '.' }, 2);
                if (parts.Length == 2)
                {
                    var section = parts[0];
                    var key = parts[1];

                    if (!sectionDict.ContainsKey(section))
                    {
                        sectionDict[section] = new Dictionary<string, object>();
                    }

                    sectionDict[section][key] = kv.Value;
                }
            }

            // 处理每个节并更新数据
            HashSet<string> writtenSections = new HashSet<string>();

            // 遍历现有数据，处理已有节和键值对
            foreach (var section in existingData)
            {
                sb.AppendLine($"[{section.Key}]");

                // 如果新数据中包含该节，更新该节的键值对
                if (sectionDict.ContainsKey(section.Key))
                {
                    foreach (var kv in section.Value)
                    {
                        var key = kv.Key;
                        var existingValue = kv.Value;

                        // 如果新数据中有该项，更新其值
                        if (sectionDict[section.Key].ContainsKey(key))
                        {
                            var newValue = sectionDict[section.Key][key];
                            if (newValue.ToString() != existingValue)
                            {
                                sb.AppendLine($"{key} = {newValue}"); // 更新值
                            }
                            else
                            {
                                sb.AppendLine($"{key} = {existingValue}"); // 保持原值
                            }

                            // 删除已更新的项，避免重复写入
                            sectionDict[section.Key].Remove(key);
                        }
                        else
                        {
                            sb.AppendLine($"{key} = {existingValue}"); // 保持原值
                        }
                    }
                    writtenSections.Add(section.Key);
                }
                else
                {
                    // 如果该节没有在新数据中出现，保留原有内容
                    foreach (var kv in section.Value)
                    {
                        sb.AppendLine($"{kv.Key} = {kv.Value}");
                    }
                }
            }

            // 添加新数据
            foreach (var section in sectionDict)
            {
                // 如果该节之前未写入过
                if (!writtenSections.Contains(section.Key))
                {
                    sb.AppendLine($"[{section.Key}]");
                    foreach (var kv in section.Value)
                    {
                        sb.AppendLine($"{kv.Key} = {kv.Value}");
                    }
                }
            }

            // 将最终内容写回文件
            File.WriteAllText(filePath, sb.ToString(), encoding);
        }

    }
}
