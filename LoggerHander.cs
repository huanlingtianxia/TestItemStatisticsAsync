using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace TestItemStatisticsAcync
{
    internal class LoggerHander
    {
        public string LogFilePath { get; private set; }

        public LoggerHander(string filePath)
        {
            LogFilePath = filePath;

            // 检查文件是否存在  
            if (!File.Exists(LogFilePath))
            {
                // 如果文件不存在，可以选择创建文件或者记录一条日志  
                using (File.Create(LogFilePath))
                {
                    // 创建文件时不会写入内容  
                }
                Console.WriteLine($"日志文件 {LogFilePath} 已创建。");
            }
        }

        public void Log(string message)
        {
            using (StreamWriter writer = new StreamWriter(LogFilePath, true))
            {
                writer.WriteLine($"{DateTime.Now}: {message}");
            }
        }
    }
}
