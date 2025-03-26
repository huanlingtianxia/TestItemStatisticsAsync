using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml;
using TestItemStatistics.ExcelOp;

namespace TestItemStatistics.ExcelOp
{
    /// <summary>
    /// 要在nuget里安装EPPlus：Install-Package EPPlus;
    /// excel必须是.xlsx格式,旧的.xls格式不支持;
    /// 整理出的测试项顺序必须和GRR的excel里的sheet顺序一致，否则会错位;  
    /// </summary>
    internal class ExcelOperation
    {
        //从测试项中提取数据
        public async Task ExtractDataFromTestItem(string WorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage)
        {
            int numSN = ParaTestItem.NumSN;//SN个数:8
            int stRow = ParaTestItem.StartRow;//数据源行开始:9
            int stCol = ParaTestItem.StartCol;//数据源列开始:1
            int stRowDest = ParaTestItem.StartRowDest;//目标行开始:1
            int stColDest = ParaTestItem.StartColDest;//目标列开始:2
            int repeat = ParaTestItem.Span;//单个SN的测试次数，即单个SN测试项跨度单元格数量:9
            int count = ParaTestItem.TotalItemCount;// test item count:229
            string fromSheet = ParaTestItem.FromSheet;
            string toSheet = ParaTestItem.ToSheet;
            numSN += 2;// 添加标题行和空行
            int num = 0;// row count

            await Task.Run(() => { /* 空操作 */ }); // 使用 Task.Run 来启动一个无操作的异步任务
            try 
            {
                if (!File.Exists(WorkbookPath))
                {
                    logMessage.Message += $"文件：{WorkbookPath}不存在\r\n";
                    return;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文                                                                  
                FileInfo workbookPath = new FileInfo(WorkbookPath);// 打开工作簿

                using (ExcelPackage Package = new ExcelPackage(workbookPath)) // 打开目标文件
                {
                    // 获取工作表
                    ExcelWorksheet sourceSheet = Package.Workbook.Worksheets[fromSheet];  // Sheet1
                    ExcelWorksheet destSheet = Package.Workbook.Worksheets[toSheet];  // Sheet2
                                                                                      //for (var i = 0; i < 5; i++)
                    for (int i = 0; i < count; i++) // total:i = 229
                    {
                        CopyRange(sourceSheet, stRow + i, stCol, stRow + i, stCol, destSheet, stRowDest + i * numSN, stColDest);
                        for (int j = 0; j < repeat+1; j++) // SN1~SN9
                        {
                            // 拷贝区域 1:fromSheet numSN 个test itme 粘贴到 toSheet中
                            CopyRange(sourceSheet, stRow + i, stCol + j * repeat + 1, stRow + i, stCol + (j + 1) * repeat, destSheet, stRowDest + ++num, stColDest + 1);
                        }
                        num++;
                        //Console.WriteLine($"序号：{i + 1} 数据提取中......，提取test item 个数:{i + 1}, 剩余test item个数: {count - (i + 1)}, item name: {sourceSheet.Cells[stRow + i, stCol].Value}\r\n");
                        logMessage.Message += $"序号：{i + 1,-6} 数据提取中......，提取test item 个数:{i + 1}, 剩余test item个数: {count - (i + 1)}, item name: {sourceSheet.Cells[stRow + i, stCol].Value}\r\n";
                    }
                    // 保存目标文件
                    Package.Save();
                }
            }
            catch(Exception ex)
            {
                logMessage.Message += "测试项数据 提取 失败：" + ex.ToString() + "\r\n";
            }
            //Console.WriteLine("测试项数据提取完成！---------------------------------------------------------------------");
            logMessage.Message += "测试项数据 提取 完成！---------------------------------------------------------------------\r\n";
        }
        //将提取数据拷贝粘贴到GRR module
        public async Task PasteToGRRModuleFromExtractData(string sourceWorkbookPath, string targetWorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage)
        {
            int numSN = ParaTestItem.NumSN;//SN个数
            int stRow = ParaTestItem.StartRow;//数据源行开始:2
            int stCol = ParaTestItem.StartCol;//数据源列开始:3
            int stRowDest = ParaTestItem.StartRowDest;//目标行开始:9
            int stColDest = ParaTestItem.StartColDest;//目标列开始:3
            int TrialsCount= ParaTestItem.Span;//模板单组列数量:3
            int count = ParaTestItem.TotalItemCount;// test item count:229
            string fromSheet = ParaTestItem.FromSheet;
            //string toSheet = ParaTestItem.ToSheet;
            numSN += 2;// 添加标题行和空行
                      
            try
            {
                if (!File.Exists(sourceWorkbookPath))
                {
                    logMessage.Message += $"文件：{sourceWorkbookPath}不存在\r\n";
                    return;
                }
                if (!File.Exists(targetWorkbookPath))
                {
                    logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n";
                    return;
                }

                string[] sheetName = await GetSheetName(targetWorkbookPath, false).ConfigureAwait(false);//get target sheet nam

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文                                                                           
                FileInfo sourceFile = new FileInfo(sourceWorkbookPath);// 打开源工作簿和目标工作簿
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);

                using (ExcelPackage sourcePackage = new ExcelPackage(sourceFile)) // 打开源文件
                using (ExcelPackage destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    // 获取工作表
                    ExcelWorksheet sourceSheet = sourcePackage.Workbook.Worksheets[fromSheet];  // Sheet1
                                                                                                //for (var i = 0; i < 5; i++)
                    for (var i = 0; i < sheetName.Length; i++)
                    {
                        ExcelWorksheet destSheet = destPackage.Workbook.Worksheets[sheetName[i]];  // Sheet2

                        for (int j = 0; j < TrialsCount; j++)
                        {
                            // 拷贝区域 1: Source Sheetxx 的 C3:E9, F3:H9, I3:K9 等 到 Dest Sheetxx 的 C3:xx, G3:xx, K3:xx
                            CopyRange(sourceSheet, stRow + i * numSN, stCol + j * 3, stRow + i * numSN + (numSN - 3), stCol + j * 3 + 2, destSheet, stRowDest, stColDest + j * 4);
                        }
                        //Console.WriteLine( $"序号：{i +1} 数据拷贝到GRR模板中......，拷贝sheet 个数:{i + 1}, 剩余sheet个数: {sheetName.Length - (i + 1)}, sheet name: {sheetName[i]}\r\n");
                        logMessage.Message += $"序号：{i +1,-6} 数据拷贝到GRR模板中......，拷贝sheet 个数:{i + 1}, 剩余sheet个数: {sheetName.Length - (i + 1)}, sheet name: {sheetName[i]}\r\n";
                    }
                    // 保存目标文件
                    destPackage.Save();
                }
            }
            catch (Exception ex)
            {
                logMessage.Message += "提取数据 拷贝到 GRR失败：" + ex.ToString() + "\r\n";
            }
            //Console.WriteLine("提取数据 拷贝到 GRR模板完成！---------------------------------------------------------------------");
            logMessage.Message += "提取数据 拷贝到 GRR模板完成！---------------------------------------------------------------------\r\n";
        }
        //将limit数据拷贝粘贴到GRR module
        public async Task PasteToGRRModuleFromLimit(string sourceWorkbookPath, string targetWorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage)
        {
            int stRow = ParaTestItem.StartRow;//数据源行开始:2
            int stCol = ParaTestItem.StartCol;//数据源列开始:3
            int stRowDest = ParaTestItem.StartRowDest;//目标行开始:9
            int stColDest = ParaTestItem.StartColDest;//目标列开始:3
            //int limitNum = 2;//limit 数量：L + H:2
            //int count = ParaTestItem.TotalItemCount;// test item count:229
            string fromSheet = ParaTestItem.FromSheet;



            try
            {
                if (!File.Exists(sourceWorkbookPath))
                {
                    logMessage.Message += $"文件：{sourceWorkbookPath}不存在\r\n";
                    return;
                }
                if (!File.Exists(targetWorkbookPath))
                {
                    logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n";
                    return;
                }

                string[] sheetName = await GetSheetName(targetWorkbookPath, false).ConfigureAwait(false);//get target sheet name

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文                                                                    
                FileInfo sourceFile = new FileInfo(sourceWorkbookPath);// 打开源工作簿和目标工作簿
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);

                using (ExcelPackage sourcePackage = new ExcelPackage(sourceFile)) // 打开源文件
                using (ExcelPackage destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    // 获取工作表
                    ExcelWorksheet sourceSheet = sourcePackage.Workbook.Worksheets[fromSheet];  // Sheet1
                                                                                                //for (var i = 0; i < 5; i++)
                    for (var i = 0; i < sheetName.Length; i++)
                    {
                        ExcelWorksheet destSheet = destPackage.Workbook.Worksheets[sheetName[i]];  // Sheet2
                                                                                                   // 拷贝区域 1: Source Sheetxx 的 C3:E9, F3:H9, I3:K9 等 到 Dest Sheetxx 的 C3:xx, G3:xx, K3:xx
                        CopyRange(sourceSheet, stRow, stCol + i, stRow + 1, stCol + i, destSheet, stRowDest, stColDest);

                        logMessage.Message += $"序号：{i + 1,-6} limit 拷贝到GRR模板中......，拷贝sheet 个数:{i + 1}, 剩余sheet个数: {sheetName.Length - (i + 1)}, sheet name: {sheetName[i]}\r\n";
                    }
                    // 保存目标文件
                    destPackage.Save();
                }
            }
            catch (Exception ex)
            {
                logMessage.Message += "limit 拷贝到 GRR失败：" + ex.ToString() + "\r\n";
            }
            //Console.WriteLine("提取数据 拷贝到 GRR模板完成！---------------------------------------------------------------------");
            logMessage.Message += "提取数据 拷贝到 GRR模板完成！---------------------------------------------------------------------\r\n";
        }


        #region internal + private
        // 新建sheet
        internal async Task CreatSheet(string targetWorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage, bool before = false)
        {
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }
                string[] sheetName = ParaTestItem.SheetName;
                var sheetNa = await GetSheetName(targetWorkbookPath, false).ConfigureAwait(false);
                string summary = sheetNa.Last();
                
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        var destSheet = destPackage.Workbook.Worksheets.Add(sheetName[i]);  // // 创建一个新的工作表，名称为 "NewSheet"

                       
                        var workbook = destPackage.Workbook; // 获取工作表集合
                        var targetSheet = workbook.Worksheets[summary];                       
                        int targetSheetIndex = targetSheet.Index;// 获取工作表的索引


                        if (before)
                        {
                            workbook.Worksheets.MoveBefore(destSheet.Index, targetSheetIndex - i);// 将工作表移到目标位置前（插入位置）
                        }
                        else
                        {
                            workbook.Worksheets.MoveAfter(destSheet.Index, targetSheetIndex - i);// 将工作表移到目标位置后（插入位置）
                        }
                        //var newSheet = package.Workbook.Worksheets.Add("NewSheet");
                        logMessage.Message += $"序号{i + 1,-6}, 创建工作表 '{sheetName[i]}'\r\n";
                    }
                    destPackage.Save();
                }

                Console.WriteLine(logMessage.Message += $"创建工作表完成！\r\n");
            }
            catch (Exception ex)
            {
                logMessage.Message += $"{ex.ToString()}\r\n";
            }

        }
        // rename sheet
        internal async Task RenameSheet(string targetWorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage)
        {
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }
                string[] newSheetName = ParaTestItem.SheetName;
                string[] oldSheetName = await GetSheetName(targetWorkbookPath, false).ConfigureAwait(false);
                int len = Math.Min(newSheetName.Length, oldSheetName.Length);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打开源工作簿和目标工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    

                    for (int i = 0; i < len; i++)
                    {
                        // 获取指定名称的工作表
                        ExcelWorksheet worksheet = destPackage.Workbook.Worksheets[oldSheetName[i]];

                        if (worksheet != null)
                        {                          
                            worksheet.Name = newSheetName[i];// 重命名工作表
                            logMessage.Message += $"序号{i + 1,-6}, '{oldSheetName[i]} ' 重命名为 --> ‘{newSheetName[i]}'\r\n";
                        }
                        else
                        {
                            logMessage.Message += $"工作表 '{oldSheetName}' 不存在！\r\n";
                        }
                    }
                    destPackage.Save();
                }

                Console.WriteLine(logMessage.Message += $"重命名工作表完成！\r\n");
            }
            catch (Exception ex)
            {
                logMessage.Message += $"{ex.ToString()}\r\n";
            }

        }
        //删除sheet
        internal async Task DeleteSheet(string targetWorkbookPath, ParametersTestItem ParaTestItem, LogMessage logMessage)
        {
            await Task.Run(() => { /* 空操作 */ }); // 使用 Task.Run 来启动一个无操作的异步任务
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }

                //string outputFilePath = @"E:\labview\other prj\IGBT cplusplus dll\MSA1\sheetname.txt";
                string[] sheetName = ParaTestItem.SheetName;

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    // 获取工作表集合
                    var workbook = destPackage.Workbook;
                   
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        var sheetToRemove = workbook.Worksheets[sheetName[i]];
                        if (sheetToRemove != null)
                        {
                            workbook.Worksheets.Delete(sheetToRemove); // 删除工作表
                            logMessage.Message += $"序号{i + 1,-6}, 工作表 '{sheetToRemove}' 已删除\r\n";
                        }
                        else
                        {
                            logMessage.Message += $"未找到工作表 '{sheetToRemove}'\r\n";
                        }
                    }

                    destPackage.Save();
                }

                logMessage.Message += "删除工作表完成！\r\n";
            }
            catch (Exception ex)
            {
                logMessage.Message += $"{ex.ToString()}\r\n";
            }

        }
        //删除sheet
        internal async Task DeleteSheet(string targetWorkbookPath, int reserveSheetCount,LogMessage logMessage)
        {
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }

                string[] sheetName = await GetSheetName(targetWorkbookPath, false).ConfigureAwait(false);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    // 获取工作表集合
                    var workbook = destPackage.Workbook;

                    // 删除名为 "Sheet1" 的工作表
                    if (sheetName.Length <= reserveSheetCount)
                    {
                        logMessage.Message += $"工作表小于 {reserveSheetCount} 个\r\n";
                        return;
                    }
                    for (int i = reserveSheetCount; i < sheetName.Length; i++)
                    {
                        var sheetToRemove = workbook.Worksheets[sheetName[i]];
                        if (sheetToRemove != null)
                        {
                            workbook.Worksheets.Delete(sheetToRemove); // 删除工作表
                            logMessage.Message += $"序号{i - reserveSheetCount + 1,-6}, 工作表 '{sheetToRemove}' 已删除\r\n";
                        }
                        else
                        {
                            logMessage.Message += $"未找到工作表 '{sheetToRemove}'\r\n";
                        }
                    }

                    destPackage.Save();
                }

                logMessage.Message += "删除工作表完成！\r\n";
            }
            catch(Exception ex)
            {
                logMessage.Message += $"{ex.ToString()}\r\n";
            }
            
        }
        //删除range单元格
        internal async Task DeleteRangeData(string targetWorkbookPath, ParametersTestItem ParaTestItem)
        {
            int stRow = ParaTestItem.StartRow;//数据源行开始:17
            int stCol = ParaTestItem.StartCol;//数据源列开始:3
            int endRow = ParaTestItem.EndRow;//数据源行结束:17
            int endCol = ParaTestItem.EndtCol;//数据源列结束:14
            string[] sheetName = ParaTestItem.SheetName;

            await Task.Run(() => { /* 空操作 */ }); // 使用 Task.Run 来启动一个无操作的异步任务
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    //Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        var destSheet = destPackage.Workbook.Worksheets[sheetName[i]];  // Sheet2
                        if (destSheet != null)
                        {
                            DeleteRangeData(destSheet, startRow: stRow, startCol: stCol, endRow: endRow, endCol: endCol);
                        }
                        else
                        {
                            Console.WriteLine("未找到工作表 ");
                        }
                    }

                    destPackage.Save();
                }
                Console.WriteLine("删除数据完成！");
            }
            catch
            {
                throw;
            }
            
        }
        //复制粘贴range单元格
        internal async Task CopyRangePaste(string targetWorkbookPath, ParametersTestItem ParaTestItem)
        {
            int stRow = ParaTestItem.StartRow;//数据源行开始
            int stCol = ParaTestItem.StartCol;//数据源列开始
            int endRow = ParaTestItem.EndRow;//数据源行结束
            int endCol = ParaTestItem.EndtCol;//数据源列结束
            int stRowDest = ParaTestItem.StartRowDest;//目标行开始
            int stColDest = ParaTestItem.StartColDest;//目标列开始
            string[] sheetName = ParaTestItem.SheetName;

            await Task.Run(() => { /* 空操作 */ }); // 使用 Task.Run 来启动一个无操作的异步任务
            try
            {
                if (!File.Exists(targetWorkbookPath))
                {
                    //Console.WriteLine(logMessage.Message += $"文件：{targetWorkbookPath}不存在\r\n");
                    return;
                }
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;// 设置 EPPlus 许可证上下文               
                FileInfo destinationFile = new FileInfo(targetWorkbookPath);// 打工作簿
                using (var destPackage = new ExcelPackage(destinationFile)) // 打开目标文件
                {
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        var destSheet = destPackage.Workbook.Worksheets[sheetName[i]];  // Sheet2
                        if (destSheet != null)
                        {
                            CopyRange(destSheet, startRow: stRow, startCol: stCol, endRow: endRow, endCol: endCol, destSheet, destStartRow: stRowDest, destStartCol: stColDest);
                        }
                        else
                        {
                            Console.WriteLine("未找到工作表 ");
                        }

                    }
                    destPackage.Save();
                }
                Console.WriteLine("数据拷贝完成！");
            }
            catch { throw; }
            
        }
        // 获取 Excel 文件所有sheet名,导出生成.txt,去除 Summary sheet。 
        internal async Task<string[]> GetSheetName(string excelFilePaht, bool allSheet, string outputFilePath = null)
        {
            string[] sheetNames;

            try
            {
                if (!File.Exists(excelFilePaht))
                {
                    return null;// new string[1] { $"文件: {excelFilePaht}不存在" };
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // 设置 EPPlus 许可证上下文// 或者 LicenseContext.Commercial             
                using (var package = new ExcelPackage(new FileInfo(excelFilePaht)))// 确保使用 EPPlus 许可证
                {
                    // 获取工作簿中的所有工作表
                    var worksheets = package.Workbook.Worksheets;
                    //sheetNames = worksheets.Select(x => x.Name).ToArray();
                    sheetNames = await Task.Run(() => { return worksheets.Select(x => x.Name).ToArray();});
                    Array.Reverse(sheetNames);
                    if (!allSheet)// 移除最后一个sheet，保留n-1个sheet
                    {
                        Array.Resize(ref sheetNames, sheetNames.Length - 1);// delete Summary sheet
                    }
                }

                if (outputFilePath != null) // save to .txt
                {
                    outputFilePath = GetTextFileName(outputFilePath);
                    using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                    {
                        writer.WriteLine($"sheet name:all count: {sheetNames.Length}");
                        foreach (var sheet in sheetNames)
                        {
                            writer.WriteLine($"{sheet}");
                        }
                    }
                }
            }
            catch
            {
                throw;
            }
            return sheetNames;
        }

        // 生产excel VBS脚本，提取同一测试项的值（测试：span 次）
        internal async Task CreatVBScript(string outputFilePath, LogMessage logMessage)
        {
            try
            {

                outputFilePath = GetTextFileName(outputFilePath);
                int span = 9;//单个SN的一个测试项的测试次数，即单个测试项跨度单元格数量
                // 打开文件流进行写入
                using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    string fromSheet = "SortSelectTrans";
                    string toSheet = "toSheet";
                    int num = 0;
                    const int count = 235;
                    await Task.Run(() => 
                    {
                        for (int i = 0; i < count; i++) // total:i = 235
                        {

                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('A' + 0)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"A{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B' + 0 * 1)}{span + i}:{(char)('B' + span * 1 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B' + span * 1)}{span + i}:{(char)('B' + span * 2 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B' + span * 2)}{span + i}:{(char)('A')}{(char)('C' + span * 0 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('A')}{(char)('C' + span * 0)}{span + i}:{(char)('A')}{(char)('C' + span * 1 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('A')}{(char)('C' + span * 1)}{span + i}:{(char)('A')}{(char)('C' + span * 2 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('A')}{(char)('C' + span * 2)}{span + i}:{(char)('B')}{(char)('D' + span * 0 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B')}{(char)('D' + span * 0)}{span + i}:{(char)('B')}{(char)('D' + span * 1 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B')}{(char)('D' + span * 1)}{span + i}:{(char)('B')}{(char)('D' + span * 2 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{(char)('B')}{(char)('D' + span * 2)}{span + i}:{(char)('C')}{(char)('E' + span * 0 - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");

                            num++;
                            if (num % (10 * 10) == 0)
                            {
                                writer.WriteLine($"DelayWithParameter (SecondValue)");
                                writer.WriteLine($"\'end count:{(num / 10)}");
                            }
                            //logMessage.Message += $"{fileName}\t\t\t\t\t\t\t\t{videoDuration}\t\t{fileSize} MB" + "\n";
                            Console.WriteLine(num);
                        }
                        writer.WriteLine($"\'end count:{(num / 10)}");
                    });
                   
                }

                //Console.WriteLine("视频文件信息已保存到 " + outputFilePath);
                //logMessage.Message += "视频文件信息已保存到 " + outputFilePath + "\n";
            }
            catch (Exception ex)
            {
                Console.WriteLine(logMessage.Message += "发生错误: " + ex.Message);
                logMessage.Message += "发生错误: " + ex.Message + "\n";
            }
        }
        // 生产excel VBS脚本，提取同一测试项的值（测试：span 次）
        internal async Task CreatVBScript(string outputFilePath, LogMessage logMessage, char startCol)
        {
            try
            {
                outputFilePath = GetTextFileName(outputFilePath);

                int span = 9;//单个SN的测试次数，即单个SN测试项跨度单元格数量
                int count = 235;// test item count
                startCol = (char)(startCol + 1); // 从 test item name 开始算，比如2.2.1开始的列号
                string fromSheet = "SortSelectTrans";
                string toSheet = "toSheet";
                int num = 0;// row count

                // 打开文件流进行写入
                using (StreamWriter writer = new StreamWriter(outputFilePath, false, Encoding.UTF8))
                {
                    await Task.Run(() =>
                    {
                        for (int i = 0; i < count; i++) // total:i = 235
                        {
                            int colCnt = 0;
                            writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{ConvertToExcelColumn(startCol - 1 + colCnt * span)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"A{2 + num}\")");// test item name
                            for (int j = 0; j < 9; j++) // SN1~SN9
                            {
                                writer.WriteLine($"Sheets(\"{fromSheet}\").Range(\"{ConvertToExcelColumn(startCol + colCnt * span)}{span + i}:" +
                                    $"{ConvertToExcelColumn(startCol + ++colCnt * span - 1)}{span + i}\").Copy Destination:=Sheets(\"{toSheet}\").Range(\"B{2 + ++num}:J{2 + num}\")");
                            }
                            num++;
                            if (num % (10 * 10) == 0)// 10 个 test item 后标记一下test item 个数
                            {
                                writer.WriteLine($"DelayWithParameter (SecondValue)");
                                writer.WriteLine($"\'end count:{(num / 10)}");
                            }
                            Console.WriteLine(num);
                        }
                        writer.WriteLine($"\'end count:{(num / 10)}");
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(logMessage.Message += "发生错误: " + ex.Message);
                logMessage.Message += "发生错误: " + ex.Message + "\n";
            }
        }

        #region Excel 
        // 拷贝指定范围的Value方法
        private void CopyRange(ExcelWorksheet sourceSheet, int startRow, int startCol, int endRow, int endCol, ExcelWorksheet destSheet, int destStartRow, int destStartCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    // 获取源单元格的公式（如果有公式的话）
                    var formula = sourceSheet.Cells[row, col].Formula;

                    // 如果源单元格有公式，则将公式复制到目标单元格
                    if (!string.IsNullOrEmpty(formula))
                    {
                        // 根据目标位置调整公式
                        string adjustedFormula = AdjustFormulaForNewLocation(formula, row, col, destStartRow, destStartCol);

                        // 将调整后的公式复制到目标单元格
                        destSheet.Cells[destStartRow + (row - startRow), destStartCol + (col - startCol)].Formula = adjustedFormula;
                    }
                    else
                    {
                        // 如果没有公式，则复制数值
                        var value = sourceSheet.Cells[row, col].Value;
                        destSheet.Cells[destStartRow + (row - startRow), destStartCol + (col - startCol)].Value = value;
                    }
                }
            }
        }
        // 根据目标单元格的位置调整公式
        private string AdjustFormulaForNewLocation(string formula, int sourceRow, int sourceCol, int destRow, int destCol)
        {
            // 公式的正则表达式，匹配单元格引用（例如 A1、$A$1、A$1、$A1）
            string pattern = @"[A-Z]+\d+";

            // 替换公式中的单元格引用
            string adjustedFormula = Regex.Replace(formula, pattern, match =>
            {
                // 获取源单元格的行列位置
                string cellReference = match.Value;
                int cellRow = int.Parse(Regex.Match(cellReference, @"\d+").Value);
                string cellColumn = Regex.Match(cellReference, @"[A-Z]+").Value;

                // 计算行和列的偏移量
                int rowOffset = destRow - sourceRow;
                int colOffset = destCol - sourceCol;

                // 计算新的行和列引用
                int newRow = cellRow + rowOffset;
                string newColumn = GetColumnName(GetColumnIndex(cellColumn) + colOffset);

                // 返回调整后的单元格引用
                return newColumn + newRow;
            });

            return adjustedFormula;
        }

        // 根据列名返回列索引
        private int GetColumnIndex(string columnName)
        {
            int columnIndex = 0;
            foreach (char c in columnName)
            {
                columnIndex = (columnIndex * 26) + (c - 'A' + 1);
            }
            return columnIndex;
        }

        // 根据列索引返回列名
        private string GetColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                columnIndex--;
                columnName = (char)(columnIndex % 26 + 'A') + columnName;
                columnIndex /= 26;
            }
            return columnName;
        }

        // 删除指定范围的数据
        private void DeleteRangeData(ExcelWorksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            for (int row = startRow; row <= endRow; row++)
            {
                for (int col = startCol; col <= endCol; col++)
                {
                    // 将单元格的值设置为 null，清除单元格数据
                    sheet.Cells[row, col].Value = null;
                }
            }
        }
        #endregion

        //获取或创建.txt文件
        private string GetTextFileName(string fullPath)
        {
            // 检查txt文件是否存在
            if (!fullPath.Contains(".txt"))
            {
                // 获取当前日期和时间并格式化为字符串
                string dateTimeString = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                // 创建文件名，包含日期和时间
                string fileName = $"{dateTimeString}.txt";

                // 设置文件路径（可以根据需要修改路径）
                fullPath = Path.Combine(fullPath + "\\", fileName);
                try
                {
                    // 创建并写入文件
                    File.WriteAllText(fullPath, "这是一个测试文件内容。");

                    // 输出文件路径
                    Console.WriteLine($"文件已创建: {fullPath}");
                }
                catch (Exception ex)
                {
                    // 如果出现错误，输出异常信息
                    Console.WriteLine($"创建文件时发生错误: {ex.Message}");
                }
            }

            if (!File.Exists(fullPath))
            {
                // 如果文件不存在，则创建文件并写入初始内容
                using (StreamWriter sw = File.CreateText(fullPath))
                {
                    sw.WriteLine("这是新创建的文件。");
                    sw.WriteLine("文件创建时间: " + DateTime.Now.ToString());
                }

                Console.WriteLine("文件已创建并写入内容。");
            }
            return fullPath;
        }
        //ASCII 转换为 Excel 列标
        private string ConvertToExcelColumn(int number)
        {
            string result = "";
            number -= 64;

            while (number > 0)
            {
                //number--; // 减 1 因为 Excel 标记系统从 1 开始
                int remainder = number % 26; // 取余数
                result = (char)(remainder + 'A' -1) + result; // 将余数转为对应字符并添加到结果
                number /= 26; // 除以 26，进入下一位
            }          
            return result;
        }
        #endregion
    }
}