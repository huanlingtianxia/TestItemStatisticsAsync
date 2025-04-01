//#define INIFILE
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestItemStatisticsAcync.ExcelOp;
using TestItemStatisticsAcync.Ini;

/* 程序包安装
安装EPPlus：Install-Package EPPlus
安装NLog：Install-Package NLog
 */

namespace TestItemStatisticsAcync
{
    public partial class Form1 : Form
    {
        #region Init and FormClose
        public Form1()
        {
            InitializeComponent();
            InitParam();
        }
        private void InitParam()
        {
            
            LogMsg.Message = String.Empty;
            ConfigLog();
            Logger.Info(">>>>>>>>>>程序启动");
            UpdateUIControlFromIniFile();
            UpdateParamFromControl();
            AdjustRichTextBoxSize(251, -1);
            // 取消选中状态并将光标移到文本框末尾
            textB_TargetPath.SelectionStart = textB_TargetPath.Text.Length;
            textB_TargetPath.SelectionLength = 0;
            LogMsg.Message += "初始化完成\r\n";
            UpdateUILog(LogMsg.Message);

        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Logger.Info("程序结束<<<<<<<<<<");
        }
        #endregion

        #region property
        internal ExcelOperation ExcelOp { get; set; } = new ExcelOperation();//Excel 操作
        //IIniReaderWriter iniReaderWriter { get; set; } = new IniReaderWriter();
        //LoggerHander loggerHander { get; set; } = new LoggerHander("log.txt");
        internal ParametersTestItem TestItem { get; set; } = new ParametersTestItem();//从测试log提取数据参数
        internal ParametersTestItem TestItemGRR { get; set; } = new ParametersTestItem();// copy paste 提取数据到GRR module 参数
        internal ParametersTestItem TestItemGRRLimit { get; set; } = new ParametersTestItem();// copy paste Limit到GRR module 参数
        internal LogMessage LogMsg { get; set; } = new LogMessage();// log message
        internal long MaxLogSize { get; set; } = 1 * 1024 * 1024; // 10 MB

        private static Logger? logger; // 定义 logger 属性  
        internal static Logger Logger => logger ?? (logger = LogManager.GetCurrentClassLogger());// 检查 logger 是否为 null，如果是则通过 LogManager 创建新的实例

        #endregion

        #region Control Click event
        //Extract data, copy paste test data to GRR, copy paste limit to GRR
        //Button_Click 是一个事件处理程序，因此它返回 void，这使得它成为 async void 方法。
        //这种情况下，虽然你不能使用 await 等待 Button_Click，但事件处理程序本身会异步执行，不会阻塞 UI 线程。
        private async void btn_SelectSourcePath_Click(object sender, EventArgs e)
        {
            await Task.CompletedTask;  // 模拟异步
            //string path1 = SelectfullPath();
            string path = SelectfullPath();
            if (path != String.Empty)
                textB_SourcePath.Text = path;
        }

        private async void btn_SelectTargetPath_Click(object sender, EventArgs e)
        {
            await Task.CompletedTask;  // 模拟异步
            //string path = SelectPath();
            string path = SelectfullPath();
            if (path != String.Empty)
                textB_TargetPath.Text = path;
        }
       
        private async void btn_ExtractData_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            UpdateParamFromControl();
            await ExcelOp.ExtractDataFromTestItem(TestItem.SourcePath, TestItem, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_PasteToGRR_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            UpdateParamFromControl();
            await ExcelOp.PasteToGRRModuleFromExtractData(TestItemGRR.SourcePath, TestItemGRR.TargetPath, TestItemGRR, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_ExtractSheetToTxt_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            UpdateParamFromControl();
            try
            {
                UpdateParamFromControl();
                string[] str = { "\\" };
                string path = string.Empty;
                string[] pathArr = TestItemGRR.TargetPath.Split(str, StringSplitOptions.None);
                for (int i = 0; i < pathArr.Length - 1; i++)
                {
                    path += pathArr[i] + "\\";
                }
                path += "GRRModuleSheetName.txt";
                string[] sheetName = await ExcelOp.GetSheetName(TestItemGRR.TargetPath, false, path).ConfigureAwait(false);
                LogMsg.Message += "提取GRR module中 sheet name 到GRRModuleSheetName.txt,\r\n" + "path: " + path + "\r\n";
                LogMsg.Message += $"Total sheet count: {sheetName.Length}\r\n";
                LogMsg.Message += $"Count{string.Empty,-5}, sheet name\r\n";
                if (sheetName != null)
                {
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        LogMsg.Message += $"{i + 1,-10}, {sheetName[i]}\r\n";
                    }
                    LogMsg.Message += $"提取sheet name 完成！ ----------------------\r\n";
                }
                else
                {
                    LogMsg.Message += $"未找到工作表\r\n";
                }
                UpdateUILog(LogMsg.Message);
            }
            catch(Exception ex)
            {
                LogMsg.Message += $"异常: {ex.ToString()}\r\n";
                UpdateUILog(LogMsg.Message);
            }
            
        }

        private async void btn_PasteLimit_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            UpdateParamFromControl();
            await ExcelOp.PasteToGRRModuleFromLimit(TestItemGRRLimit.SourcePath, TestItemGRRLimit.TargetPath, TestItemGRRLimit, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
        }
        // update ini file
        private void btn_WriteIni_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            try
            {
                DialogResult result = MessageBox.Show(this, "确定将UI中的数据更新到 Ini 文件?", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                if (result == DialogResult.OK)
                {
                    bool state = UpdateIniFileFromUIControl();
                    LogMsg.Message += state ? "已将UI数据更新到Ini文件\r\n" : "UI数据更新到Ini文件 异常\r\n";
                    MessageBox.Show(LogMsg.Message);// 用户点击“确认”
                }
                else
                {
                    MessageBox.Show(this, LogMsg.Message += "操作已取消\r\n");// 用户点击“取消”
                }
                UpdateUILog(LogMsg.Message);
            }
            catch(Exception ex)
            {
                UpdateUILog(LogMsg.Message += ex.ToString());
            }
        }
        private void btn_ReadIni_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            bool state = UpdateUIControlFromIniFile();
            LogMsg.Message += state ? "read ini file and update to UI Control completed\r\n" : "read ini file failed\r\n";
            UpdateUILog(LogMsg.Message);
        }
        //General: CopyPaste And Delete
        private async void btn_CopyPaste_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            ParametersTestItem para = new ParametersTestItem();
            UpdateParamFromControl(para, true);
            await ExcelOp.CopyRangePaste(para.TargetPath, para, LogMsg).ConfigureAwait(false); // 复制 公式单元格
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_DeleteRange_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            //await Task.Delay(5000);
            ParametersTestItem para = new ParametersTestItem();
            UpdateParamFromControl(para, false);
            await ExcelOp.DeleteRangeData(para.TargetPath, para, LogMsg).ConfigureAwait(false); // 删除 17行单元格，作用域：11.xx测试项
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_DeleteSheet_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            ParametersTestItem parametersTestItem = new ParametersTestItem();
            UpdateParamFromControl(parametersTestItem, false);
            if (parametersTestItem.ReserveSheetCount == -1)
            {
                await ExcelOp.DeleteSheet(parametersTestItem.TargetPath, parametersTestItem, LogMsg).ConfigureAwait(false);//删除SheetName中的工作表
            }
            else
            {
                await ExcelOp.DeleteSheet(parametersTestItem.TargetPath, parametersTestItem.ReserveSheetCount, LogMsg).ConfigureAwait(false);//保留ReserveSheetCount个工作表
            }
            //string[] sheet = ExcelOp.GetSheetName(parametersTestItem.TargetPath,true);
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_CreatSheet_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            ParametersTestItem para = new ParametersTestItem();
            UpdateParamFromControl(para, false);
            await ExcelOp.CreatSheet(para.TargetPath, para, LogMsg).ConfigureAwait(false); // 删除 17行单元格，作用域：11.xx测试项
            UpdateUILog(LogMsg.Message);
        }

        private async void btn_RemaneSheet_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            ParametersTestItem para = new ParametersTestItem();
            UpdateParamFromControl(para, false);
            await ExcelOp.RenameSheet(para.TargetPath, para, LogMsg).ConfigureAwait(false); // 删除 17行单元格，作用域：11.xx测试项
            UpdateUILog(LogMsg.Message);
        }

        private void lab_EnableMaskArrow_Click(object sender, EventArgs e)
        {
            flowLayoutP_Mask.Visible = !flowLayoutP_Mask.Visible;
            if(flowLayoutP_Mask.Visible)
            {
                AdjustRichTextBoxSize(251, -1);
            }
            else
            {
                AdjustRichTextBoxSize(144, 1);
            }
            //Thread.Sleep(100);
        }
        #endregion

        #region private function
        private string SelectPath()
        {
            // 创建一个 FolderBrowserDialog 实例
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();

            // 设置初始路径（可选）
            //folderBrowserDialog.SelectedPath = @"E:\labview\《Labview从入门到精通》视频教程\";

            // 显示文件夹选择对话框
            DialogResult result = folderBrowserDialog.ShowDialog();

            // 如果用户选择了一个文件夹
            if (result == DialogResult.OK)
            {
                // 获取用户选择的文件夹路径
                string folderPath = folderBrowserDialog.SelectedPath;

                // 将文件夹路径显示到 TextBox 中
                return folderPath;
            }
            return string.Empty;
        }
        private string SelectfullPath()
        {
            // 创建 OpenFileDialog 实例
            OpenFileDialog openFileDialog = new OpenFileDialog();

            // 设置初始目录和过滤器（可选）
            //openFileDialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            openFileDialog.InitialDirectory = "C:\\";  // 你可以设置你希望打开的默认目录
            openFileDialog.Filter = "所有文件 (*.*)|*.*";  // 允许选择所有文件类型
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            string fullPath = string.Empty;

            // 显示对话框并检查用户是否选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取选择的文件路径
                fullPath = openFileDialog.FileName;
            }
            return fullPath;
        }
        private string SaveFile()
        {
            // 创建 SaveFileDialog 实例
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            // 设置初始目录和过滤器（可选）
            saveFileDialog.InitialDirectory = "C:\\";
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;
            string fullPath = string.Empty;
            // 显示对话框并检查用户是否选择了保存路径
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 获取选择的保存文件路径
                fullPath += saveFileDialog.FileName;

                // 显示文件路径
                //filePathTextBox.Text = selectedFile;
            }
            return fullPath;
        }
        private void MessagePop(string aa, bool mode = false)
        {
            if(mode)
            {
                DialogResult result = MessageBox.Show(this, "是否保存更改?", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // 用户点击“是”
                    MessageBox.Show("保存中...");
                }
                else
                {
                    // 用户点击“否”
                    MessageBox.Show("未保存");
                }
            }
            else
            {
                DialogResult result = MessageBox.Show(this, "确定要删除文件吗?", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                if (result == DialogResult.OK)
                {
                    // 用户点击“确认”
                    MessageBox.Show("文件已删除");
                }
                else
                {
                    // 用户点击“取消”
                    MessageBox.Show("操作已取消");
                }
            }
        }
        // config log
        private void ConfigLog()
        {  
            var config = new LoggingConfiguration();// 明确配置 NLog，包含内部日志设置             
            var logfile = new FileTarget("logfile")// 创建文件目标
            {
                FileName = "${basedir}/logs/log.txt",
                Layout = "${longdate} ${level} ${message}"
            };

            config.AddTarget(logfile);
            config.LoggingRules.Add(new LoggingRule("*", LogLevel.Debug, logfile));

            // 定义内部日志文件  
            //ConfigSetting.SetInternalLogLevel(LogLevel.Debug);
            //ConfigSetting.SetInternalLogFile("${basedir}/nlog-internal.log");

            LogManager.Configuration = config;// 应用配置 

            string logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
            Directory.CreateDirectory(logDir); // 确保 logs 目录存在 
            string logFilePath = Path.Combine(logDir, "log.txt");
            CheckLogFileSizeAndRecreate(logFilePath);
        }
        private void CheckLogFileSizeAndRecreate(string logFilePath)
        {
            // 如果日志文件存在，检查大小  
            if (File.Exists(logFilePath))
            {
                FileInfo fileInfo = new FileInfo(logFilePath);
                if (fileInfo.Length > MaxLogSize)
                {
                    DialogResult result = MessageBox.Show(this, $"日志文件 {logFilePath} 超过 {MaxLogSize / (1024 * 1024)}MB，是否删除并创建新日志？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                    if (result == DialogResult.OK)
                    {
                        // 删除日志文件  
                        File.Delete(logFilePath);
                        Console.WriteLine(LogMsg.Message += $"日志文件 {logFilePath} 超过 {MaxLogSize / (1024 * 1024)}MB，已被删除。\r\n");

                        // 重新创建日志文件  
                        using (File.Create(logFilePath))
                        {
                            // 空文件创建，不需要任何操作  
                        }
                        Console.WriteLine(LogMsg.Message += $"日志文件 {logFilePath} 已重新创建。\r\n");
                        MessageBox.Show(LogMsg.Message);// 用户点击“确认”
                    }
                    else
                    {
                        MessageBox.Show(this, LogMsg.Message += "操作已取消\r\n");// 用户点击“取消”
                    }
                }
            }
        }
        private void AdjustRichTextBoxSize(int newHeight, int stretch)
        {
            // 获取当前的尺寸和位置  
            int currentHeight = richTB_Log.Height;
            int dtHeight = Math.Abs(newHeight - currentHeight) * stretch; // 向上拉伸或缩放 xx 像素  

              
            richTB_Log.Size = new Size(richTB_Log.Width, newHeight);// 更新 RichTextBox 的高度

            Point currentLocation = richTB_Log.Location;
            richTB_Log.Location = new Point(currentLocation.X, currentLocation.Y + dtHeight); // 向上或向下移动 xx 像素
            //this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + (251 - 144));// 可选：更新窗口尺寸以适应新大小  
        }
        //GRR parameters
        private void UpdateParamFromControl()
        {
            try
            {
                //Extract data
                TestItem.StartRow = int.Parse(textB_StartRow.Text);
                TestItem.StartCol = int.Parse(textB_StartCol.Text);
                TestItem.StartRowDest = int.Parse(textB_StartRowDest.Text);
                TestItem.StartColDest = int.Parse(textB_StartColDest.Text);
                TestItem.Repeat = int.Parse(textB_Repeat.Text);
                TestItem.NumSN = int.Parse(textB_NumSN.Text);
                TestItem.TotalItemCount = int.Parse(textB_TotalItem.Text);
                TestItem.FromSheet = textB_FromSheet.Text;
                TestItem.ToSheet = textB_ToSheet.Text;
                TestItem.SourcePath = textB_SourcePath.Text;

                //Paste to GRR module test item
                TestItemGRR.StartRow = int.Parse(textB_StartRow_GRR.Text);
                TestItemGRR.StartCol = int.Parse(textB_StartCol_GRR.Text);
                TestItemGRR.StartRowDest = int.Parse(textB_StartRowDest_GRR.Text);
                TestItemGRR.StartColDest = int.Parse(textB_StartColDest_GRR.Text);
                TestItemGRR.Repeat = int.Parse(textB_TrialsCount_GRR.Text);
                TestItemGRR.NumSN = int.Parse(textB_NumSN_GRR.Text);
                //TestItemGRR.TotalItemCount = int.Parse(textB_TotalItem_GRR.Text);
                TestItemGRR.FromSheet = textB_FromSheet_GRR.Text;
                TestItemGRR.SourcePath = textB_SourcePath.Text;
                TestItemGRR.TargetPath = textB_TargetPath.Text;

                //Paste to GRR module Limit
                TestItemGRRLimit.StartRow = int.Parse(textB_StartRowLimit.Text);
                TestItemGRRLimit.StartCol = int.Parse(textB_StartColLimit.Text);
                TestItemGRRLimit.StartRowDest = int.Parse(textB_StartRowDestLimit.Text);
                TestItemGRRLimit.StartColDest = int.Parse(textB_StartColDestLimit.Text);
                TestItemGRRLimit.FromSheet = textB_FromSheetLimit.Text;
                TestItemGRRLimit.SourcePath = textB_SourcePath.Text;
                TestItemGRRLimit.TargetPath = textB_TargetPath.Text;

                // option

                // common
            }
            catch (Exception ex)
            {
                LogMsg.Message += "输入控件不是数字：" + ex.ToString() + "\r\n";
                UpdateUILog(LogMsg.Message);
            }
            
        }
        // General parameters
        private void UpdateParamFromControl(ParametersTestItem parameters, bool copyPast)
        {
            try
            {
                parameters.TargetPath = textB_ExcelPath.Text;
                parameters.SheetName = richT_SheetName.Text.Trim().Split(new string[1] { "\n" }, StringSplitOptions.None);
                string[] cnt = textB_ReserveSheetCount.Text.Trim().Split(new string[1] { ":" }, StringSplitOptions.None);
                parameters.ReserveSheetCount = (cnt.Length == 2 && cnt[0] == ":") ? int.Parse(cnt[1]) : -1;
                int[] CopyPastePara = textB_CopyPastePara.Text.Trim().Split(new string[1] { "," }, StringSplitOptions.None).Select(item => int.Parse(item)).ToArray();
                int[] DeletePara = textB_DeletePara.Text.Trim().Split(new string[1] { "," }, StringSplitOptions.None).Select(item => int.Parse(item)).ToArray();
                //parameters.CopyPastePara = textB_CopyPastePara.Text;
                //parameters.DeletePara = textB_DeletePara.Text;
                //parameters.SheetName = richT_SheetName.Text;
                if (copyPast)
                {
                    for (int i = 0; i < CopyPastePara.Length; i++)
                    {
                        switch (i)
                        {
                            case 0: parameters.StartRow = CopyPastePara[i]; break;
                            case 1: parameters.StartCol = CopyPastePara[i]; break;
                            case 2: parameters.EndRow = CopyPastePara[i]; break;
                            case 3: parameters.EndtCol = CopyPastePara[i]; break;
                            case 4: parameters.StartRowDest = CopyPastePara[i]; break;
                            case 5: parameters.StartColDest = CopyPastePara[i]; break;
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < DeletePara.Length; i++)
                    {
                        switch (i)
                        {
                            case 0: parameters.StartRow = DeletePara[i]; break;
                            case 1: parameters.StartCol = DeletePara[i]; break;
                            case 2: parameters.EndRow = DeletePara[i]; break;
                            case 3: parameters.EndtCol = DeletePara[i]; break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMsg.Message += "输入控件不是数字：" + ex.ToString() + "\r\n";
                UpdateUILog(LogMsg.Message);
            }
        }
        // update richtextbox log,BeginInvoke 是异步执行的，不会阻塞当前线程，而 Invoke 是同步执行的，会等待操作完成。
        private void UpdateUILog(string msg)
        {
            if (richTB_Log.InvokeRequired)
            {
                // 调用 UI 线程来更新 RichTextBox
                richTB_Log.BeginInvoke(new Action(() => {
                    richTB_Log.AppendText(msg);
                    Logger.Info(msg);
                }));
                
            }
            else
            {
                richTB_Log.AppendText(msg);
                Logger.Info(msg);
            }
        }
        // init richtextbox log
        private void InitUILog(string msg)
        {
            richTB_Log.Clear();
            LogMsg.Message = string.Empty;
            UpdateUILog(msg);
        }

        // update INI file from UI Control
        private bool UpdateIniFileFromUIControl()
        {
            #region INI_FILE_WRITE
            string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);//获取当前执行文件所在路径
            string path = executablePath + "\\Ini\\TestItemStatistics.ini";
            if (!File.Exists(path))
            {
                LogMsg.Message += $"文件：'{path}' 不存在\r\n";
                return false;
            }
            IniFile iniFile = new IniFile(path);

            // 写入数据  
            // ExtractData
            iniFile.Write("GRR_ExtractData", "SourcePath", textB_SourcePath.Text);
            iniFile.Write("GRR_ExtractData", "StartRow", textB_StartRow.Text);
            iniFile.Write("GRR_ExtractData", "StartCol", textB_StartCol.Text);
            iniFile.Write("GRR_ExtractData", "FromSheet", textB_FromSheet.Text);
            iniFile.Write("GRR_ExtractData", "StartRowDest", textB_StartRowDest.Text);
            iniFile.Write("GRR_ExtractData", "StartColDest", textB_StartColDest.Text);
            iniFile.Write("GRR_ExtractData", "ToSheet", textB_ToSheet.Text);
            iniFile.Write("GRR_ExtractData", "Repeat", textB_Repeat.Text);
            iniFile.Write("GRR_ExtractData", "NumSN", textB_NumSN.Text);
            iniFile.Write("GRR_ExtractData", "TotalItem", textB_TotalItem.Text);

            // PasteToGRR
            iniFile.Write("GRR_PasteToGRR", "StartRow", textB_StartRow_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "StartCol", textB_StartCol_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "FromSheet", textB_FromSheet_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "TargetPath", textB_TargetPath.Text);
            iniFile.Write("GRR_PasteToGRR", "StartRowDest", textB_StartRowDest_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "StartColDest", textB_StartColDest_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "TrialsCount", textB_TrialsCount_GRR.Text);
            iniFile.Write("GRR_PasteToGRR", "NumSN", textB_NumSN_GRR.Text);

            // Limit
            iniFile.Write("GRR_Limit", "StartRow", textB_StartRowLimit.Text);
            iniFile.Write("GRR_Limit", "StartCol", textB_StartColLimit.Text);
            iniFile.Write("GRR_Limit", "FromSheet", textB_FromSheetLimit.Text);
            iniFile.Write("GRR_Limit", "StartRowDest", textB_StartRowDestLimit.Text);
            iniFile.Write("GRR_Limit", "StartColDest", textB_StartColDestLimit.Text);

            // General Excel parameter
            iniFile.Write("General_ExcelParam", "ExcelPath", textB_ExcelPath.Text);
            iniFile.Write("General_ExcelParam", "SheetName", richT_SheetName.Text.Replace("\r\n", "\\n").Replace("\n", "\\n"));
            iniFile.Write("General_ExcelParam", "ReserveCnt", textB_ReserveSheetCount.Text);
            iniFile.Write("General_ExcelParam", "PasteParam", textB_CopyPastePara.Text);
            iniFile.Write("General_ExcelParam", "DeleteParam", textB_DeletePara.Text);

            return true;
            #endregion
        }
        // update UI Control from INI File
        private bool UpdateUIControlFromIniFile()
        {
            #region INI_FILE_READ
            string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);//获取当前执行文件所在路径
            string path = executablePath + "\\Ini\\TestItemStatistics.ini";
            if (!File.Exists(path))
            {
                LogMsg.Message += $"文件：'{path}' 不存在\r\n";
                return false;
            }
            IniFile iniFile = new IniFile(path);

            // 读取数据  
            // Extract Data 文本框控件 初始化
            textB_SourcePath.Text = iniFile.Read("GRR_ExtractData", "SourcePath");
            textB_StartRow.Text = iniFile.Read("GRR_ExtractData", "StartRow");
            textB_StartCol.Text = iniFile.Read("GRR_ExtractData", "StartCol");
            textB_FromSheet.Text = iniFile.Read("GRR_ExtractData", "FromSheet");
            textB_StartRowDest.Text = iniFile.Read("GRR_ExtractData", "StartRowDest");
            textB_StartColDest.Text = iniFile.Read("GRR_ExtractData", "StartColDest");
            textB_ToSheet.Text = iniFile.Read("GRR_ExtractData", "ToSheet");
            textB_Repeat.Text = iniFile.Read("GRR_ExtractData", "Repeat");
            textB_NumSN.Text = iniFile.Read("GRR_ExtractData", "NumSN");
            textB_TotalItem.Text = iniFile.Read("GRR_ExtractData", "TotalItem");

            // Paste to GRR 文本框控件 初始化
            textB_StartRow_GRR.Text = iniFile.Read("GRR_PasteToGRR", "StartRow");
            textB_StartCol_GRR.Text = iniFile.Read("GRR_PasteToGRR", "StartCol");
            textB_FromSheet_GRR.Text = iniFile.Read("GRR_PasteToGRR", "FromSheet");
            textB_TargetPath.Text = iniFile.Read("GRR_PasteToGRR", "TargetPath");
            textB_StartRowDest_GRR.Text = iniFile.Read("GRR_PasteToGRR", "StartRowDest");
            textB_StartColDest_GRR.Text = iniFile.Read("GRR_PasteToGRR", "StartColDest");
            textB_TrialsCount_GRR.Text = iniFile.Read("GRR_PasteToGRR", "TrialsCount");
            textB_NumSN_GRR.Text = iniFile.Read("GRR_PasteToGRR", "NumSN");

            // limit 文本框控件 初始化
            textB_StartRowLimit.Text = iniFile.Read("GRR_Limit", "StartRow");
            textB_StartColLimit.Text = iniFile.Read("GRR_Limit", "StartCol");
            textB_FromSheetLimit.Text = iniFile.Read("GRR_Limit", "FromSheet");
            textB_StartRowDestLimit.Text = iniFile.Read("GRR_Limit", "StartRowDest");
            textB_StartColDestLimit.Text = iniFile.Read("GRR_Limit", "StartColDest");

            // Genenal Excel parameter 初始化
            textB_ExcelPath.Text = iniFile.Read("General_ExcelParam", "ExcelPath");
            richT_SheetName.Text = iniFile.Read("General_ExcelParam", "SheetName").Replace(@"\n", Environment.NewLine);
            textB_ReserveSheetCount.Text = iniFile.Read("General_ExcelParam", "ReserveCnt");
            textB_CopyPastePara.Text = iniFile.Read("General_ExcelParam", "PasteParam");
            textB_DeletePara.Text = iniFile.Read("General_ExcelParam", "DeleteParam");

            // Max Log Size
            long logSize;
            MaxLogSize = long.TryParse(iniFile.Read("Setting", "MaxLogSize"),out logSize) ? logSize : MaxLogSize;

            return true;
            #endregion

        }


#if INIFILE
        // read ini file
        private Dictionary<string, string> ReadIni()
        {
            //BaseDirectory 返回的是应用程序的启动目录，通常是项目构建后生成的输出文件夹（如 bin\Debug 或 bin\Release 目录）。如果你在开发阶段使用它，路径可能是编译后的输出目录
            //string projectPath = AppDomain.CurrentDomain.BaseDirectory;//获取应用程序的根目录路径
            string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);//获取当前执行文件所在路径
            string path = executablePath + "\\Ini\\TestItemStatistics.ini";
            var iniData = iniReaderWriter.ReadIniFile(path, Encoding.GetEncoding("GB2312"));// 使用 GB2312 编码读取
            return iniData;
        }

        // update UI Control from INI File
        private void UpdateUIControlFromIniFile_1()
        {
        #region IniReaderWriter
            var iniData = ReadIni();
            if (iniData.Count == 0)
            {
                InitUILog("配置文件读取失败......\r\n");
                return;
            }
            // ExtractData 文本框控件 初始化
            textB_SourcePath.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.SourcePath");
            textB_StartRow.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.StartRow");
            textB_StartCol.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.StartCol");
            textB_FromSheet.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.FromSheet");
            textB_StartRowDest.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.StartRowDest");
            textB_StartColDest.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.StartColDest");
            textB_ToSheet.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.ToSheet");
            textB_Repeat.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.Repeat");
            textB_NumSN.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.NumSN");
            textB_TotalItem.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_ExtractData.TotalItem");

            // Paste to GRR 文本框控件 初始化
            textB_StartRow_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.StartRow");
            textB_StartCol_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.StartCol");
            textB_FromSheet_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.FromSheet");
            textB_TargetPath.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.TargetPath");
            textB_StartRowDest_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.StartRowDest");
            textB_StartColDest_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.StartColDest");
            textB_TrialsCount_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.TrialsCount");
            textB_NumSN_GRR.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_PasteToGRR.NumSN");

            // limit 文本框控件 初始化
            textB_StartRowLimit.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_Limit.StartRow");
            textB_StartColLimit.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_Limit.StartCol");
            textB_FromSheetLimit.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_Limit.FromSheet");
            textB_StartRowDestLimit.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_Limit.StartRowDest");
            textB_StartColDestLimit.Text = iniReaderWriter.GetValue<string>(iniData, "GRR_Limit.StartColDest");

            // Genenal Excel parameter 初始化
            textB_ExcelPath.Text = iniReaderWriter.GetValue<string>(iniData, "General_ExcelParam.ExcelPath");
            richT_SheetName.Text = iniReaderWriter.GetValue<string>(iniData, "General_ExcelParam.SheetName").Replace(@"\n", Environment.NewLine);
            textB_ReserveSheetCount.Text = iniReaderWriter.GetValue<string>(iniData, "General_ExcelParam.ReserveCnt");
            textB_CopyPastePara.Text = iniReaderWriter.GetValue<string>(iniData, "General_ExcelParam.PasteParam");
            textB_DeletePara.Text = iniReaderWriter.GetValue<string>(iniData, "General_ExcelParam.DeleteParam");
        #endregion
            //InitUILog("初始化完成......\r\n");
            string ss = richT_SheetName.Text.Replace(Environment.NewLine, "\n");
        }
        // update INI file from UI Control
        private void UpdateIniFileFromUIControl_1()
        {
        #region IniReaderWriter
            string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);//获取当前执行文件所在路径
            string path = executablePath + "\\Ini\\TestItemStatistics.ini";

            // 创建一个字典来存储键值对
            Dictionary<string, object> iniData = new Dictionary<string, object>();

            // ExtractData
            iniData.Add("GRR_ExtractData.SourcePath", textB_SourcePath.Text);
            iniData.Add("GRR_ExtractData.StartRow", textB_StartRow.Text);
            iniData.Add("GRR_ExtractData.StartCol", textB_StartCol.Text);
            iniData.Add("GRR_ExtractData.FromSheet", textB_FromSheet.Text);
            iniData.Add("GRR_ExtractData.StartRowDest", textB_StartRowDest.Text);
            iniData.Add("GRR_ExtractData.StartColDest", textB_StartColDest.Text);
            iniData.Add("GRR_ExtractData.ToSheet", textB_ToSheet.Text);
            iniData.Add("GRR_ExtractData.Repeat", textB_Repeat.Text);
            iniData.Add("GRR_ExtractData.NumSN", textB_NumSN.Text);
            iniData.Add("GRR_ExtractData.TotalItem", textB_TotalItem.Text);

            // PasteToGRR
            iniData.Add("GRR_PasteToGRR.StartRow", textB_StartRow_GRR.Text);
            iniData.Add("GRR_PasteToGRR.StartCol", textB_StartCol_GRR.Text);
            iniData.Add("GRR_PasteToGRR.FromSheet", textB_FromSheet_GRR.Text);
            iniData.Add("GRR_PasteToGRR.TargetPath", textB_TargetPath.Text);
            iniData.Add("GRR_PasteToGRR.StartRowDest", textB_StartRowDest_GRR.Text);
            iniData.Add("GRR_PasteToGRR.StartColDest", textB_StartColDest_GRR.Text);
            iniData.Add("GRR_PasteToGRR.TrialsCount", textB_TrialsCount_GRR.Text);
            iniData.Add("GRR_PasteToGRR.NumSN", textB_NumSN_GRR.Text);

            // Limit
            iniData.Add("GRR_Limit.StartRow", textB_StartRowLimit.Text);
            iniData.Add("GRR_Limit.StartCol", textB_StartColLimit.Text);
            iniData.Add("GRR_Limit.FromSheet", textB_FromSheetLimit.Text);
            iniData.Add("GRR_Limit.StartRowDest", textB_StartRowDestLimit.Text);
            iniData.Add("GRR_Limit.StartColDest", textB_StartColDestLimit.Text);

            // General Excel parameter
            iniData.Add("General_ExcelParam.ExcelPath", textB_ExcelPath.Text);
            iniData.Add("General_ExcelParam.SheetName", richT_SheetName.Text.Replace("\r\n", "\\n").Replace("\n", "\\n"));
            iniData.Add("General_ExcelParam.ReserveCnt", textB_ReserveSheetCount.Text);
            iniData.Add("General_ExcelParam.PasteParam", textB_CopyPastePara.Text);
            iniData.Add("General_ExcelParam.DeleteParam", textB_DeletePara.Text);

            //var iniDataOld = ReadIni();
            //if (iniDataOld.Count == 0)
            //{
            //    InitUILog("配置文件读取失败......\r\n");
            //    return;
            //}
            //Dictionary<string, object> combinedDict = iniDataOld
            //.Concat(iniData)
            //.GroupBy(kvp => kvp.Key)
            //.ToDictionary(g => g.Key, g => g.Last()); // 选择最后一个值，如果有重复键  

            // 写入新的数据到 INI 文件
            //iniReaderWriter.WriteIniFile(path, iniData, Encoding.GetEncoding("GB2312"));
            iniReaderWriter.WriteIniFile(path, iniData, Encoding.GetEncoding("GB2312"), true);
            //Console.WriteLine("\nUI Control data written to INI file!");
        #endregion
        }
#endif

        #endregion


    }
}
