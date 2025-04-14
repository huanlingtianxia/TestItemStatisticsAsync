//#define INIFILE
using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestItemStatisticsAcync.ExcelOperation;
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
            UpdateParamFromUIControl();
            RichUILogAdjustSize(251, -1);
            // 取消选中状态并将光标移到文本框末尾
            textB_TargetPath.SelectionStart = textB_TargetPath.Text.Length;
            textB_TargetPath.SelectionLength = 0;
            LogMsg.Message += "Initialization completed...\r\n";
            UpdateUILog(LogMsg.Message);

        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Logger.Info("程序结束<<<<<<<<<<");
        }
        #endregion

        #region private member        
        private bool _richLogHightStretch = true;
        private static Logger? _logger; // 定义 logger 属性
        private static SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1); // 只允许一个异步操作
        private Stopwatch _stopwatch = new Stopwatch();
        #endregion

        #region internal property
        //IIniReaderWriter iniReaderWriter { get; set; } = new IniReaderWriter();
        //LoggerHander loggerHander { get; set; } = new LoggerHander("log.txt");
        internal ExcelOperater ExcelOp { get; set; } = new ExcelOperater();//Excel 操作
        internal ParametersTestItem TestItem { get; set; } = new ParametersTestItem();//从测试log提取数据参数
        internal ParametersTestItem TestItemGRR { get; set; } = new ParametersTestItem();
        internal ParametersTestItem TestItemGRRLimit { get; set; } = new ParametersTestItem();
        internal ParametersTestItem TestItemGRRSummary { get; set; } = new ParametersTestItem();
        internal LogMessage LogMsg { get; set; } = new LogMessage();// log message
        internal long MaxLogSize { get; set; } = 10 * 1024 * 1024; // 10 MB
        internal static Logger Logger => _logger ?? (_logger = LogManager.GetCurrentClassLogger());// 检查 logger 是否为 null，如果是则通过 LogManager 创建新的实例

        #endregion

        #region Path and Ini config button Control Click event
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
        
        // update ini file
        private async void btn_ReadIni_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            try
            {
                InitUILog("waiting......\r\n");
                bool state = UpdateUIControlFromIniFile();
                LogMsg.Message += state ? "read ini file and update to UI Control completed\r\n" : "read ini file failed\r\n";
                UpdateUILog(LogMsg.Message);
            }
            catch (Exception ex) { UpdateUILog(LogMsg.Message += ex.ToString()); }
            finally { _semaphore.Release(); }// 释放信号量，允许下一个操作执行
            StopWatchTime();
        }
        private async void btn_WriteIni_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            InitUILog("waiting......\r\n");
            try
            {
                DialogResult result = MessageBox.Show(this, "确定将UI中的数据更新到 Ini 文件?", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
                StartWatchTime();
                if (result == DialogResult.OK)
                {
                    bool state = UpdateIniFileFromUIControl();
                    LogMsg.Message += state ? "已将UI数据更新到Ini文件\r\n" : "UI数据更新到Ini文件 异常\r\n";
                    _stopwatch.Stop();
                    MessageBox.Show(LogMsg.Message);// 用户点击“确认”
                }
                else
                {

                    _stopwatch.Stop();
                    MessageBox.Show(this, LogMsg.Message += "操作已取消\r\n");// 用户点击“取消”
                }
                UpdateUILog(LogMsg.Message);
            }
            catch (Exception ex) { UpdateUILog(LogMsg.Message += ex.ToString()); }
            finally { _semaphore.Release(); }// 释放信号量，允许下一个操作执行
            StopWatchTime();
        }
        #endregion

        #region GRR button Control Click event
        // Extract Data & GRR Paste
        private async void btn_ExtractData_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            bool state = UpdateParamFromUIControl();
            if(state)
                await ExcelOp.ExtractDataFromTestItem(TestItem.SourcePath, TestItem, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_PasteToGRR_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            bool state = UpdateParamFromUIControl();
            if( state)
                await ExcelOp.PasteToGRRModuleFromExtractData(TestItemGRR.SourcePath, TestItemGRR.TargetPath, TestItemGRR, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_PasteLimit_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            bool state = UpdateParamFromUIControl();
            if( state )
                await ExcelOp.PasteToGRRModuleFromLimit(TestItemGRRLimit.SourcePath, TestItemGRRLimit.TargetPath, TestItemGRRLimit, LogMsg).ConfigureAwait(false);
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_ExtractSheetToTxt_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            bool state = UpdateParamFromUIControl();
            try
            {
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
            catch (Exception ex)
            {
                LogMsg.Message += $"异常: {ex.ToString()}\r\n";
                UpdateUILog(LogMsg.Message);
            }
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void Btn_SummaryFormula_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            bool state = UpdateParamFromUIControl();
            try
            {
                string[] paramsArr = textB_SummaryParam.Text.Split(new[] { "\\n" }, StringSplitOptions.RemoveEmptyEntries);
                int[][] paramsArray = GetParam(paramsArr.ElementAtOrDefault(0) + "\\n" + paramsArr.ElementAtOrDefault(1));
                string targetSheet = paramsArr.ElementAtOrDefault(2);
                if (state)
                    await ExcelOp.PasteToGRRModuleForSummaryFormula(TestItemGRR.TargetPath, TestItemGRR, paramsArray, LogMsg, targetSheet).ConfigureAwait(false);
                UpdateUILog(LogMsg.Message);
            }
            catch (Exception ex)
            {
                LogMsg.Message += $"异常: {ex.ToString()}\r\n";
                UpdateUILog(LogMsg.Message);
            }
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        #endregion

        #region General button Control Click event
        //General: CopyPaste And Delete
        private async void btn_CopyPaste_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            ParametersTestItem[] param = null;
            bool state = UpdateParamFromUIControl(ref param, GeneralMode.CellCopyPaste);
            if(state)
            {
                for (int i = 0; i < param.Length; i++)
                {
                    await ExcelOp.CopyRangePaste(param[i].TargetPath, param[i], LogMsg).ConfigureAwait(false); // 复制 公式单元格
                }
            }
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_DeleteRange_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            DialogResult result = MessageBox.Show(this, "确定删除 DeleteParam 参数中的数据？", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Cancel)
            {
                LogMsg.Message += "取消删除 DeleteParam 操作\r\n";
                UpdateUILog(LogMsg.Message);
                return;
            }
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            //await Task.Delay(5000);
            ParametersTestItem[] param = null;
            bool state = UpdateParamFromUIControl(ref param, GeneralMode.CellDelete);
            if(state)
            {
                for (int i = 0; i < param.Length; i++)
                    await ExcelOp.DeleteRangeData(param[i].TargetPath, param[i], LogMsg).ConfigureAwait(false); // 删除 17行单元格，作用域：11.xx测试项
            }
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_DeleteSheet_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            ParametersTestItem[] param = null;
            bool state = UpdateParamFromUIControl(ref param, GeneralMode.SheetOperater);
            string content = param.ElementAtOrDefault(0).ReserveSheetCount == -1 ? "确定删除SheetName中的所有sheet？" :
                            $"确定删除sheet,仅保留最右侧 {param.ElementAtOrDefault(0).ReserveSheetCount} 个sheet？\r\n注意：最左侧的sheet（Summary）不计算在内";
            DialogResult result = MessageBox.Show(this, content, "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Cancel)
            {
                LogMsg.Message += "取消删除Sheet操作\r\n";
                UpdateUILog(LogMsg.Message);
                return;
            }
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
             if(state)
            {
                if (param.ElementAtOrDefault(0).ReserveSheetCount == -1)
                {
                    await ExcelOp.DeleteSheet(param.ElementAtOrDefault(0).TargetPath, param.ElementAtOrDefault(0), LogMsg).ConfigureAwait(false);//删除SheetName中的工作表
                }
                else
                {
                    await ExcelOp.DeleteSheet(param.ElementAtOrDefault(0).TargetPath, param.ElementAtOrDefault(0).ReserveSheetCount, LogMsg).ConfigureAwait(false);//保留ReserveSheetCount个工作表
                }
            }
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_CreatSheet_Click(object sender, EventArgs e)
        {
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            InitUILog("waiting......\r\n");
            ParametersTestItem[] param = null;
            bool state = UpdateParamFromUIControl(ref param, GeneralMode.SheetOperater);
            if(state)
            {
                await ExcelOp.CreatSheet(param.ElementAtOrDefault(0).TargetPath, param.ElementAtOrDefault(0), LogMsg).ConfigureAwait(false);
            }
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }
        private async void btn_RemaneSheet_Click(object sender, EventArgs e)
        {
            InitUILog("waiting......\r\n");
            DialogResult result = MessageBox.Show(this, "确定将excel中从右->左的所有sheet\r\n按SheetName中从上->下的字符重命名？\r\n注意：sheet数量必须要相等", "确认", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Cancel)
            {
                LogMsg.Message += "取消重命名Sheet操作\r\n";
                UpdateUILog(LogMsg.Message);
                return;
            }
            await _semaphore.WaitAsync();// 请求信号量，如果已经有一个操作在执行，其他的操作会被挂起
            StartWatchTime();
            ParametersTestItem[] param = null;
            bool state = UpdateParamFromUIControl(ref param, GeneralMode.SheetOperater);
            if( state )
            {
                await ExcelOp.RenameSheet(param.ElementAtOrDefault(0).TargetPath, param.ElementAtOrDefault(0), LogMsg).ConfigureAwait(false); // 删除 17行单元格，作用域：11.xx测试项
            }
            UpdateUILog(LogMsg.Message);
            StopWatchTime();
            _semaphore.Release();// 释放信号量，允许下一个操作执行
        }

        // UI log size adjust
        private void lab_EnableMaskArrow_Click(object sender, EventArgs e)
        {
            _richLogHightStretch = !_richLogHightStretch;
            Action action = _richLogHightStretch ? (Action)(() => { RichUILogAdjustSize(251, -1); }) : () => { RichUILogAdjustSize(144, 1); };
            action();
        }
        #endregion

        #region private UI
        /* UI Control update */
        // update richtextbox log,BeginInvoke 是异步执行的，不会阻塞当前线程，而 Invoke 是同步执行的，会等待操作完成。
        private void UpdateUILog(string msg)
        {
            if (richTB_Log.InvokeRequired)
            {
                // 调用 UI 线程来更新 RichTextBox
                richTB_Log.BeginInvoke(new Action(() => {
                    richTB_Log.AppendText(msg);
                    RichUILogScrollToEnd();
                    Logger.Info(msg);
                }));

            }
            else
            {
                richTB_Log.AppendText(msg);
                RichUILogScrollToEnd();
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
        private void RichUILogScrollToEnd()
        {
            // 将光标移动到文本末尾
            richTB_Log.SelectionStart = richTB_Log.Text.Length;
            // 滚动到光标所在的位置
            richTB_Log.ScrollToCaret();
        }
        private void RichUILogAdjustSize(int newHeight, int stretch)
        {
            // 获取当前的尺寸和位置  
            int currentHeight = richTB_Log.Height;
            int dtHeight = Math.Abs(newHeight - currentHeight) * stretch; // 向上拉伸或缩放 dtHeight 像素  
            richTB_Log.Size = new Size(richTB_Log.Width, newHeight);// 更新 RichTextBox 的高度
            Point currentLocation = richTB_Log.Location;
            richTB_Log.Location = new Point(currentLocation.X, currentLocation.Y + dtHeight); // 向上或向下移动 dtHeight 像素
            //this.ClientSize = new Size(this.ClientSize.Width, this.ClientSize.Height + dtHeight);// 可选：更新窗口尺寸以适应新大小  
        }

        //Get GRR parameters
        private bool UpdateParamFromUIControl()
        {
            try
            {
                BindFromUIToModel(TestItem, new Dictionary<string, string>
                {
                    { "StartRow", "textB_StartRow" },
                    { "StartCol", "textB_StartCol" },
                    { "StartRowDest", "textB_StartRowDest" },
                    { "StartColDest", "textB_StartColDest" },
                    { "Repeat", "textB_Repeat" },
                    { "NumSN", "textB_NumSN" },
                    { "TotalItemCount", "textB_TotalItem" },
                    { "FromSheet", "textB_FromSheet" },
                    { "ToSheet", "textB_ToSheet" },
                    { "SourcePath", "textB_SourcePath" }
                });

                BindFromUIToModel(TestItemGRR, new Dictionary<string, string>
                {
                    { "StartRow", "textB_StartRow_GRR" },
                    { "StartCol", "textB_StartCol_GRR" },
                    { "StartRowDest", "textB_StartRowDest_GRR" },
                    { "StartColDest", "textB_StartColDest_GRR" },
                    { "Repeat", "textB_TrialsCount_GRR" },
                    { "NumSN", "textB_NumSN_GRR" },
                    { "FromSheet", "textB_FromSheet_GRR" },
                    { "SourcePath", "textB_SourcePath" },
                    { "TargetPath", "textB_TargetPath" }
                });

                BindFromUIToModel(TestItemGRRLimit, new Dictionary<string, string>
                {
                    { "StartRow", "textB_StartRowLimit" },
                    { "StartCol", "textB_StartColLimit" },
                    { "StartRowDest", "textB_StartRowDestLimit" },
                    { "StartColDest", "textB_StartColDestLimit" },
                    { "FromSheet", "textB_FromSheetLimit" },
                    { "SourcePath", "textB_SourcePath" },
                    { "TargetPath", "textB_TargetPath" }
                });
                //TestItemGRRSummary
                return true;
            }
            catch (Exception ex)
            {
                LogMsg.Message += "UI非法数据输入：" + ex + "\r\n";
                UpdateUILog(LogMsg.Message);
                return false;
            }
        }
        // Get General parameters
        private bool UpdateParamFromUIControl(ref ParametersTestItem[] testItems, GeneralMode mode)
        {
            try
            {
                // 处理 ReserveSheetCount 和 SheetName
                string[] cnt = textB_ReserveSheetCount.Text.Trim().Split(':');
                if (mode == GeneralMode.SheetOperater) // sheet operater
                {
                    testItems = new ParametersTestItem[]
                    {
                        new ParametersTestItem
                        {
                            TargetPath = textB_ExcelPath.Text,
                            SheetName = richT_SheetName.Text.Trim().Split('\n'),
                            ReserveSheetCount = (cnt.Length == 2 && cnt[0] == "") ? int.Parse(cnt[1]) : -1,
                            PosSheetName = textB_PosSheet.Text.Trim()
                        }
                    };
                }
                else if (mode == GeneralMode.CellCopyPaste || mode == GeneralMode.CellDelete) // cell operater
                {
                    // 获取参数输入数据
                    string rawInput = mode == GeneralMode.CellCopyPaste ? textB_CopyPastePara.Text : textB_DeletePara.Text;
                    int[][] paramsArray = GetParam(rawInput);
                    int expectedLength = (mode == GeneralMode.CellCopyPaste) ? 6 : 4;
                    // 检查数据格式
                    if (paramsArray.Any(arr => arr.Length < expectedLength) || paramsArray == null || paramsArray.Length == 0)
                    {
                        LogMsg.Message += (mode == GeneralMode.CellCopyPaste)
                            ? "CopyPasteParam 输入数据格式异常\r\n"
                            : "DeletePara 输入数据格式异常\r\n";
                        return false;
                    }
                    testItems = paramsArray
                    .Select(arr => new ParametersTestItem
                    {
                        StartRow = arr.ElementAtOrDefault(0),
                        StartCol = arr.ElementAtOrDefault(1),
                        EndRow = arr.ElementAtOrDefault(2),
                        EndtCol = arr.ElementAtOrDefault(3),
                        StartRowDest = (mode == GeneralMode.CellCopyPaste) ? arr.ElementAtOrDefault(4) : 0,
                        StartColDest = (mode == GeneralMode.CellCopyPaste) ? arr.ElementAtOrDefault(5) : 0,

                        TargetPath = textB_ExcelPath.Text,
                        SheetName = richT_SheetName.Text.Trim().Split('\n'),
                        ReserveSheetCount = (cnt.Length == 2 && cnt[0] == "") ? int.Parse(cnt[1]) : -1,
                        PosSheetName = textB_PosSheet.Text.Trim()
                    })
                    .ToArray();
                }
                return true;
            }
            catch (Exception ex)
            {
                LogMsg.Message += "UI非法数据输入：" + ex.ToString() + "\r\n";
                UpdateUILog(LogMsg.Message);
                return false;
            }
        }
        
        // INI write :update INI file from UI Control
        private bool UpdateIniFileFromUIControl()
        {
            try
            {
                string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string path = Path.Combine(executablePath, "Ini", "TestItemStatistics.ini");
                if (!File.Exists(path))
                {
                    LogMsg.Message += $"文件：'{path}' 不存在\r\n";
                    return false;
                }

                IniFile iniFile = new IniFile(path);

                // Define the sections and keys for the INI file
                var sectionsAndKeys = new (string Section, string Key, string Value)[]
                {
                    ("GRR_ExtractData", "SourcePath", textB_SourcePath.Text),
                    ("GRR_ExtractData", "StartRow", textB_StartRow.Text),
                    ("GRR_ExtractData", "StartCol", textB_StartCol.Text),
                    ("GRR_ExtractData", "FromSheet", textB_FromSheet.Text),
                    ("GRR_ExtractData", "StartRowDest", textB_StartRowDest.Text),
                    ("GRR_ExtractData", "StartColDest", textB_StartColDest.Text),
                    ("GRR_ExtractData", "ToSheet", textB_ToSheet.Text),
                    ("GRR_ExtractData", "Repeat", textB_Repeat.Text),
                    ("GRR_ExtractData", "NumSN", textB_NumSN.Text),
                    ("GRR_ExtractData", "TotalItem", textB_TotalItem.Text),

                    ("GRR_PasteToGRR", "StartRow", textB_StartRow_GRR.Text),
                    ("GRR_PasteToGRR", "StartCol", textB_StartCol_GRR.Text),
                    ("GRR_PasteToGRR", "FromSheet", textB_FromSheet_GRR.Text),
                    ("GRR_PasteToGRR", "TargetPath", textB_TargetPath.Text),
                    ("GRR_PasteToGRR", "StartRowDest", textB_StartRowDest_GRR.Text),
                    ("GRR_PasteToGRR", "StartColDest", textB_StartColDest_GRR.Text),
                    ("GRR_PasteToGRR", "TrialsCount", textB_TrialsCount_GRR.Text),
                    ("GRR_PasteToGRR", "NumSN", textB_NumSN_GRR.Text),

                    ("GRR_Limit", "StartRow", textB_StartRowLimit.Text),
                    ("GRR_Limit", "StartCol", textB_StartColLimit.Text),
                    ("GRR_Limit", "FromSheet", textB_FromSheetLimit.Text),
                    ("GRR_Limit", "StartRowDest", textB_StartRowDestLimit.Text),
                    ("GRR_Limit", "StartColDest", textB_StartColDestLimit.Text),

                    ("GRR_Summary", "TargetParam", textB_SummaryParam.Text),

                    ("General_ExcelParam", "ExcelPath", textB_ExcelPath.Text),
                    ("General_ExcelParam", "SheetName", richT_SheetName.Text.Replace("\r\n", "\\n").Replace("\n", "\\n")),
                    ("General_ExcelParam", "ReserveCnt", textB_ReserveSheetCount.Text),
                    ("General_ExcelParam", "PasteParam", textB_CopyPastePara.Text),
                    ("General_ExcelParam", "DeleteParam", textB_DeletePara.Text),
                    ("General_ExcelParam", "PosSheetName", textB_PosSheet.Text)
                };

                // Loop through the sections and write values
                foreach (var item in sectionsAndKeys)
                {
                    iniFile.Write(item.Section, item.Key, item.Value);
                }

                return true;
            }
            catch (Exception ex)
            {
                LogMsg.Message += "UI非法数据输入：" + ex.ToString() + "\r\n";
                UpdateUILog(LogMsg.Message);
                return false;
            }
        }
        // INI read :update UI Control from INI File
        private bool UpdateUIControlFromIniFile()
        {
            try
            {
                string executablePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                string path = Path.Combine(executablePath, "Ini", "TestItemStatistics.ini");
                if (!File.Exists(path))
                {
                    LogMsg.Message += $"文件：'{path}' 不存在\r\n";
                    return false;
                }

                IniFile iniFile = new IniFile(path);

                // Define the sections and keys for the INI file
                var sectionsAndKeys = new (string Section, string Key, Action<string> SetValue)[]
                {
                    ("GRR_ExtractData", "SourcePath", value => textB_SourcePath.Text = value),
                    ("GRR_ExtractData", "StartRow", value => textB_StartRow.Text = value),
                    ("GRR_ExtractData", "StartCol", value => textB_StartCol.Text = value),
                    ("GRR_ExtractData", "FromSheet", value => textB_FromSheet.Text = value),
                    ("GRR_ExtractData", "StartRowDest", value => textB_StartRowDest.Text = value),
                    ("GRR_ExtractData", "StartColDest", value => textB_StartColDest.Text = value),
                    ("GRR_ExtractData", "ToSheet", value => textB_ToSheet.Text = value),
                    ("GRR_ExtractData", "Repeat", value => textB_Repeat.Text = value),
                    ("GRR_ExtractData", "NumSN", value => textB_NumSN.Text = value),
                    ("GRR_ExtractData", "TotalItem", value => textB_TotalItem.Text = value),

                    ("GRR_PasteToGRR", "StartRow", value => textB_StartRow_GRR.Text = value),
                    ("GRR_PasteToGRR", "StartCol", value => textB_StartCol_GRR.Text = value),
                    ("GRR_PasteToGRR", "FromSheet", value => textB_FromSheet_GRR.Text = value),
                    ("GRR_PasteToGRR", "TargetPath", value => textB_TargetPath.Text = value),
                    ("GRR_PasteToGRR", "StartRowDest", value => textB_StartRowDest_GRR.Text = value),
                    ("GRR_PasteToGRR", "StartColDest", value => textB_StartColDest_GRR.Text = value),
                    ("GRR_PasteToGRR", "TrialsCount", value => textB_TrialsCount_GRR.Text = value),
                    ("GRR_PasteToGRR", "NumSN", value => textB_NumSN_GRR.Text = value),

                    ("GRR_Limit", "StartRow", value => textB_StartRowLimit.Text = value),
                    ("GRR_Limit", "StartCol", value => textB_StartColLimit.Text = value),
                    ("GRR_Limit", "FromSheet", value => textB_FromSheetLimit.Text = value),
                    ("GRR_Limit", "StartRowDest", value => textB_StartRowDestLimit.Text = value),
                    ("GRR_Limit", "StartColDest", value => textB_StartColDestLimit.Text = value),

                    ("GRR_Summary", "TargetParam", value => textB_SummaryParam.Text = value),

                    ("General_ExcelParam", "ExcelPath", value => textB_ExcelPath.Text = value),
                    ("General_ExcelParam", "SheetName", value => richT_SheetName.Text = value.Replace(@"\n", Environment.NewLine)),
                    ("General_ExcelParam", "ReserveCnt", value => textB_ReserveSheetCount.Text = value),
                    ("General_ExcelParam", "PasteParam", value => textB_CopyPastePara.Text = value),
                    ("General_ExcelParam", "DeleteParam", value => textB_DeletePara.Text = value),
                    ("General_ExcelParam", "PosSheetName", value => textB_PosSheet.Text = value)
                };

                // Loop through the sections and read values
                foreach (var item in sectionsAndKeys)
                {
                    string value = iniFile.Read(item.Section, item.Key);
                    item.SetValue(value);
                }

                // Max Log Size
                long logSize;
                MaxLogSize = long.TryParse(iniFile.Read("Setting", "MaxLogSize"), out logSize) ? logSize : MaxLogSize;

                return true;
            }
            catch (Exception ex)
            {
                LogMsg.Message += "UI非法数据输入：" + ex.ToString() + "\r\n";
                UpdateUILog(LogMsg.Message);
                return false;
            }
        }
        #endregion

        #region private base func
        /* path */
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
            if (mode)
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

        /* Stopwatch */
        private void StartWatchTime()
        {
            _stopwatch.Reset();
            _stopwatch.Start();
        }
        private void StopWatchTime()
        {
            _stopwatch.Stop();
            UpdateUILog($"耗时: {_stopwatch.ElapsedMilliseconds} ms\r\n");
        }

        /* config log */
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

        // private parameters
        private int[][] GetParam(string param)
        {
            int[][] result = param
                            .Split(new[] { "\\n" }, StringSplitOptions.RemoveEmptyEntries)
                            .Select(line => line
                                .Split(',')
                                .Select(s => int.Parse(s.Trim()))
                                .ToArray())
                            .ToArray();
            return result;
        }

        #region general reflect Get Control value to object
        // reflect
        private void BindFromUIToModel(object model, Dictionary<string, string> propertyToControlMap)
        {
            Type modelType = model.GetType();

            foreach (var kvp in propertyToControlMap)
            {
                string propertyName = kvp.Key;
                string controlName = kvp.Value;

                PropertyInfo prop = modelType.GetProperty(propertyName);
                if (prop == null || !prop.CanWrite) continue;

                Control[] found = this.Controls.Find(controlName, true);
                if (found.Length == 0 || !(found[0] is TextBox textBox))
                    throw new Exception($"控件未找到或不是 TextBox：{controlName}");

                object value = ConvertValue(prop.PropertyType, textBox.Text, controlName);
                prop.SetValue(model, value);
            }
        }
        private object ConvertValue(Type targetType, string input, string controlName)
        {
            try
            {
                if (targetType == typeof(int))
                    return int.Parse(input);
                if (targetType == typeof(string))
                    return input;
                // 可以扩展更多类型
                throw new NotSupportedException($"不支持的目标类型：{targetType.Name}");
            }
            catch
            {
                throw new FormatException($"控件 {controlName} 输入格式错误，无法转换为 {targetType.Name}");
            }
        }
        #endregion

        #endregion


    }
}
