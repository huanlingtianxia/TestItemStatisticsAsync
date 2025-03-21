using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TestItem.Excel;

namespace TestItemStatistics
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            InitPara();
        }
        private void InitPara()
        {
            textB_SourcePath.Text = @"E:\labview\MSA\D474 D475\ProdDataMSA\ExtractData\done\";
            textB_TargetPath.Text = @"E:\labview\MSA\D474 D475\ProdDataMSA\ExtractData\done\";
            // 取消选中状态并将光标移到文本框末尾
            textB_TargetPath.SelectionStart = textB_TargetPath.Text.Length;
            textB_TargetPath.SelectionLength = 0;
            UpdateParaFromControl();
        }

        ExcelOperation excelOperation { get; set; } = new ExcelOperation();//Excel 操作
        ParametersTestItem testItem { get; set; } = new ParametersTestItem();//从测试log提取数据参数
        ParametersTestItem testItemGRR { get; set; } = new ParametersTestItem();// copy paste 提取数据到GRR module 参数
        ParametersTestItem testItemGRRLimit { get; set; } = new ParametersTestItem();// copy paste Limit到GRR module 参数
        string msg = string.Empty;

        #region Control Click event
        //Extract data, copy paste test data to GRR, copy paste limit to GRR
        private void btn_SelectSourcePath_Click(object sender, EventArgs e)
        {
            //string path1 = SelectfullPath();
            string path = SelectfullPath();
            if (path != String.Empty)
                textB_SourcePath.Text = path;
        }

        private void btn_SelectTargetPath_Click(object sender, EventArgs e)
        {
            //string path = SelectPath();
            string path = SelectfullPath();
            if (path != String.Empty)
                textB_TargetPath.Text = path;
        }
       
        private void btn_ExtractData_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            UpdateParaFromControl();
            excelOperation.ExtractDataFromTestItem(testItem.SourcePath, testItem, ref msg);
            richTB_Log.Text += msg;
        }

        private void btn_PasteToGRR_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            UpdateParaFromControl();
            excelOperation.PasteToGRRModuleFromExtractData(testItemGRR.SourcePath, testItemGRR.TargetPath, testItemGRR, ref msg);
            richTB_Log.Text += msg;
        }

        private void btn_ExtractSheetToTxt_Click(object sender, EventArgs e)
        {
            try
            {
                UpdateParaFromControl();
                string[] str = { "\\" };
                string path = string.Empty;
                string[] pathArr = testItemGRR.TargetPath.Split(str, StringSplitOptions.None);
                for (int i = 0; i < pathArr.Length - 1; i++)
                {
                    path += pathArr[i] + "\\";
                }
                path += "GRRModuleSheetName.txt";
                string[] sheetName = excelOperation.GetSheetName(testItemGRR.TargetPath, false, path);
                msg += "提取GRR module中 test item sheet name 到GRRModuleSheetName.txt,\r\n path: " + path + "\r\n";
                if(sheetName != null)
                {
                    for (int i = 0; i < sheetName.Length; i++)
                    {
                        msg += $"序号：{i + 1,-6} {sheetName[i]}\r\n";
                    }
                    msg += $"提取sheet name 完成！sheet count: {sheetName.Length} ----------------------\r\n";
                }
                else
                {
                    msg += $"未找到工作表\r\n";
                }
                richTB_Log.Text += msg;
            }
            catch(Exception ex)
            {
                msg += $"异常: {ex.ToString()}\r\n";
                richTB_Log.Text += msg;
            }
            
        }

        private void btn_PasteLimit_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            UpdateParaFromControl();
            excelOperation.PasteToGRRModuleFromLimit(testItemGRRLimit.SourcePath, testItemGRRLimit.TargetPath, testItemGRRLimit, ref msg);
            richTB_Log.Text += msg;
        }

        //General: CopyPaste And Delete
        private void btn_CopyPaste_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            ParametersTestItem para = new ParametersTestItem();
            UpdateParaFromControl(para, true);
            excelOperation.CopyRangePaste(para.TargetPath, para); // 复制 公式单元格
            msg += $"拷贝粘贴完成！";
            richTB_Log.Text = msg;
        }

        private void btn_DeleteRange_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            ParametersTestItem para = new ParametersTestItem();
            UpdateParaFromControl(para, false);
            excelOperation.DeleteRangeData(para.TargetPath, para); // 删除 17行单元格，作用域：11.xx测试项
            msg += $"删除：开始行{para.StartRow}，开始列{para.StartCol}，结束行{para.EndRow}，结束始列{para.EndtCol} 完成！";
            richTB_Log.Text = msg;
        }

        private void btn_DeleteSheet_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            ParametersTestItem parametersTestItem = new ParametersTestItem();
            UpdateParaFromControl(parametersTestItem, false);
            if (parametersTestItem.ReserveSheetCount == -1)
            {
                excelOperation.DeleteSheet(parametersTestItem.TargetPath, parametersTestItem, ref msg);//删除SheetName中的工作表
            }
            else
            {
                excelOperation.DeleteSheet(parametersTestItem.TargetPath, parametersTestItem.ReserveSheetCount, ref msg);//保留ReserveSheetCount个工作表
            }
            //string[] sheet = excelOperation.GetSheetName(parametersTestItem.TargetPath,true);
            richTB_Log.Text = msg;
        }

        private void btn_CreatSheet_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            ParametersTestItem para = new ParametersTestItem();
            UpdateParaFromControl(para, false);
            excelOperation.CreatSheet(para.TargetPath, para, ref msg); // 删除 17行单元格，作用域：11.xx测试项
            richTB_Log.Text = msg;
        }
        #endregion

        #region function
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
        private void UpdateParaFromControl()
        {
            try
            {
                //Extract data
                testItem.StartRow = int.Parse(textB_StartRow.Text);
                testItem.StartCol = int.Parse(textB_StartCol.Text);
                testItem.StartRowDest = int.Parse(textB_StartRowDest.Text);
                testItem.StartColDest = int.Parse(textB_StartColDest.Text);
                testItem.Span = int.Parse(textB_Repeat.Text);
                testItem.NumSN = int.Parse(textB_NumSN.Text);
                testItem.TotalItemCount = int.Parse(textB_TotalItem.Text);
                testItem.FromSheet = textB_FromSheet.Text;
                testItem.ToSheet = textB_ToSheet.Text;
                testItem.SourcePath = textB_SourcePath.Text;

                //Paste to GRR module test item
                testItemGRR.StartRow = int.Parse(textB_StartRow_GRR.Text);
                testItemGRR.StartCol = int.Parse(textB_StartCol_GRR.Text);
                testItemGRR.StartRowDest = int.Parse(textB_StartRowDest_GRR.Text);
                testItemGRR.StartColDest = int.Parse(textB_StartColDest_GRR.Text);
                testItemGRR.Span = int.Parse(textB_TrialsCount_GRR.Text);
                testItemGRR.NumSN = int.Parse(textB_NumSN_GRR.Text);
                //testItemGRR.TotalItemCount = int.Parse(textB_TotalItem_GRR.Text);
                testItemGRR.FromSheet = textB_FromSheet_GRR.Text;
                testItemGRR.SourcePath = textB_SourcePath.Text;
                testItemGRR.TargetPath = textB_TargetPath.Text;

                //Paste to GRR module Limit
                testItemGRRLimit.StartRow = int.Parse(textB_StartRowLimit.Text);
                testItemGRRLimit.StartCol = int.Parse(textB_StartColLimit.Text);
                testItemGRRLimit.StartRowDest = int.Parse(textB_StartRowDestLimit.Text);
                testItemGRRLimit.StartColDest = int.Parse(textB_StartColDestLimit.Text);
                testItemGRRLimit.FromSheet = textB_FromSheetLimit.Text;
                testItemGRRLimit.SourcePath = textB_SourcePath.Text;
                testItemGRRLimit.TargetPath = textB_TargetPath.Text;

                // option

                // common
            }
            catch (Exception ex)
            {
                msg += "输入控件不是数字：" + ex.ToString() + "\r\n";
                richTB_Log.Text += msg;
            }
            
        }
        private void UpdateParaFromControl(ParametersTestItem parameters, bool copyPast)
        {
            try
            {
                parameters.TargetPath = textB_ExcelPath.Text;
                parameters.SheetName = richT_SheetName.Text.Trim().Split(new string[1] { "\n" }, StringSplitOptions.None);
                string[] cnt = textB_ReserveSheetCount.Text.Trim().Split(new string[1] { ":" }, StringSplitOptions.None);
                parameters.ReserveSheetCount = cnt.Length == 2 ? int.Parse(cnt[1]) : -1;
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
                msg += "输入控件不是数字：" + ex.ToString() + "\r\n";
                richTB_Log.Text += msg;
            }
        }
        #endregion

    }
}
