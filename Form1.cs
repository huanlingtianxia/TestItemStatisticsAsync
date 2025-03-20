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
            textB_SourcePath.Text = @"E:\labview\MSA\AllLoginOneSheet_C001D471\testCreat\op4_B001_C001.xlsx";
            textB_TargetPath.Text = @"E:\labview\MSA\AllLoginOneSheet_C001D471\testCreat\GRR_20250317_D471_FCT1_No.1&2&3_1.xlsx";
            // 取消选中状态并将光标移到文本框末尾
            textB_TargetPath.SelectionStart = textB_TargetPath.Text.Length;
            textB_TargetPath.SelectionLength = 0;
>>>>>>>>> Temporary merge branch 2
            UpdateParaFromControl();
        }

        ExcelOperation excelOperation = new ExcelOperation();//从测试log提取数据参数
        ParametersTestItem testItem = new ParametersTestItem();//从测试log提取数据参数
        ParametersTestItem testItemGRR = new ParametersTestItem();// copy paste 提取数据到GRR module 参数
        ParametersTestItem testItemGRRLimit = new ParametersTestItem();// copy paste Limit到GRR module 参数
        ParametersTestItem tempTestItem = new ParametersTestItem();//临时 参数
        string msg = string.Empty;

        //string sourceExcelFilePaht = @"E:\labview\other prj\IGBT cplusplus dll\MSA1\test1\op4_Test.xlsx";
        //string targetExccelFilePaht = @"E:\labview\other prj\IGBT cplusplus dll\MSA1\test1\GRR_20250317_D471_FCT1_No.1&2&3_Test.xlsx";

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

        private void btn_ExtracSheetToTxt_Click(object sender, EventArgs e)
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
            string[] sheetName = excelOperation.GetSheetName(testItemGRR.TargetPath, path);
            msg += "提取GRR module中 test item sheet name 到GRRModuleSheetName.txt,\r\n path: " + path + "\r\n";
            for (int i = 0; i < sheetName.Length; i++)
            {
                msg += sheetName[i] + "\r\n";
            }

            msg += $"提取sheet name 完成！sheet count: {sheetName.Length} ----------------------\r\n";
            richTB_Log.Text += msg;
        }

        private void btn_PasteLimit_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            UpdateParaFromControl();
            excelOperation.PasteToGRRModuleFromLimit(testItemGRRLimit.SourcePath, testItemGRRLimit.TargetPath, testItemGRRLimit, ref msg);
            richTB_Log.Text += msg;
        }

        private void btn_CopyGRRModuleAndDelete_Click(object sender, EventArgs e)
        {
            richTB_Log.Clear();
            msg = string.Empty;
            tempTestItem.StartRow = 16;
            tempTestItem.StartCol = 6;
            tempTestItem.EndRow = 16;
            tempTestItem.EndtCol = 6;
            tempTestItem.StartRowDest = 17;
            tempTestItem.StartColDest = 6;
            UpdateParaFromControl();
            excelOperation.PasteToGRRModuleFromFormula(tempTestItem.TargetPath, tempTestItem); // 复制 公式单元格
            msg += "复制粘贴公式到F17，J17，N17公式完成\r\n";
            richTB_Log.Text += msg;

            tempTestItem.StartRow = 17;
            tempTestItem.StartCol = 3;
            tempTestItem.EndRow = 17;
            tempTestItem.EndtCol = 14;
            excelOperation.DeleteRangeDataFromGRRModule(tempTestItem.TargetPath, tempTestItem, 175, 18); // 删除 17行单元格，作用域：11.xx测试项
            msg += "删除11.x item C17:N17完成\r\n";
            richTB_Log.Text += msg;
        }
        //function
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

                // option tempTestItem
                tempTestItem.SourcePath = textB_SourcePath.Text;
                tempTestItem.TargetPath = textB_TargetPath.Text;

                // common
            }
            catch (Exception ex)
            {
                msg += "输入控件不是数字：" + ex.ToString() + "\r\n";
                richTB_Log.Text += msg;
            }
            
        }


    }
}
