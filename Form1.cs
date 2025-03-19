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
        }

        ExcelOperation excelOperation = new ExcelOperation();
        ParametersTestItem testItem = new ParametersTestItem();
        ParametersTestItem testItemGRR = new ParametersTestItem();
        string msg = string.Empty;

        //string sourceExcelFilePaht = @"E:\labview\other prj\IGBT cplusplus dll\MSA1\test1\op4_Test.xlsx";
        //string targetExccelFilePaht = @"E:\labview\other prj\IGBT cplusplus dll\MSA1\test1\GRR_20250317_D471_FCT1_No.1&2&3_Test.xlsx";

        private void btn_SelectSourcePath_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
            if (path != String.Empty)
                textB_SourcePath.Text = path;
        }

        private void btn_SelectTargetPath_Click(object sender, EventArgs e)
        {
            string path = SelectPath();
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

                //Paste to GRR module
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
