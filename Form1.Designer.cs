﻿namespace TestItemStatistics
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.textB_TargetPath = new System.Windows.Forms.TextBox();
            this.textB_SourcePath = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.richTB_Log = new System.Windows.Forms.RichTextBox();
            this.btn_SelectTargetPath = new System.Windows.Forms.Button();
            this.btn_SelectSourcePath = new System.Windows.Forms.Button();
            this.textB_NumSN = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.textB_StartRow = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.textB_StartCol = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.textB_StartRowDest = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.textB_StartColDest = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.textB_TotalItem = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.textB_FromSheet = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.textB_ToSheet = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textB_Repeat = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.btn_ExtractData = new System.Windows.Forms.Button();
            this.btn_PasteToGRR = new System.Windows.Forms.Button();
            this.label22 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.textB_NumSN_GRR = new System.Windows.Forms.TextBox();
            this.label19 = new System.Windows.Forms.Label();
            this.textB_StartRow_GRR = new System.Windows.Forms.TextBox();
            this.label18 = new System.Windows.Forms.Label();
            this.textB_StartCol_GRR = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.textB_TrialsCount_GRR = new System.Windows.Forms.TextBox();
            this.label17 = new System.Windows.Forms.Label();
            this.textB_StartRowDest_GRR = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.textB_StartColDest_GRR = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.textB_FromSheet_GRR = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_ExtracSheetToTxt = new System.Windows.Forms.Button();
            this.textB_StartColLimit = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.textB_StartRowLimit = new System.Windows.Forms.TextBox();
            this.label21 = new System.Windows.Forms.Label();
            this.textB_StartColDestLimit = new System.Windows.Forms.TextBox();
            this.textB_StartRowDestLimit = new System.Windows.Forms.TextBox();
            this.label25 = new System.Windows.Forms.Label();
            this.label26 = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.textB_FromSheetLimit = new System.Windows.Forms.TextBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label28 = new System.Windows.Forms.Label();
            this.btn_PasteLimit = new System.Windows.Forms.Button();
            this.label32 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btn_CopyGRRModuleAndDelete = new System.Windows.Forms.Button();
            this.label29 = new System.Windows.Forms.Label();
            this.label30 = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.label33 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "TargetPath";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "SourcePath";
            // 
            // textB_TargetPath
            // 
            this.textB_TargetPath.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.textB_TargetPath.Location = new System.Drawing.Point(80, 32);
            this.textB_TargetPath.Name = "textB_TargetPath";
            this.textB_TargetPath.Size = new System.Drawing.Size(564, 20);
            this.textB_TargetPath.TabIndex = 6;
            this.textB_TargetPath.Text = "E:\\labview\\MSA\\AllLoginOneSheet_C001D471\\testCreat\\GRR_20250317_D471_FCT1_No.1&2&" +
    "3.xlsx";
            // 
            // textB_SourcePath
            // 
            this.textB_SourcePath.Location = new System.Drawing.Point(80, 6);
            this.textB_SourcePath.Name = "textB_SourcePath";
            this.textB_SourcePath.Size = new System.Drawing.Size(564, 20);
            this.textB_SourcePath.TabIndex = 7;
            this.textB_SourcePath.Text = "E:\\labview\\MSA\\AllLoginOneSheet_C001D471\\testCreat\\op4_B001_C001.xlsx";
            // 
            // richTB_Log
            // 
            this.richTB_Log.Location = new System.Drawing.Point(0, 291);
            this.richTB_Log.Name = "richTB_Log";
            this.richTB_Log.Size = new System.Drawing.Size(732, 203);
            this.richTB_Log.TabIndex = 14;
            this.richTB_Log.Text = "";
            this.richTB_Log.WordWrap = false;
            // 
            // btn_SelectTargetPath
            // 
            this.btn_SelectTargetPath.Location = new System.Drawing.Point(650, 30);
            this.btn_SelectTargetPath.Name = "btn_SelectTargetPath";
            this.btn_SelectTargetPath.Size = new System.Drawing.Size(82, 23);
            this.btn_SelectTargetPath.TabIndex = 11;
            this.btn_SelectTargetPath.Text = "SelectPathT";
            this.btn_SelectTargetPath.UseVisualStyleBackColor = true;
            this.btn_SelectTargetPath.Click += new System.EventHandler(this.btn_SelectTargetPath_Click);
            // 
            // btn_SelectSourcePath
            // 
            this.btn_SelectSourcePath.Location = new System.Drawing.Point(650, 6);
            this.btn_SelectSourcePath.Name = "btn_SelectSourcePath";
            this.btn_SelectSourcePath.Size = new System.Drawing.Size(82, 23);
            this.btn_SelectSourcePath.TabIndex = 12;
            this.btn_SelectSourcePath.Text = "SelectPathS";
            this.btn_SelectSourcePath.UseVisualStyleBackColor = true;
            this.btn_SelectSourcePath.Click += new System.EventHandler(this.btn_SelectSourcePath_Click);
            // 
            // textB_NumSN
            // 
            this.textB_NumSN.Location = new System.Drawing.Point(174, 22);
            this.textB_NumSN.Name = "textB_NumSN";
            this.textB_NumSN.Size = new System.Drawing.Size(35, 20);
            this.textB_NumSN.TabIndex = 15;
            this.textB_NumSN.Text = "9";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(128, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "NumSN";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(24, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(51, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "StartRow";
            // 
            // textB_StartRow
            // 
            this.textB_StartRow.Location = new System.Drawing.Point(77, 19);
            this.textB_StartRow.Name = "textB_StartRow";
            this.textB_StartRow.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRow.TabIndex = 15;
            this.textB_StartRow.Text = "9";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(31, 44);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(44, 13);
            this.label6.TabIndex = 8;
            this.label6.Text = "StartCol";
            // 
            // textB_StartCol
            // 
            this.textB_StartCol.Location = new System.Drawing.Point(77, 41);
            this.textB_StartCol.Name = "textB_StartCol";
            this.textB_StartCol.Size = new System.Drawing.Size(35, 20);
            this.textB_StartCol.TabIndex = 15;
            this.textB_StartCol.Text = "1";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(2, 66);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(73, 13);
            this.label7.TabIndex = 8;
            this.label7.Text = "StartRowDest";
            // 
            // textB_StartRowDest
            // 
            this.textB_StartRowDest.Location = new System.Drawing.Point(77, 63);
            this.textB_StartRowDest.Name = "textB_StartRowDest";
            this.textB_StartRowDest.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRowDest.TabIndex = 15;
            this.textB_StartRowDest.Text = "1";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(9, 87);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(66, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "StartColDest";
            // 
            // textB_StartColDest
            // 
            this.textB_StartColDest.Location = new System.Drawing.Point(77, 84);
            this.textB_StartColDest.Name = "textB_StartColDest";
            this.textB_StartColDest.Size = new System.Drawing.Size(35, 20);
            this.textB_StartColDest.TabIndex = 15;
            this.textB_StartColDest.Text = "2";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(121, 48);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(51, 13);
            this.label9.TabIndex = 8;
            this.label9.Text = "TotalItem";
            // 
            // textB_TotalItem
            // 
            this.textB_TotalItem.Location = new System.Drawing.Point(174, 45);
            this.textB_TotalItem.Name = "textB_TotalItem";
            this.textB_TotalItem.Size = new System.Drawing.Size(35, 20);
            this.textB_TotalItem.TabIndex = 15;
            this.textB_TotalItem.Text = "229";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(114, 70);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(58, 13);
            this.label10.TabIndex = 8;
            this.label10.Text = "FromSheet";
            // 
            // textB_FromSheet
            // 
            this.textB_FromSheet.Location = new System.Drawing.Point(174, 67);
            this.textB_FromSheet.Name = "textB_FromSheet";
            this.textB_FromSheet.Size = new System.Drawing.Size(90, 20);
            this.textB_FromSheet.TabIndex = 15;
            this.textB_FromSheet.Text = "SortSelectTrans";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(124, 92);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(48, 13);
            this.label11.TabIndex = 8;
            this.label11.Text = "ToSheet";
            // 
            // textB_ToSheet
            // 
            this.textB_ToSheet.Location = new System.Drawing.Point(174, 89);
            this.textB_ToSheet.Name = "textB_ToSheet";
            this.textB_ToSheet.Size = new System.Drawing.Size(90, 20);
            this.textB_ToSheet.TabIndex = 15;
            this.textB_ToSheet.Text = "toSheetAll";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textB_Repeat
            // 
            this.textB_Repeat.Location = new System.Drawing.Point(77, 107);
            this.textB_Repeat.Name = "textB_Repeat";
            this.textB_Repeat.Size = new System.Drawing.Size(35, 20);
            this.textB_Repeat.TabIndex = 15;
            this.textB_Repeat.Text = "9";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(28, 110);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(42, 13);
            this.label13.TabIndex = 8;
            this.label13.Text = "Repeat";
            // 
            // btn_ExtractData
            // 
            this.btn_ExtractData.Location = new System.Drawing.Point(562, 57);
            this.btn_ExtractData.Name = "btn_ExtractData";
            this.btn_ExtractData.Size = new System.Drawing.Size(82, 23);
            this.btn_ExtractData.TabIndex = 12;
            this.btn_ExtractData.Text = "ExtractData";
            this.btn_ExtractData.UseVisualStyleBackColor = true;
            this.btn_ExtractData.Click += new System.EventHandler(this.btn_ExtractData_Click);
            // 
            // btn_PasteToGRR
            // 
            this.btn_PasteToGRR.Location = new System.Drawing.Point(647, 55);
            this.btn_PasteToGRR.Name = "btn_PasteToGRR";
            this.btn_PasteToGRR.Size = new System.Drawing.Size(82, 23);
            this.btn_PasteToGRR.TabIndex = 11;
            this.btn_PasteToGRR.Text = "PasteToGRR";
            this.btn_PasteToGRR.UseVisualStyleBackColor = true;
            this.btn_PasteToGRR.Click += new System.EventHandler(this.btn_PasteToGRR_Click);
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(554, 107);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(175, 91);
            this.label22.TabIndex = 8;
            this.label22.Text = "1. 剔除异常数据，排序转置\r\n2. 选择路径：sourcePath是提取\r\n   和GRR的源文件共同路径\r\n3. 设置参数\r\n4. 提取数据\r\n5. 将提取数" +
    "据粘贴到GRR模板：\r\n   TargetPath";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(129, 26);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(44, 13);
            this.label20.TabIndex = 8;
            this.label20.Text = "NumSN";
            // 
            // textB_NumSN_GRR
            // 
            this.textB_NumSN_GRR.Location = new System.Drawing.Point(175, 23);
            this.textB_NumSN_GRR.Name = "textB_NumSN_GRR";
            this.textB_NumSN_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_NumSN_GRR.TabIndex = 15;
            this.textB_NumSN_GRR.Text = "9";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(25, 23);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(51, 13);
            this.label19.TabIndex = 8;
            this.label19.Text = "StartRow";
            // 
            // textB_StartRow_GRR
            // 
            this.textB_StartRow_GRR.Location = new System.Drawing.Point(78, 20);
            this.textB_StartRow_GRR.Name = "textB_StartRow_GRR";
            this.textB_StartRow_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRow_GRR.TabIndex = 15;
            this.textB_StartRow_GRR.Text = "2";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(32, 45);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(44, 13);
            this.label18.TabIndex = 8;
            this.label18.Text = "StartCol";
            // 
            // textB_StartCol_GRR
            // 
            this.textB_StartCol_GRR.Location = new System.Drawing.Point(78, 42);
            this.textB_StartCol_GRR.Name = "textB_StartCol_GRR";
            this.textB_StartCol_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_StartCol_GRR.TabIndex = 15;
            this.textB_StartCol_GRR.Text = "3";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(115, 49);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(60, 13);
            this.label15.TabIndex = 8;
            this.label15.Text = "TrialsCount";
            // 
            // textB_TrialsCount_GRR
            // 
            this.textB_TrialsCount_GRR.Location = new System.Drawing.Point(175, 46);
            this.textB_TrialsCount_GRR.Name = "textB_TrialsCount_GRR";
            this.textB_TrialsCount_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_TrialsCount_GRR.TabIndex = 15;
            this.textB_TrialsCount_GRR.Text = "3";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(3, 67);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(73, 13);
            this.label17.TabIndex = 8;
            this.label17.Text = "StartRowDest";
            // 
            // textB_StartRowDest_GRR
            // 
            this.textB_StartRowDest_GRR.Location = new System.Drawing.Point(78, 64);
            this.textB_StartRowDest_GRR.Name = "textB_StartRowDest_GRR";
            this.textB_StartRowDest_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRowDest_GRR.TabIndex = 15;
            this.textB_StartRowDest_GRR.Text = "9";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(10, 88);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(66, 13);
            this.label16.TabIndex = 8;
            this.label16.Text = "StartColDest";
            // 
            // textB_StartColDest_GRR
            // 
            this.textB_StartColDest_GRR.Location = new System.Drawing.Point(78, 85);
            this.textB_StartColDest_GRR.Name = "textB_StartColDest_GRR";
            this.textB_StartColDest_GRR.Size = new System.Drawing.Size(35, 20);
            this.textB_StartColDest_GRR.TabIndex = 15;
            this.textB_StartColDest_GRR.Text = "3";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(115, 71);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(58, 13);
            this.label14.TabIndex = 8;
            this.label14.Text = "FromSheet";
            // 
            // textB_FromSheet_GRR
            // 
            this.textB_FromSheet_GRR.Location = new System.Drawing.Point(175, 68);
            this.textB_FromSheet_GRR.Name = "textB_FromSheet_GRR";
            this.textB_FromSheet_GRR.Size = new System.Drawing.Size(77, 20);
            this.textB_FromSheet_GRR.TabIndex = 15;
            this.textB_FromSheet_GRR.Text = "toSheetAll";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.textB_ToSheet);
            this.groupBox1.Controls.Add(this.textB_NumSN);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.textB_StartRow);
            this.groupBox1.Controls.Add(this.textB_FromSheet);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.textB_StartCol);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.textB_TotalItem);
            this.groupBox1.Controls.Add(this.textB_StartRowDest);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.textB_StartColDest);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.textB_Repeat);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(6, 59);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(270, 131);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Extract data: SourcePath";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_ExtracSheetToTxt);
            this.groupBox2.Controls.Add(this.textB_StartRowDest_GRR);
            this.groupBox2.Controls.Add(this.label20);
            this.groupBox2.Controls.Add(this.textB_FromSheet_GRR);
            this.groupBox2.Controls.Add(this.textB_NumSN_GRR);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.label19);
            this.groupBox2.Controls.Add(this.textB_StartColDest_GRR);
            this.groupBox2.Controls.Add(this.textB_StartRow_GRR);
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Controls.Add(this.label18);
            this.groupBox2.Controls.Add(this.textB_StartCol_GRR);
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.textB_TrialsCount_GRR);
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Location = new System.Drawing.Point(282, 67);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(259, 123);
            this.groupBox2.TabIndex = 18;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Paste to GRR module: SourcePath to TargetPath";
            // 
            // btn_ExtracSheetToTxt
            // 
            this.btn_ExtracSheetToTxt.Location = new System.Drawing.Point(119, 94);
            this.btn_ExtracSheetToTxt.Name = "btn_ExtracSheetToTxt";
            this.btn_ExtracSheetToTxt.Size = new System.Drawing.Size(132, 24);
            this.btn_ExtracSheetToTxt.TabIndex = 16;
            this.btn_ExtracSheetToTxt.Text = "ExtracSheetToTxt";
            this.btn_ExtracSheetToTxt.UseVisualStyleBackColor = true;
            this.btn_ExtracSheetToTxt.Click += new System.EventHandler(this.btn_ExtracSheetToTxt_Click);
            // 
            // textB_StartColLimit
            // 
            this.textB_StartColLimit.Location = new System.Drawing.Point(119, 41);
            this.textB_StartColLimit.Name = "textB_StartColLimit";
            this.textB_StartColLimit.Size = new System.Drawing.Size(35, 20);
            this.textB_StartColLimit.TabIndex = 15;
            this.textB_StartColLimit.Text = "9";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(73, 44);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(44, 13);
            this.label12.TabIndex = 8;
            this.label12.Text = "StartCol";
            // 
            // textB_StartRowLimit
            // 
            this.textB_StartRowLimit.Location = new System.Drawing.Point(119, 19);
            this.textB_StartRowLimit.Name = "textB_StartRowLimit";
            this.textB_StartRowLimit.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRowLimit.TabIndex = 15;
            this.textB_StartRowLimit.Text = "2";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(66, 22);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(51, 13);
            this.label21.TabIndex = 8;
            this.label21.Text = "StartRow";
            // 
            // textB_StartColDestLimit
            // 
            this.textB_StartColDestLimit.Location = new System.Drawing.Point(299, 41);
            this.textB_StartColDestLimit.Name = "textB_StartColDestLimit";
            this.textB_StartColDestLimit.Size = new System.Drawing.Size(35, 20);
            this.textB_StartColDestLimit.TabIndex = 15;
            this.textB_StartColDestLimit.Text = "7";
            // 
            // textB_StartRowDestLimit
            // 
            this.textB_StartRowDestLimit.Location = new System.Drawing.Point(299, 17);
            this.textB_StartRowDestLimit.Name = "textB_StartRowDestLimit";
            this.textB_StartRowDestLimit.Size = new System.Drawing.Size(35, 20);
            this.textB_StartRowDestLimit.TabIndex = 15;
            this.textB_StartRowDestLimit.Text = "3";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Location = new System.Drawing.Point(6, 17);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(41, 13);
            this.label25.TabIndex = 8;
            this.label25.Text = "Source";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Location = new System.Drawing.Point(183, 31);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(38, 13);
            this.label26.TabIndex = 8;
            this.label26.Text = "Target";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Location = new System.Drawing.Point(4, 65);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(58, 13);
            this.label27.TabIndex = 8;
            this.label27.Text = "FromSheet";
            // 
            // textB_FromSheetLimit
            // 
            this.textB_FromSheetLimit.Location = new System.Drawing.Point(64, 62);
            this.textB_FromSheetLimit.Name = "textB_FromSheetLimit";
            this.textB_FromSheetLimit.Size = new System.Drawing.Size(90, 20);
            this.textB_FromSheetLimit.TabIndex = 15;
            this.textB_FromSheetLimit.Text = "limit";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label28);
            this.groupBox3.Controls.Add(this.btn_PasteLimit);
            this.groupBox3.Controls.Add(this.textB_StartColLimit);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.textB_StartRowLimit);
            this.groupBox3.Controls.Add(this.label32);
            this.groupBox3.Controls.Add(this.label27);
            this.groupBox3.Controls.Add(this.label21);
            this.groupBox3.Controls.Add(this.label24);
            this.groupBox3.Controls.Add(this.label25);
            this.groupBox3.Controls.Add(this.label26);
            this.groupBox3.Controls.Add(this.textB_FromSheetLimit);
            this.groupBox3.Controls.Add(this.textB_StartColDestLimit);
            this.groupBox3.Controls.Add(this.label23);
            this.groupBox3.Controls.Add(this.textB_StartRowDestLimit);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Location = new System.Drawing.Point(11, 196);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(347, 89);
            this.groupBox3.TabIndex = 19;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Paste to GRR: Limit";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.label28.Location = new System.Drawing.Point(168, 8);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(9, 78);
            this.label28.TabIndex = 20;
            this.label28.Text = "|\r\n|\r\n|\r\n|\r\n|\r\n|\r\n";
            // 
            // btn_PasteLimit
            // 
            this.btn_PasteLimit.Location = new System.Drawing.Point(256, 62);
            this.btn_PasteLimit.Name = "btn_PasteLimit";
            this.btn_PasteLimit.Size = new System.Drawing.Size(78, 22);
            this.btn_PasteLimit.TabIndex = 20;
            this.btn_PasteLimit.Text = "PasteLimit";
            this.btn_PasteLimit.UseVisualStyleBackColor = true;
            this.btn_PasteLimit.Click += new System.EventHandler(this.btn_PasteLimit_Click);
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.ForeColor = System.Drawing.Color.Red;
            this.label32.Location = new System.Drawing.Point(193, 66);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(57, 13);
            this.label32.TabIndex = 8;
            this.label32.Text = "3.粘贴limit";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Location = new System.Drawing.Point(6, 36);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(50, 13);
            this.label24.TabIndex = 8;
            this.label24.Text = "LimitHigh";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(220, 18);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(73, 13);
            this.label23.TabIndex = 8;
            this.label23.Text = "StartRowDest";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(227, 41);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "StartColDest";
            // 
            // btn_CopyGRRModuleAndDelete
            // 
            this.btn_CopyGRRModuleAndDelete.Location = new System.Drawing.Point(577, 262);
            this.btn_CopyGRRModuleAndDelete.Name = "btn_CopyGRRModuleAndDelete";
            this.btn_CopyGRRModuleAndDelete.Size = new System.Drawing.Size(152, 23);
            this.btn_CopyGRRModuleAndDelete.TabIndex = 20;
            this.btn_CopyGRRModuleAndDelete.Text = "CopyGRRModuleAndDelete";
            this.btn_CopyGRRModuleAndDelete.UseVisualStyleBackColor = true;
            this.btn_CopyGRRModuleAndDelete.Click += new System.EventHandler(this.btn_CopyGRRModuleAndDelete_Click);
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Location = new System.Drawing.Point(601, 246);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(111, 13);
            this.label29.TabIndex = 8;
            this.label29.Text = "特殊处理11.x测试项";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.ForeColor = System.Drawing.Color.Red;
            this.label30.Location = new System.Drawing.Point(559, 81);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(64, 13);
            this.label30.TabIndex = 8;
            this.label30.Text = "1.提取数据";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.ForeColor = System.Drawing.Color.Red;
            this.label31.Location = new System.Drawing.Point(650, 81);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(76, 13);
            this.label31.TabIndex = 8;
            this.label31.Text = "2.粘贴到GRR";
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.ForeColor = System.Drawing.Color.Red;
            this.label33.Location = new System.Drawing.Point(532, 269);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(48, 13);
            this.label33.TabIndex = 8;
            this.label33.Text = "4(option)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 506);
            this.Controls.Add(this.btn_CopyGRRModuleAndDelete);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label33);
            this.Controls.Add(this.label31);
            this.Controls.Add(this.label30);
            this.Controls.Add(this.label29);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textB_TargetPath);
            this.Controls.Add(this.textB_SourcePath);
            this.Controls.Add(this.richTB_Log);
            this.Controls.Add(this.btn_PasteToGRR);
            this.Controls.Add(this.btn_ExtractData);
            this.Controls.Add(this.btn_SelectTargetPath);
            this.Controls.Add(this.btn_SelectSourcePath);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textB_TargetPath;
        private System.Windows.Forms.TextBox textB_SourcePath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.RichTextBox richTB_Log;
        private System.Windows.Forms.Button btn_SelectTargetPath;
        private System.Windows.Forms.Button btn_SelectSourcePath;
        private System.Windows.Forms.TextBox textB_NumSN;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textB_StartRow;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox textB_StartCol;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox textB_StartRowDest;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox textB_StartColDest;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox textB_TotalItem;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox textB_FromSheet;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox textB_ToSheet;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textB_Repeat;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Button btn_ExtractData;
        private System.Windows.Forms.Button btn_PasteToGRR;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label label20;
        private System.Windows.Forms.TextBox textB_NumSN_GRR;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox textB_StartRow_GRR;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.TextBox textB_StartCol_GRR;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox textB_TrialsCount_GRR;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.TextBox textB_StartRowDest_GRR;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.TextBox textB_StartColDest_GRR;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox textB_FromSheet_GRR;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_ExtracSheetToTxt;
        private System.Windows.Forms.TextBox textB_StartColLimit;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox textB_StartRowLimit;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.TextBox textB_StartColDestLimit;
        private System.Windows.Forms.TextBox textB_StartRowDestLimit;
        private System.Windows.Forms.Label label25;
        private System.Windows.Forms.Label label26;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.TextBox textB_FromSheetLimit;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btn_PasteLimit;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label28;
        private System.Windows.Forms.Label label24;
        private System.Windows.Forms.Button btn_CopyGRRModuleAndDelete;
        private System.Windows.Forms.Label label29;
        private System.Windows.Forms.Label label32;
        private System.Windows.Forms.Label label30;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.Label label33;
    }
}

