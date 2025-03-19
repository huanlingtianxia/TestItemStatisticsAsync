namespace TestItemStatistics
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
            this.label3 = new System.Windows.Forms.Label();
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
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
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
            this.textB_TargetPath.Text = "E:\\labview\\other prj\\IGBT cplusplus dll\\MSA1\\test1\\testSim\\GRR_20250317_D471_FCT1" +
    "_No.1&2&3_Test.xlsx";
            // 
            // textB_SourcePath
            // 
            this.textB_SourcePath.Location = new System.Drawing.Point(80, 6);
            this.textB_SourcePath.Name = "textB_SourcePath";
            this.textB_SourcePath.Size = new System.Drawing.Size(564, 20);
            this.textB_SourcePath.TabIndex = 7;
            this.textB_SourcePath.Text = "E:\\labview\\other prj\\IGBT cplusplus dll\\MSA1\\test1\\testSim\\op4_Test.xlsx";
            // 
            // richTB_Log
            // 
            this.richTB_Log.Location = new System.Drawing.Point(0, 209);
            this.richTB_Log.Name = "richTB_Log";
            this.richTB_Log.Size = new System.Drawing.Size(732, 173);
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
            this.btn_SelectTargetPath.Text = "TargetPath";
            this.btn_SelectTargetPath.UseVisualStyleBackColor = true;
            this.btn_SelectTargetPath.Click += new System.EventHandler(this.btn_SelectTargetPath_Click);
            // 
            // btn_SelectSourcePath
            // 
            this.btn_SelectSourcePath.Location = new System.Drawing.Point(650, 6);
            this.btn_SelectSourcePath.Name = "btn_SelectSourcePath";
            this.btn_SelectSourcePath.Size = new System.Drawing.Size(82, 23);
            this.btn_SelectSourcePath.TabIndex = 12;
            this.btn_SelectSourcePath.Text = "SourcePath";
            this.btn_SelectSourcePath.UseVisualStyleBackColor = true;
            this.btn_SelectSourcePath.Click += new System.EventHandler(this.btn_SelectSourcePath_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 193);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(25, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Log";
            // 
            // textB_NumSN
            // 
            this.textB_NumSN.Location = new System.Drawing.Point(174, 22);
            this.textB_NumSN.Name = "textB_NumSN";
            this.textB_NumSN.Size = new System.Drawing.Size(35, 20);
            this.textB_NumSN.TabIndex = 15;
            this.textB_NumSN.Text = "8";
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
            this.btn_PasteToGRR.Location = new System.Drawing.Point(650, 59);
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
            this.label22.Location = new System.Drawing.Point(547, 93);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(175, 91);
            this.label22.TabIndex = 8;
            this.label22.Text = "1. 排序转置\r\n2. 选择路径：sourcePath是提取\r\n   和GRR的源文件共同路径\r\n3. 设置参数\r\n4. 提取数据\r\n5. 将提取数据粘贴到GRR" +
    "模板：\r\n   TargetPath";
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
            this.textB_NumSN_GRR.Text = "8";
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
            this.groupBox2.Text = "Paste to GRR module: SourcePaht to TargetPaht";
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
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 391);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
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
            this.Controls.Add(this.label3);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
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
        private System.Windows.Forms.Label label3;
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
    }
}

