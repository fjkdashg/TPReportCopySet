namespace TUFO公式复制工具
{
    partial class ChildListSet
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.ZTListPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.GetExcelPath = new System.Windows.Forms.Button();
            this.BNoCol = new System.Windows.Forms.NumericUpDown();
            this.JNoCol = new System.Windows.Forms.NumericUpDown();
            this.ChildStartRow = new System.Windows.Forms.NumericUpDown();
            this.BNoCol_Str = new System.Windows.Forms.TextBox();
            this.JNoCol_Str = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.BaseBNo = new System.Windows.Forms.TextBox();
            this.BaseJNo = new System.Windows.Forms.TextBox();
            this.BaseNoErrList = new System.Windows.Forms.RichTextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.DGV = new System.Windows.Forms.DataGridView();
            this.BackStep = new System.Windows.Forms.Button();
            this.NextStep = new System.Windows.Forms.Button();
            this.SheetTable = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.BNoCol)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.JNoCol)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChildStartRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).BeginInit();
            this.SuspendLayout();
            // 
            // ZTListPath
            // 
            this.ZTListPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ZTListPath.Location = new System.Drawing.Point(82, 10);
            this.ZTListPath.Margin = new System.Windows.Forms.Padding(2);
            this.ZTListPath.Name = "ZTListPath";
            this.ZTListPath.ReadOnly = true;
            this.ZTListPath.Size = new System.Drawing.Size(322, 21);
            this.ZTListPath.TabIndex = 0;
            this.ZTListPath.TextChanged += new System.EventHandler(this.ZTListPath_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 14);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(65, 12);
            this.label1.TabIndex = 1;
            this.label1.Text = "账套定义表";
            this.label1.DoubleClick += new System.EventHandler(this.Label1_DoubleClick);
            // 
            // GetExcelPath
            // 
            this.GetExcelPath.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.GetExcelPath.Location = new System.Drawing.Point(535, 6);
            this.GetExcelPath.Margin = new System.Windows.Forms.Padding(2);
            this.GetExcelPath.Name = "GetExcelPath";
            this.GetExcelPath.Size = new System.Drawing.Size(56, 26);
            this.GetExcelPath.TabIndex = 2;
            this.GetExcelPath.Text = "选择";
            this.GetExcelPath.UseVisualStyleBackColor = true;
            this.GetExcelPath.Click += new System.EventHandler(this.GetExcelPath_Click);
            // 
            // BNoCol
            // 
            this.BNoCol.Location = new System.Drawing.Point(274, 41);
            this.BNoCol.Margin = new System.Windows.Forms.Padding(2);
            this.BNoCol.Name = "BNoCol";
            this.BNoCol.Size = new System.Drawing.Size(36, 21);
            this.BNoCol.TabIndex = 3;
            this.BNoCol.ValueChanged += new System.EventHandler(this.Col_ValueChanged);
            // 
            // JNoCol
            // 
            this.JNoCol.Location = new System.Drawing.Point(457, 41);
            this.JNoCol.Margin = new System.Windows.Forms.Padding(2);
            this.JNoCol.Name = "JNoCol";
            this.JNoCol.Size = new System.Drawing.Size(36, 21);
            this.JNoCol.TabIndex = 4;
            this.JNoCol.ValueChanged += new System.EventHandler(this.Col_ValueChanged);
            // 
            // ChildStartRow
            // 
            this.ChildStartRow.Location = new System.Drawing.Point(113, 41);
            this.ChildStartRow.Margin = new System.Windows.Forms.Padding(2);
            this.ChildStartRow.Name = "ChildStartRow";
            this.ChildStartRow.Size = new System.Drawing.Size(36, 21);
            this.ChildStartRow.TabIndex = 5;
            this.ChildStartRow.ValueChanged += new System.EventHandler(this.Col_ValueChanged);
            // 
            // BNoCol_Str
            // 
            this.BNoCol_Str.Location = new System.Drawing.Point(321, 41);
            this.BNoCol_Str.Margin = new System.Windows.Forms.Padding(2);
            this.BNoCol_Str.Name = "BNoCol_Str";
            this.BNoCol_Str.ReadOnly = true;
            this.BNoCol_Str.Size = new System.Drawing.Size(30, 21);
            this.BNoCol_Str.TabIndex = 6;
            // 
            // JNoCol_Str
            // 
            this.JNoCol_Str.Location = new System.Drawing.Point(504, 41);
            this.JNoCol_Str.Margin = new System.Windows.Forms.Padding(2);
            this.JNoCol_Str.Name = "JNoCol_Str";
            this.JNoCol_Str.ReadOnly = true;
            this.JNoCol_Str.Size = new System.Drawing.Size(30, 21);
            this.JNoCol_Str.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 45);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 8;
            this.label2.Text = "子公司账套起始行";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(178, 45);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 12);
            this.label3.TabIndex = 9;
            this.label3.Text = "基础账套编码列";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(362, 45);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(89, 12);
            this.label4.TabIndex = 10;
            this.label4.Text = "尽调账套编码列";
            // 
            // BaseBNo
            // 
            this.BaseBNo.Location = new System.Drawing.Point(9, 90);
            this.BaseBNo.Margin = new System.Windows.Forms.Padding(2);
            this.BaseBNo.Name = "BaseBNo";
            this.BaseBNo.ReadOnly = true;
            this.BaseBNo.Size = new System.Drawing.Size(76, 21);
            this.BaseBNo.TabIndex = 11;
            // 
            // BaseJNo
            // 
            this.BaseJNo.Location = new System.Drawing.Point(148, 90);
            this.BaseJNo.Margin = new System.Windows.Forms.Padding(2);
            this.BaseJNo.Name = "BaseJNo";
            this.BaseJNo.ReadOnly = true;
            this.BaseJNo.Size = new System.Drawing.Size(76, 21);
            this.BaseJNo.TabIndex = 12;
            // 
            // BaseNoErrList
            // 
            this.BaseNoErrList.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.BaseNoErrList.Location = new System.Drawing.Point(238, 71);
            this.BaseNoErrList.Margin = new System.Windows.Forms.Padding(2);
            this.BaseNoErrList.Name = "BaseNoErrList";
            this.BaseNoErrList.ReadOnly = true;
            this.BaseNoErrList.Size = new System.Drawing.Size(213, 39);
            this.BaseNoErrList.TabIndex = 13;
            this.BaseNoErrList.Text = "";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(10, 71);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 12);
            this.label5.TabIndex = 14;
            this.label5.Text = "模板基础编码";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(150, 71);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(77, 12);
            this.label6.TabIndex = 15;
            this.label6.Text = "模板尽调编码";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(92, 71);
            this.button2.Margin = new System.Windows.Forms.Padding(2);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(48, 38);
            this.button2.TabIndex = 16;
            this.button2.Text = "交换\r\n编码\r\n";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // DGV
            // 
            this.DGV.AllowUserToAddRows = false;
            this.DGV.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGV.Location = new System.Drawing.Point(3, 115);
            this.DGV.Margin = new System.Windows.Forms.Padding(2);
            this.DGV.Name = "DGV";
            this.DGV.RowHeadersWidth = 51;
            this.DGV.RowTemplate.Height = 27;
            this.DGV.Size = new System.Drawing.Size(597, 309);
            this.DGV.TabIndex = 17;
            // 
            // BackStep
            // 
            this.BackStep.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.BackStep.Location = new System.Drawing.Point(466, 71);
            this.BackStep.Margin = new System.Windows.Forms.Padding(2);
            this.BackStep.Name = "BackStep";
            this.BackStep.Size = new System.Drawing.Size(56, 38);
            this.BackStep.TabIndex = 18;
            this.BackStep.Text = "上一步";
            this.BackStep.UseVisualStyleBackColor = true;
            this.BackStep.Click += new System.EventHandler(this.BackStep_Click);
            // 
            // NextStep
            // 
            this.NextStep.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.NextStep.Location = new System.Drawing.Point(535, 71);
            this.NextStep.Margin = new System.Windows.Forms.Padding(2);
            this.NextStep.Name = "NextStep";
            this.NextStep.Size = new System.Drawing.Size(56, 38);
            this.NextStep.TabIndex = 19;
            this.NextStep.Text = "下一步";
            this.NextStep.UseVisualStyleBackColor = true;
            this.NextStep.Click += new System.EventHandler(this.NextStep_Click);
            // 
            // SheetTable
            // 
            this.SheetTable.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SheetTable.FormattingEnabled = true;
            this.SheetTable.Location = new System.Drawing.Point(410, 10);
            this.SheetTable.Name = "SheetTable";
            this.SheetTable.Size = new System.Drawing.Size(121, 20);
            this.SheetTable.TabIndex = 20;
            // 
            // ChildListSet
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(602, 422);
            this.Controls.Add(this.SheetTable);
            this.Controls.Add(this.NextStep);
            this.Controls.Add(this.BackStep);
            this.Controls.Add(this.DGV);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.BaseNoErrList);
            this.Controls.Add(this.BaseJNo);
            this.Controls.Add(this.BaseBNo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.JNoCol_Str);
            this.Controls.Add(this.BNoCol_Str);
            this.Controls.Add(this.ChildStartRow);
            this.Controls.Add(this.JNoCol);
            this.Controls.Add(this.BNoCol);
            this.Controls.Add(this.GetExcelPath);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ZTListPath);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "ChildListSet";
            this.Text = "ChildListSet";
            this.Load += new System.EventHandler(this.ChildListSet_Load);
            ((System.ComponentModel.ISupportInitialize)(this.BNoCol)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.JNoCol)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChildStartRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGV)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox ZTListPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button GetExcelPath;
        private System.Windows.Forms.NumericUpDown BNoCol;
        private System.Windows.Forms.NumericUpDown JNoCol;
        private System.Windows.Forms.NumericUpDown ChildStartRow;
        private System.Windows.Forms.TextBox BNoCol_Str;
        private System.Windows.Forms.TextBox JNoCol_Str;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox BaseBNo;
        private System.Windows.Forms.TextBox BaseJNo;
        private System.Windows.Forms.RichTextBox BaseNoErrList;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView DGV;
        private System.Windows.Forms.Button BackStep;
        private System.Windows.Forms.Button NextStep;
        private System.Windows.Forms.ComboBox SheetTable;
    }
}