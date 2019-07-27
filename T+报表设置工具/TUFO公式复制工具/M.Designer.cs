namespace TUFO公式复制工具
{
    partial class M
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
            this.BaseBNo = new System.Windows.Forms.TextBox();
            this.TarJNo = new System.Windows.Forms.TextBox();
            this.TarBNo = new System.Windows.Forms.TextBox();
            this.BaseJNo = new System.Windows.Forms.TextBox();
            this.TarFileName = new System.Windows.Forms.TextBox();
            this.WriteNewTxt = new System.Windows.Forms.Button();
            this.OuterPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.WorkMsg = new System.Windows.Forms.RichTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // BaseBNo
            // 
            this.BaseBNo.Location = new System.Drawing.Point(120, 59);
            this.BaseBNo.Name = "BaseBNo";
            this.BaseBNo.ReadOnly = true;
            this.BaseBNo.Size = new System.Drawing.Size(123, 25);
            this.BaseBNo.TabIndex = 0;
            // 
            // TarJNo
            // 
            this.TarJNo.Location = new System.Drawing.Point(400, 105);
            this.TarJNo.Name = "TarJNo";
            this.TarJNo.Size = new System.Drawing.Size(123, 25);
            this.TarJNo.TabIndex = 1;
            // 
            // TarBNo
            // 
            this.TarBNo.Location = new System.Drawing.Point(399, 59);
            this.TarBNo.Name = "TarBNo";
            this.TarBNo.Size = new System.Drawing.Size(123, 25);
            this.TarBNo.TabIndex = 2;
            // 
            // BaseJNo
            // 
            this.BaseJNo.Location = new System.Drawing.Point(120, 105);
            this.BaseJNo.Name = "BaseJNo";
            this.BaseJNo.ReadOnly = true;
            this.BaseJNo.Size = new System.Drawing.Size(123, 25);
            this.BaseJNo.TabIndex = 3;
            // 
            // TarFileName
            // 
            this.TarFileName.Location = new System.Drawing.Point(120, 159);
            this.TarFileName.Name = "TarFileName";
            this.TarFileName.Size = new System.Drawing.Size(244, 25);
            this.TarFileName.TabIndex = 4;
            // 
            // WriteNewTxt
            // 
            this.WriteNewTxt.Location = new System.Drawing.Point(381, 153);
            this.WriteNewTxt.Name = "WriteNewTxt";
            this.WriteNewTxt.Size = new System.Drawing.Size(142, 36);
            this.WriteNewTxt.TabIndex = 5;
            this.WriteNewTxt.Text = "写入";
            this.WriteNewTxt.UseVisualStyleBackColor = true;
            this.WriteNewTxt.Click += new System.EventHandler(this.Button1_Click);
            // 
            // OuterPath
            // 
            this.OuterPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OuterPath.Location = new System.Drawing.Point(120, 13);
            this.OuterPath.Name = "OuterPath";
            this.OuterPath.ReadOnly = true;
            this.OuterPath.Size = new System.Drawing.Size(403, 25);
            this.OuterPath.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(67, 15);
            this.label1.TabIndex = 8;
            this.label1.Text = "输出路径";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(16, 64);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(97, 15);
            this.label5.TabIndex = 15;
            this.label5.Text = "模板基础编码";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 110);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 15);
            this.label2.TabIndex = 16;
            this.label2.Text = "模板尽调编码";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(278, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 15);
            this.label3.TabIndex = 17;
            this.label3.Text = "目标基础编码";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(278, 110);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 15);
            this.label4.TabIndex = 18;
            this.label4.Text = "目标尽调编码";
            // 
            // WorkMsg
            // 
            this.WorkMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.WorkMsg.Location = new System.Drawing.Point(19, 195);
            this.WorkMsg.Name = "WorkMsg";
            this.WorkMsg.ReadOnly = true;
            this.WorkMsg.Size = new System.Drawing.Size(503, 243);
            this.WorkMsg.TabIndex = 19;
            this.WorkMsg.Text = "";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(17, 164);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(82, 15);
            this.label6.TabIndex = 20;
            this.label6.Text = "目标文件名";
            // 
            // M
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(545, 450);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.WorkMsg);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OuterPath);
            this.Controls.Add(this.WriteNewTxt);
            this.Controls.Add(this.TarFileName);
            this.Controls.Add(this.BaseJNo);
            this.Controls.Add(this.TarBNo);
            this.Controls.Add(this.TarJNo);
            this.Controls.Add(this.BaseBNo);
            this.Name = "M";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "手动执行";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox BaseBNo;
        private System.Windows.Forms.TextBox TarJNo;
        private System.Windows.Forms.TextBox TarBNo;
        private System.Windows.Forms.TextBox BaseJNo;
        private System.Windows.Forms.TextBox TarFileName;
        private System.Windows.Forms.Button WriteNewTxt;
        private System.Windows.Forms.TextBox OuterPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RichTextBox WorkMsg;
        private System.Windows.Forms.Label label6;
    }
}