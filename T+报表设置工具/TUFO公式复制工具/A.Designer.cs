namespace TUFO公式复制工具
{
    partial class A
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
            this.BaseJNo = new System.Windows.Forms.TextBox();
            this.WriteNewTxt = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.OuterPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.WorkMsg = new System.Windows.Forms.RichTextBox();
            this.BackStep = new System.Windows.Forms.Button();
            this.SetB = new System.Windows.Forms.Button();
            this.SetJ = new System.Windows.Forms.Button();
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
            // BaseJNo
            // 
            this.BaseJNo.Location = new System.Drawing.Point(396, 59);
            this.BaseJNo.Name = "BaseJNo";
            this.BaseJNo.ReadOnly = true;
            this.BaseJNo.Size = new System.Drawing.Size(123, 25);
            this.BaseJNo.TabIndex = 3;
            // 
            // WriteNewTxt
            // 
            this.WriteNewTxt.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.WriteNewTxt.Location = new System.Drawing.Point(122, 139);
            this.WriteNewTxt.Name = "WriteNewTxt";
            this.WriteNewTxt.Size = new System.Drawing.Size(167, 36);
            this.WriteNewTxt.TabIndex = 5;
            this.WriteNewTxt.Text = "生成基础公式";
            this.WriteNewTxt.UseVisualStyleBackColor = true;
            this.WriteNewTxt.Click += new System.EventHandler(this.StartWork_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(20, 101);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(623, 32);
            this.progressBar1.TabIndex = 6;
            // 
            // OuterPath
            // 
            this.OuterPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.OuterPath.Location = new System.Drawing.Point(120, 13);
            this.OuterPath.Name = "OuterPath";
            this.OuterPath.ReadOnly = true;
            this.OuterPath.Size = new System.Drawing.Size(523, 25);
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
            this.label2.Location = new System.Drawing.Point(292, 64);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 15);
            this.label2.TabIndex = 16;
            this.label2.Text = "模板尽调编码";
            // 
            // WorkMsg
            // 
            this.WorkMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.WorkMsg.Location = new System.Drawing.Point(1, 181);
            this.WorkMsg.Name = "WorkMsg";
            this.WorkMsg.ReadOnly = true;
            this.WorkMsg.Size = new System.Drawing.Size(654, 266);
            this.WorkMsg.TabIndex = 19;
            this.WorkMsg.Text = "";
            // 
            // BackStep
            // 
            this.BackStep.Location = new System.Drawing.Point(20, 139);
            this.BackStep.Name = "BackStep";
            this.BackStep.Size = new System.Drawing.Size(93, 36);
            this.BackStep.TabIndex = 20;
            this.BackStep.Text = "上一步";
            this.BackStep.UseVisualStyleBackColor = true;
            this.BackStep.Click += new System.EventHandler(this.BackStep_Click);
            // 
            // SetB
            // 
            this.SetB.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SetB.Location = new System.Drawing.Point(299, 139);
            this.SetB.Name = "SetB";
            this.SetB.Size = new System.Drawing.Size(167, 36);
            this.SetB.TabIndex = 21;
            this.SetB.Text = "处理标准合并公式";
            this.SetB.UseVisualStyleBackColor = true;
            this.SetB.Click += new System.EventHandler(this.SetB_Click);
            // 
            // SetJ
            // 
            this.SetJ.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SetJ.Location = new System.Drawing.Point(476, 139);
            this.SetJ.Name = "SetJ";
            this.SetJ.Size = new System.Drawing.Size(167, 36);
            this.SetJ.TabIndex = 22;
            this.SetJ.Text = "处理尽调合并公式";
            this.SetJ.UseVisualStyleBackColor = true;
            this.SetJ.Click += new System.EventHandler(this.SetJ_Click);
            // 
            // A
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(655, 450);
            this.Controls.Add(this.SetJ);
            this.Controls.Add(this.SetB);
            this.Controls.Add(this.BackStep);
            this.Controls.Add(this.WriteNewTxt);
            this.Controls.Add(this.WorkMsg);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OuterPath);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.BaseJNo);
            this.Controls.Add(this.BaseBNo);
            this.Name = "A";
            this.Text = "M";
            this.Activated += new System.EventHandler(this.A_Activated);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox BaseBNo;
        private System.Windows.Forms.TextBox BaseJNo;
        private System.Windows.Forms.Button WriteNewTxt;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.TextBox OuterPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox WorkMsg;
        private System.Windows.Forms.Button BackStep;
        private System.Windows.Forms.Button SetB;
        private System.Windows.Forms.Button SetJ;
    }
}