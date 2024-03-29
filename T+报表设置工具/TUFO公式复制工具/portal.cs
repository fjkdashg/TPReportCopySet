﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TUFO公式复制工具
{
    public partial class portal : Form
    {
        WorkSet WS;
        public portal(WorkSet _WS)
        {
            InitializeComponent();

            WS = _WS;
            if (WS.Work_Path != null)
            {
                if (!string.IsNullOrEmpty(WS.Work_Path.OuterPath))
                {
                    this.OuterPath.Text = WS.Work_Path.OuterPath;
                }
                if (!string.IsNullOrEmpty(WS.Work_Path.BaseRepTemp_Path))
                {
                    this.BaseRepTemp_Path.Text = WS.Work_Path.BaseRepTemp_Path;
                }

                if (!string.IsNullOrEmpty(WS.Work_Path.BCD_Path))
                {
                    this.BCD_Path.Text = WS.Work_Path.BCD_Path;
                }
                if (!string.IsNullOrEmpty(WS.Work_Path.JCD_Path))
                {
                    this.JCD_Path.Text = WS.Work_Path.JCD_Path;
                }

                if (!string.IsNullOrEmpty(WS.Work_Path.BZCFZ_Path))
                {
                    this.BZCFZ_Path.Text = WS.Work_Path.BZCFZ_Path;
                }
                if (!string.IsNullOrEmpty(WS.Work_Path.JZCFZ_Path))
                {
                    this.JZCFZ_Path.Text = WS.Work_Path.JZCFZ_Path;
                }

                if (!string.IsNullOrEmpty(WS.Work_Path.BLR_Path))
                {
                    this.BLR_Path.Text = WS.Work_Path.BLR_Path;
                }
                if (!string.IsNullOrEmpty(WS.Work_Path.JLR_Path))
                {
                    this.JLR_Path.Text = WS.Work_Path.JLR_Path;
                }
            }
            else
            {
                WS.Work_Path = new WorkPath();
            }
        }

        protected override void OnLoad(EventArgs e)
        {
            WS.Work_Path.OuterPath = @"E:\Temp\北京";
            WS.Work_Path.BaseRepTemp_Path = @"E:\Temp\北京\北京瑞克博云-1.txt";
            WS.Work_Path.BCD_Path = @"E:\Temp\北京\冲抵数据汇总表.txt";
            WS.Work_Path.JCD_Path= @"E:\Temp\北京\冲抵数据尽调汇总表.txt";
            WS.Work_Path.BZCFZ_Path= @"E:\Temp\北京\横排资产负债工作表.txt";
            WS.Work_Path.JZCFZ_Path= @"E:\Temp\北京\横排资产负债尽调工作表.txt";
            WS.Work_Path.BLR_Path= @"E:\Temp\北京\横排利润工作表.txt";
            WS.Work_Path.JLR_Path= @"E:\Temp\北京\横排利润尽调工作表.txt";

            if (!string.IsNullOrEmpty(WS.Work_Path.OuterPath))
            {
                this.OuterPath.Text = WS.Work_Path.OuterPath;
            }
            if (!string.IsNullOrEmpty(WS.Work_Path.BaseRepTemp_Path))
            {
                this.BaseRepTemp_Path.Text = WS.Work_Path.BaseRepTemp_Path;
            }

            if (!string.IsNullOrEmpty(WS.Work_Path.BCD_Path))
            {
                this.BCD_Path.Text = WS.Work_Path.BCD_Path;
            }
            if (!string.IsNullOrEmpty(WS.Work_Path.JCD_Path))
            {
                this.JCD_Path.Text = WS.Work_Path.JCD_Path;
            }

            if (!string.IsNullOrEmpty(WS.Work_Path.BZCFZ_Path))
            {
                this.BZCFZ_Path.Text = WS.Work_Path.BZCFZ_Path;
            }
            if (!string.IsNullOrEmpty(WS.Work_Path.JZCFZ_Path))
            {
                this.JZCFZ_Path.Text = WS.Work_Path.JZCFZ_Path;
            }

            if (!string.IsNullOrEmpty(WS.Work_Path.BLR_Path))
            {
                this.BLR_Path.Text = WS.Work_Path.BLR_Path;
            }
            if (!string.IsNullOrEmpty(WS.Work_Path.JLR_Path))
            {
                this.JLR_Path.Text = WS.Work_Path.JLR_Path;
            }
        }
        private void GetPath_Click(object sender, EventArgs e)
        {
            Button BTN = (Button)sender;
            
            if (BTN.Name == "GetOuterPath")
            {
                FolderBrowserDialog dialog = new FolderBrowserDialog();
                dialog.Description = "请选输出路径";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string file = dialog.SelectedPath;
                    BTN.Parent.GetNextControl(BTN, false).Text = file;
                }
            }
            else
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;//该值确定是否可以选择多个文件
                dialog.Title = "请选择公式文件";
                dialog.Filter = "文本文件(*.txt)|*.txt";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string file = dialog.FileName;
                    BTN.Parent.GetNextControl(BTN, false).Text = file;
                }
            }
        }

        private void NextStep_Click(object sender, EventArgs e)
        {
            if (WS.Work_Path == null)
            {
                WS.Work_Path = new WorkPath();
            }
            WS.Work_Path.OuterPath = this.OuterPath.Text;
            WS.Work_Path.BaseRepTemp_Path = this.BaseRepTemp_Path.Text;
            WS.Work_Path.BCD_Path = this.BCD_Path.Text;
            WS.Work_Path.JCD_Path = this.JCD_Path.Text;
            WS.Work_Path.BZCFZ_Path = this.BZCFZ_Path.Text;
            WS.Work_Path.JZCFZ_Path = this.JZCFZ_Path.Text;
            WS.Work_Path.BLR_Path = this.BLR_Path.Text;
            WS.Work_Path.JLR_Path = this.JLR_Path.Text;

            if ((!string.IsNullOrWhiteSpace(WS.Work_Path.OuterPath)) && (!string.IsNullOrWhiteSpace(WS.Work_Path.BaseRepTemp_Path)))
            {
                main.ShowNewForm(this.MdiParent, new ChildListSet(WS));
            }
            else
            {
                MessageBox.Show("关键参数不能为空");
            }
        }


        private void Path_TextChanged(object sender, EventArgs e)
        {
            TextBox ak = (TextBox)sender;
            if (string.IsNullOrWhiteSpace(ak.Text))
            {

            }
        }
        private void OuterPath_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
