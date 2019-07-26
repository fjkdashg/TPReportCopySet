using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TUFO公式复制工具
{
    public partial class A : Form
    {
        WorkSet WS;
        DataTable ChildDT;

        public A(WorkSet _WS,DataTable _ChildDT)
        {
            InitializeComponent();

            WS = _WS;
            ChildDT = _ChildDT;
            this.OuterPath.Text = WS.Work_Path.OuterPath;
            this.BaseBNo.Text = WS.BaseBNo;
            this.BaseJNo.Text = WS.BaseJNo;
        }

        public void StartWork_Click(object sender, EventArgs e)
        {
            BackgroundWorker BW = new BackgroundWorker();
            BW.WorkerReportsProgress = true;
            BW.DoWork += Work_Do;
            BW.ProgressChanged += Work_Progress;
            BW.RunWorkerCompleted += Work_Done;

            this.progressBar1.Maximum = ChildDT.Rows.Count +1;
            List<object> arg = new List<object>();
            arg.Add(WS);
            arg.Add(ChildDT);
            BW.RunWorkerAsync(arg);

        }

        public void Work_Done(object sender, RunWorkerCompletedEventArgs e)
        {
            this.progressBar1.Value = this.progressBar1.Maximum;
        }

        public void Work_Do(object sender, DoWorkEventArgs e)
        {
            List<object> arg = (List<object>)e.Argument;
            BackgroundWorker BW = (BackgroundWorker)sender;
            WorkSet WS = (WorkSet)arg[0];
            DataTable ChildDT = (DataTable)arg[1];
            int si = 1;
            foreach (DataRow dr in ChildDT.Rows)
            {
                string filename = dr["ShortTitle"].ToString() + "-" + dr["ChildIndex"].ToString() + ".txt";
                if (string.IsNullOrWhiteSpace(dr["ChildIndex"].ToString()) || string.IsNullOrWhiteSpace(dr["ShortTitle"].ToString()) || string.IsNullOrWhiteSpace(dr["BNo"].ToString()) || string.IsNullOrWhiteSpace(dr["JNo"].ToString()))
                {
                    MessageBox.Show("关键参数不能为空");
                    BW.ReportProgress(si, filename + "  有空参数");
                }
                else
                {
                    using (StreamReader sr = new StreamReader(WS.Work_Path.BaseRepTemp_Path))
                    {
                        string outtxtfilepath = WS.Work_Path.OuterPath +"/" + filename;
                        if (!File.Exists(outtxtfilepath))
                        {
                            FileStream fs = new FileStream(outtxtfilepath, FileMode.Create, FileAccess.Write);
                            fs.Close();
                        }
                        using (StreamWriter sw = new StreamWriter(outtxtfilepath))
                        {
                            string line;
                            // 从文件读取并显示行，直到文件的末尾 
                            while ((line = sr.ReadLine()) != null)
                            {
                                sw.Write("");
                                sw.WriteLine(line.Replace(WS.BaseBNo, dr["BNo"].ToString()).Replace(WS.BaseJNo, dr["JNo"].ToString()));
                            }
                            sw.Close();
                            BW.ReportProgress(si, filename + "  写入完成");
                        }
                        sr.Close();
                    }
                }
                si++;
            }
        }

        public void Work_Progress(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
            this.WorkMsg.AppendText("\n" + e.UserState.ToString());
            this.WorkMsg.ScrollToCaret();
        }

        private void BackStep_Click(object sender, EventArgs e)
        {
            main.ShowNewForm(this.MdiParent, new ChildListSet(WS));
        }
    }
}
