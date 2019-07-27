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
    public partial class JRepWork : Form
    {
        WorkSet WS;
        DataTable ChildDT;
        DataTable ChildDTin;

        TempRepCol trc=new TempRepCol();
        public JRepWork(WorkSet _ws,DataTable _ChildDT)
        {
            InitializeComponent();
            WS = _ws;
            ChildDTin = _ChildDT;

            
        }

        private void JRepWork_Activated(object sender, EventArgs e)
        {
            this.JCD_Path.Text = WS.Work_Path.JCD_Path;
            this.JZCFZ_Path.Text = WS.Work_Path.JZCFZ_Path;
            this.JLR_Path.Text = WS.Work_Path.JLR_Path;

            string LastCDC = "";
            List<string> CD_CName = new List<string>();
            using (StreamReader sr = new StreamReader(this.JCD_Path.Text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.IndexOf("=") >= 0)
                    {
                        string pattern = @"^[a-z,A-Z]{1,2}";
                        foreach (Match match in Regex.Matches(line, pattern))
                        {
                            LastCDC = match.Value;
                            CD_CName.Add(match.Value);
                        }
                    }
                }
                sr.Close();
            }
            //Console.WriteLine(LastCDC);

            string Lastzcfz2C = "";
            List<string> zcfz2_CName = new List<string>();
            using (StreamReader sr = new StreamReader(this.JZCFZ_Path.Text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.IndexOf("=") >= 0)
                    {
                        string pattern = @"^[a-z,A-Z]{1,2}";
                        foreach (Match match in Regex.Matches(line, pattern))
                        {
                            Lastzcfz2C = match.Value;
                            zcfz2_CName.Add(match.Value);
                        }
                    }
                }
                sr.Close();
            }
            //Console.WriteLine(Lastzcfz2C);

            string Lastlr2C = "";
            List<string> lr2_CName = new List<string>();
            using (StreamReader sr = new StreamReader(this.JLR_Path.Text))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.IndexOf("=") >= 0)
                    {
                        string pattern = @"^[a-z,A-Z]{1,2}";
                        foreach (Match match in Regex.Matches(line, pattern))
                        {
                            Lastlr2C = match.Value;
                            lr2_CName.Add(match.Value);
                        }
                    }
                }
                sr.Close();
            }
            //Console.WriteLine(Lastlr2C);
            ChildDT = new DataTable();
            ChildDT = ChildDTin.Clone();
            foreach (DataRow dr in ChildDTin.Rows)
            {
                if (CD_CName.IndexOf(dr["CD_Cname"].ToString()) < 0 || zcfz2_CName.IndexOf(dr["zcfz2_Cname"].ToString()) < 0 || lr2_CName.IndexOf(dr["lr2_Cname"].ToString()) < 0)
                {
                    DataRow ak = ChildDT.NewRow();
                    for (int i = 0; i < ChildDTin.Columns.Count; i++)
                    {
                        ak[i] = dr[i];
                    }
                    ChildDT.Rows.Add(ak);
                }
                if (dr["CD_Cname"].ToString() == LastCDC && dr["zcfz2_Cname"].ToString() == Lastzcfz2C && dr["lr2_Cname"].ToString() == Lastlr2C)
                {
                    trc.RepNo = dr["RepNo"].ToString();
                    trc.cd = LastCDC;
                    trc.zcfz1 = dr["zcfz1_Cname"].ToString();
                    trc.zcfz2 = Lastzcfz2C;
                    trc.lr1 = dr["lr1_Cname"].ToString();
                    trc.lr2 = Lastlr2C;
                    trc.state = true;
                }
            }
            DGV1.DataSource = ChildDT;

            DataGridViewCheckBoxColumn dgvc = new DataGridViewCheckBoxColumn();
            dgvc.Name = "Select";
            dgvc.HeaderText = "添加进公式";
            DGV1.Columns.Insert(0, dgvc);
        }

        public void Work_Done(object sender, RunWorkerCompletedEventArgs e)
        {
            this.progressBar1.Value = this.progressBar1.Maximum;
        }

        public void Work_Progress(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = this.progressBar1.Value+1;
            //this.WorkMsg.AppendText("\n" + e.UserState.ToString());
            //this.WorkMsg.ScrollToCaret();
        }

        
        private void Button1_Click(object sender, EventArgs e)
        {
            this.progressBar1.Maximum = DGV1.Rows.Count * 4 + 6;
            this.progressBar1.Value = 0;

            BackgroundWorker BW = new BackgroundWorker();
            BW.WorkerReportsProgress = true;
            BW.DoWork += Work_Do;
            BW.ProgressChanged += Work_Progress;
            BW.RunWorkerCompleted += Work_Done;

            List<object> arg = new List<object>();
            arg.Add(WS);
            arg.Add(DGV1);
            arg.Add(trc);
            BW.RunWorkerAsync(arg);
        }

        public void Work_Do(object sender, DoWorkEventArgs e)
        {
            List<object> arg = (List<object>)e.Argument;
            BackgroundWorker BW = (BackgroundWorker)sender;
            WorkSet WS = (WorkSet)arg[0];
            DataGridView DGV1 = (DataGridView)arg[1];
            TempRepCol _trc = (TempRepCol)arg[2];

            BW.ReportProgress(1);
            if (_trc.state)
            {
                List<string> cd_fl = new List<string>();
                List<string> zcfz_fl = new List<string>();
                List<string> lr_fl = new List<string>();

                using (StreamReader sr = new StreamReader(this.JCD_Path.Text))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.IndexOf("=") >= 0)
                        {
                            if (line.Substring(0, _trc.cd.Length) == _trc.cd)
                            {
                                cd_fl.Add(line);
                                //Console.WriteLine(line);
                            }
                        }
                    }
                    sr.Close();
                }
                BW.ReportProgress(1);
                using (StreamReader sr = new StreamReader(this.JZCFZ_Path.Text))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.IndexOf("=") >= 0)
                        {
                            if (line.Substring(0, _trc.zcfz1.Length) == _trc.zcfz1)
                            {
                                zcfz_fl.Add(line);
                                //Console.WriteLine(line);
                            }
                            if (line.Substring(0, _trc.zcfz2.Length) == _trc.zcfz2)
                            {
                                zcfz_fl.Add(line);
                                //Console.WriteLine(line);
                            }
                        }
                    }
                    sr.Close();
                }
                BW.ReportProgress(1);
                using (StreamReader sr = new StreamReader(this.JLR_Path.Text))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.IndexOf("=") >= 0)
                        {
                            if (line.Substring(0, _trc.lr1.Length) == _trc.lr1)
                            {
                                lr_fl.Add(line);
                                //Console.WriteLine(line);
                            }
                            if (line.Substring(0, _trc.lr2.Length) == _trc.lr2)
                            {
                                lr_fl.Add(line);
                                //Console.WriteLine(line);
                            }
                        }
                    }
                    sr.Close();
                }
                BW.ReportProgress(1);
                foreach (DataGridViewRow vr in DGV1.Rows)
                {
                    //
                    Boolean isselect = (Boolean)vr.Cells["Select"].EditedFormattedValue;
                    if ((Boolean)vr.Cells["Select"].EditedFormattedValue)
                    {
                        using (StreamWriter sw = new StreamWriter(this.JCD_Path.Text, append: true))
                        {
                            foreach (string f in cd_fl)
                            {
                                sw.WriteLine(f.Replace(_trc.RepNo, vr.Cells["RepNo"].Value.ToString()).Replace(_trc.cd, vr.Cells["CD_Cname"].Value.ToString()));
                            }
                            sw.Close();
                            //BW.ReportProgress(si, filename + "  写入完成");
                        }
                        BW.ReportProgress(1);
                        using (StreamWriter sw = new StreamWriter(this.JZCFZ_Path.Text, append: true))
                        {
                            foreach (string f in zcfz_fl)
                            {
                                sw.WriteLine(f.Replace(_trc.RepNo, vr.Cells["RepNo"].Value.ToString()).Replace(_trc.zcfz1, vr.Cells["zcfz1_Cname"].Value.ToString()).Replace(_trc.zcfz2, vr.Cells["zcfz2_Cname"].Value.ToString()));
                            }
                            sw.Close();
                            //BW.ReportProgress(si, filename + "  写入完成");
                        }
                        BW.ReportProgress(1);
                        using (StreamWriter sw = new StreamWriter(this.JLR_Path.Text, append: true))
                        {
                            foreach (string f in lr_fl)
                            {
                                sw.WriteLine(f.Replace(_trc.RepNo, vr.Cells["RepNo"].Value.ToString()).Replace(_trc.lr1, vr.Cells["lr1_Cname"].Value.ToString()).Replace(_trc.lr2, vr.Cells["lr2_Cname"].Value.ToString()));
                            }
                            sw.Close();
                            //BW.ReportProgress(si, filename + "  写入完成");
                        }
                        BW.ReportProgress(1);
                    }
                }
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            main.ShowNewForm(this.MdiParent, new A(WS, ChildDT));
        }
    }
}
