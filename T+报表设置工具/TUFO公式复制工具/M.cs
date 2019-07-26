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
    public partial class M : Form
    {
        WorkSet WS;
        public M(WorkSet _WS)
        {
            InitializeComponent();

            WS = _WS;
            this.OuterPath.Text = WS.Work_Path.OuterPath;
            this.BaseBNo.Text = WS.BaseBNo;
            this.BaseJNo.Text = WS.BaseJNo;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TarBNo.Text) || string.IsNullOrWhiteSpace(TarJNo.Text) || string.IsNullOrWhiteSpace(TarFileName.Text))
            {
                MessageBox.Show("关键参数不能为空");
            }
            else
            {
                using (StreamReader sr = new StreamReader(WS.Work_Path.BaseRepTemp_Path))
                {
                    string outtxtfilepath= WS.Work_Path.OuterPath + "/" + this.TarFileName.Text.Trim() + ".txt";
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
                            sw.WriteLine(line.Replace(WS.BaseBNo, TarBNo.Text).Replace(WS.BaseJNo, TarJNo.Text));
                            this.WorkMsg.AppendText("\n" + line.Replace(WS.BaseBNo, TarBNo.Text).Replace(WS.BaseJNo, TarJNo.Text));
                        }
                        sw.Close();
                    } 
                    sr.Close();
                }
            }
        }
    }
}
