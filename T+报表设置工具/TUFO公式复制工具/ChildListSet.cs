﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TUFO公式复制工具
{
    public partial class ChildListSet : Form
    {
        WorkSet WS;

        DataTable ChildDT = new DataTable();
        DataTable ExcelDT = new DataTable();
        public ChildListSet(WorkSet _WS)
        {
            InitializeComponent();
            WS = _WS;
            this.BNoCol.Value = WS.BNoCol;
            this.JNoCol.Value = WS.JNoCol;
            this.ChildStartRow.Value = WS.ChildStartRow;
            
            Encoding fe = Encoding.Default;
            try
            {
                fe = Encoding.GetEncoding(WS.Work_Path.BaseRepTemp_Path);
            }
            catch
            {
                fe = Encoding.Default;
            }
            Console.WriteLine(fe.EncodingName);

            List<string> ztno = new List<string>();
            string pattern = @"[^(],""\d{6}""";
            using (StreamReader sr = new StreamReader(WS.Work_Path.BaseRepTemp_Path))
            {
                string line;
                // 从文件读取并显示行，直到文件的末尾 
                while ((line = sr.ReadLine()) != null)
                {
                    foreach (Match match in Regex.Matches(line, pattern))
                    {
                        if (ztno.IndexOf(match.Value.Replace(",\"","").Replace("\"","")) < 0)
                        {
                            ztno.Add(match.Value.Replace(",\"", "").Replace("\"", ""));
                        }
                    }
                }
                sr.Close();
            }
            if (ztno.Count >= 2)
            {
                this.BaseBNo.Text = ztno[0];
                this.BaseJNo.Text = ztno[1];
            }
            foreach (string no in ztno)
            {
                string nos = no;
                if (BaseNoErrList.Text.Length > 1)
                {
                    nos = "," + nos;
                }
                BaseNoErrList.AppendText(nos);
            }
        }



        private void GetExcelPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择公式文件";
            dialog.Filter = "文本文件(*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string file = dialog.FileName;
                this.ZTListPath.Text = file;
            }
        }

        private void ChildListSet_Load(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {
            string temp = this.BaseBNo.Text;
            this.BaseBNo.Text = this.BaseJNo.Text;
            this.BaseJNo.Text = temp;
        }


        private void Col_ValueChanged(object sender, EventArgs e)
        {
            this.BNoCol_Str.Text = IndexToASCiiTitle(this.BNoCol.Value);
            this.JNoCol_Str.Text = IndexToASCiiTitle(this.JNoCol.Value);

            ChildDT = new DataTable();
            DataColumn c1 = new DataColumn();
            c1.ColumnName = "ChildIndex";
            c1.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c1);

            DataColumn c2 = new DataColumn();
            c2.ColumnName = "ShortTitle";
            c2.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c2);

            DataColumn c3 = new DataColumn();
            c3.ColumnName = "RepNo";
            c3.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c3);

            DataColumn c4 = new DataColumn();
            c4.ColumnName = "BNo";
            c4.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c4);

            DataColumn c5 = new DataColumn();
            c5.ColumnName = "JNo";
            c5.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c5);

            DataColumn c6 = new DataColumn();
            c6.ColumnName = "CD_Cname";
            c6.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c6);

            DataColumn c7 = new DataColumn();
            c7.ColumnName = "zcfz1_Cname";
            c7.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c7);

            DataColumn c8 = new DataColumn();
            c8.ColumnName = "zcfz2_Cname";
            c8.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c8);

            DataColumn c9 = new DataColumn();
            c9.ColumnName = "lr1_Cname";
            c9.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c9);

            DataColumn c10 = new DataColumn();
            c10.ColumnName = "lr2_Cname";
            c10.DataType = Type.GetType("System.String");
            ChildDT.Columns.Add(c10);

            if (ExcelDT.Rows.Count > 0)
            {
                for (decimal i = this.ChildStartRow.Value; i < ExcelDT.Rows.Count; i++)
                {
                    int r = int.Parse(i.ToString());
                    int bc = int.Parse(this.BNoCol.Value.ToString());
                    int jc = int.Parse(this.JNoCol.Value.ToString());
                    int ci = bc - 4;
                    int st = bc - 2;
                    int rn = bc - 3;

                    int cd = bc - 1;
                    int zcfz1 = jc + 4;
                    int zcfz2 = jc + 5;
                    int lr1 = jc + 2;
                    int lr2 = jc + 3;

                    if (!string.IsNullOrWhiteSpace(ExcelDT.Rows[r][bc].ToString()))
                    {
                        DataRow dr = ChildDT.NewRow();
                        dr["ChildIndex"] = ExcelDT.Rows[r][ci].ToString();
                        dr["ShortTitle"] = ExcelDT.Rows[r][st].ToString();
                        dr["RepNo"] = ExcelDT.Rows[r][rn].ToString();
                        dr["BNo"] = ExcelDT.Rows[r][bc].ToString();
                        dr["JNo"] = ExcelDT.Rows[r][jc].ToString();

                        dr["CD_Cname"] = ExcelDT.Rows[r][cd].ToString();
                        dr["zcfz1_Cname"] = ExcelDT.Rows[r][zcfz1].ToString();
                        dr["zcfz2_Cname"] = ExcelDT.Rows[r][zcfz2].ToString();
                        dr["lr1_Cname"] = ExcelDT.Rows[r][lr1].ToString();
                        dr["lr2_Cname"] = ExcelDT.Rows[r][lr2].ToString();

                        ChildDT.Rows.Add(dr);
                    }
                }
            }
            DGV.DataSource = ChildDT;
        }

        public string IndexToASCiiTitle(decimal Index)
        {
            List<decimal> IndexList = new List<decimal>();
            for (int i = 1; i <= Index; i++)
            {
                if (IndexList.Count <= 0)
                {
                    IndexList.Add(0);
                }
                IndexList[0] += 1;
                for (int j = 0; j < IndexList.Count; j++)
                {
                    if (IndexList[j] == 27)
                    {
                        IndexList[j] = 1;
                        if (j + 1 == IndexList.Count)
                        {
                            IndexList.Add(1);
                        }
                        else
                        {
                            IndexList[j + 1] += 1;
                        }
                    }
                }
            }
            IndexList.Reverse();
            string ShowCS = "";
            foreach (decimal IndexOut in IndexList)
            {
                ShowCS += Encoding.ASCII.GetString(new byte[] { (byte)(IndexOut + 64) });
            }
            return ShowCS;
        }

        private void NextStep_Click(object sender, EventArgs e)
        {
            WS.BaseBNo = this.BaseBNo.Text;
            WS.BaseJNo = this.BaseJNo.Text;
            WS.ChildStartRow = this.ChildStartRow.Value;
            WS.BNoCol = this.BNoCol.Value;
            WS.JNoCol = this.JNoCol.Value;

            if (DGV.Rows.Count <= 0)
            {
                Form n = new M(WS);
                n.Show();
            }
            else
            {
                main.ShowNewForm(this.MdiParent, new A(WS, ChildDT));
            }
        }

        private void ZTListPath_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(ZTListPath.Text))
            {
                string connstr = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="+ ZTListPath.Text.Trim()+ ";Extended Properties='Excel 12.0;HDR=false;IMEX=2;';Persist Security Info=False";
                OleDbConnection ODC = new OleDbConnection(connstr);
                ODC.Open();
                DataTable schemaTable = ODC.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);

                SheetTable.DataSource = schemaTable;
                SheetTable.DisplayMember = "name";
                SheetTable.ValueMember = "name";

                /*
                OleDbDataAdapter odda = new OleDbDataAdapter("select * from ["+ schemaTable.Rows[0][2].ToString().Trim() + "]", connstr);
                DataTable dt = new DataTable();
                odda.Fill(ExcelDT);
                Col_ValueChanged(new object(), new EventArgs());
                */
                ODC.Close();
            }
        }

        private void Label1_DoubleClick(object sender, EventArgs e)
        {
            WS.BaseBNo = this.BaseBNo.Text;
            WS.BaseJNo = this.BaseJNo.Text;
            WS.ChildStartRow = this.ChildStartRow.Value;
            WS.BNoCol = this.BNoCol.Value;
            WS.JNoCol = this.JNoCol.Value;

            if (MessageBox.Show("要手动执行吗？", "", buttons: MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                //main.ShowNewForm(this.MdiParent, new M(WS));
                Form m = new M(WS);
                m.Show();
            }
        }

        private void BackStep_Click(object sender, EventArgs e)
        {
            main.ShowNewForm(this.MdiParent, new portal(WS));
        }
    }
}
