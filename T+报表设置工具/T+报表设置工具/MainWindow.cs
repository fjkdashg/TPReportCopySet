using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace T_报表设置工具
{
    public partial class MainWindow : Form
    {
        SqlConnection MainSQL;

        RepXML BaseRepXML;
        RepXML CDBXML;
        RepXML CDJXML;
        RepXML ZCFZBXML;
        RepXML ZCFZJXML;
        RepXML LRBXML;
        RepXML LRJXML;

        /*
        XmlNode BaseRepXML_root;
        XmlNode CDBXML_root;
        XmlNode CDJXML_root;
        XmlNode ZCFZBXML_root;
        XmlNode ZCFZJXML_root;
        XmlNode LRBXML_root;
        XmlNode LRJXML_root;
        */
        DataTable rDT = new DataTable();
        DataTable zDT = new DataTable();


        public MainWindow()
        {
            InitializeComponent();

            //GetWorkDB_Click(new object(), new EventArgs());
            DataColumn c1 = new DataColumn();
            c1.ColumnName = "RepID";
            c1.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c1);

            DataColumn c2 = new DataColumn();
            c2.ColumnName = "RepNo";
            c2.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c2);

            DataColumn c3 = new DataColumn();
            c3.ColumnName = "RepName";
            c3.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c3);

            DataColumn c4 = new DataColumn();
            c4.ColumnName = "zName";
            c4.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c4);

            DataColumn c5 = new DataColumn();
            c5.ColumnName = "BzNo";
            c5.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c5);

            DataColumn c6 = new DataColumn();
            c6.ColumnName = "BzCDCol";
            c6.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c6);

            DataColumn c7 = new DataColumn();
            c7.ColumnName = "BzZCFZCol1";
            c7.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c7);

            DataColumn c8 = new DataColumn();
            c8.ColumnName = "BzZCFZCol2";
            c8.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c8);

            DataColumn c9 = new DataColumn();
            c9.ColumnName = "BzLRCol1";
            c9.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c9);

            DataColumn c10 = new DataColumn();
            c10.ColumnName = "BzLRCol2";
            c10.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c10);

            //------------
            DataColumn c11 = new DataColumn();
            c11.ColumnName = "JzNo";
            c11.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c11);

            DataColumn c12 = new DataColumn();
            c12.ColumnName = "JzCDCol";
            c12.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c12);

            DataColumn c13 = new DataColumn();
            c13.ColumnName = "JzZCFZCol1";
            c13.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c13);

            DataColumn c14 = new DataColumn();
            c14.ColumnName = "JzZCFZCol2";
            c14.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c14);

            DataColumn c15 = new DataColumn();
            c15.ColumnName = "JzLRCol1";
            c15.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c15);

            DataColumn c16 = new DataColumn();
            c16.ColumnName = "JzLRCol2";
            c16.DataType = Type.GetType("System.String");
            zDT.Columns.Add(c16);

        }

        protected override void OnLoad(EventArgs e)
        {
            GetWorkDB.PerformClick();
            SaveSQLConn.PerformClick();
        }

        private void GetWorkDB_Click(object sender, EventArgs e)
        {
            string ConnStr = String.Format("data source={0},{1};initial catalog=Master;user id={2};pwd={3}", SQL_Host.Text.Trim(), SQL_Port.Text.Trim(), SQL_User.Text.Trim(), SQL_Pwd.Text.Trim());
            //Console.WriteLine(ConnStr);
            MainSQL = new SqlConnection(ConnStr);
            MainSQL.Open();
            SqlDataAdapter SDA = new SqlDataAdapter("SELECT Name FROM Master..SysDatabases ORDER BY Name", MainSQL);
            DataTable dblist = new DataTable();
            SDA.Fill(dblist);
            this.SQL_WorkDB.DataSource = dblist;
            this.SQL_WorkDB.ValueMember = "Name";

            this.SQL_WorkDB.SelectedIndex = 17;
        }

        private void SaveSQLConn_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(this.SQL_WorkDB.Text))
            {
                this.WorkSetting.Enabled = true;
                this.GetWorkDB.Enabled = false;
                this.SQL_Host.Enabled = false;
                this.SQL_Port.Enabled = false;
                this.SQL_User.Enabled = false;
                this.SQL_Pwd.Enabled = false;
                this.SQL_WorkDB.Enabled = false;

                MainSQL.Close();
                string ConnStr = String.Format("data source={0},{1};initial catalog={2};user id={3};pwd={4}", SQL_Host.Text.Trim(), SQL_Port.Text.Trim(), this.SQL_WorkDB.Text.Trim(), SQL_User.Text.Trim(), SQL_Pwd.Text.Trim());
                //Console.WriteLine(ConnStr);
                MainSQL = new SqlConnection(ConnStr);
                MainSQL.Open();
                try
                {
                    SqlDataAdapter SDA = new SqlDataAdapter("select  * from   TUFO_ReportTemplateBasic where    isSysTemplate <> 1 and Creator <> 'admin' ", MainSQL);
                    SDA.Fill(rDT);

                    DataTable dblist1 = new DataTable();
                    DataTable dblist2 = new DataTable();
                    DataTable dblist3 = new DataTable();
                    DataTable dblist4 = new DataTable();
                    DataTable dblist5 = new DataTable();
                    DataTable dblist6 = new DataTable();
                    DataTable dblist7 = new DataTable();
                    SDA.Fill(dblist1);
                    SDA.Fill(dblist2);
                    SDA.Fill(dblist3);
                    SDA.Fill(dblist4);
                    SDA.Fill(dblist5);
                    SDA.Fill(dblist6);
                    SDA.Fill(dblist7);

                    this.Temp_BaseRep_Name.DataSource = dblist1;
                    this.Temp_BaseRep_No.DataSource = this.Temp_BaseRep_Name.DataSource;
                    this.Temp_BaseRep_ID.DataSource = this.Temp_BaseRep_Name.DataSource;
                    this.Temp_BaseRep_Name.ValueMember = "Name";
                    this.Temp_BaseRep_No.ValueMember = "Code";
                    this.Temp_BaseRep_ID.ValueMember = "TemplateID";

                    this.Temp_CDB_Name.DataSource = dblist2;
                    this.Temp_CDB_No.DataSource = this.Temp_CDB_Name.DataSource;
                    this.Temp_CDB_ID.DataSource = this.Temp_CDB_Name.DataSource;
                    this.Temp_CDB_Name.ValueMember = "Name";
                    this.Temp_CDB_No.ValueMember = "Code";
                    this.Temp_CDB_ID.ValueMember = "TemplateID";


                    this.Temp_CDJ_Name.DataSource = dblist3;
                    this.Temp_CDJ_No.DataSource = this.Temp_CDJ_Name.DataSource;
                    this.Temp_CDJ_ID.DataSource = this.Temp_CDJ_Name.DataSource;
                    this.Temp_CDJ_Name.ValueMember = "Name";
                    this.Temp_CDJ_No.ValueMember = "Code";
                    this.Temp_CDJ_ID.ValueMember = "TemplateID";

                    this.Temp_ZCFZB_Name.DataSource = dblist4;
                    this.Temp_ZCFZB_No.DataSource = this.Temp_ZCFZB_Name.DataSource;
                    this.Temp_ZCFZB_ID.DataSource = this.Temp_ZCFZB_Name.DataSource;
                    this.Temp_ZCFZB_Name.ValueMember = "Name";
                    this.Temp_ZCFZB_No.ValueMember = "Code";
                    this.Temp_ZCFZB_ID.ValueMember = "TemplateID";

                    this.Temp_ZCFZJ_Name.DataSource = dblist5;
                    this.Temp_ZCFZJ_No.DataSource = this.Temp_ZCFZJ_Name.DataSource;
                    this.Temp_ZCFZJ_ID.DataSource = this.Temp_ZCFZJ_Name.DataSource;
                    this.Temp_ZCFZJ_Name.ValueMember = "Name";
                    this.Temp_ZCFZJ_No.ValueMember = "Code";
                    this.Temp_ZCFZJ_ID.ValueMember = "TemplateID";

                    this.Temp_LRB_Name.DataSource = dblist6;
                    this.Temp_LRB_No.DataSource = this.Temp_LRB_Name.DataSource;
                    this.Temp_LRB_ID.DataSource = this.Temp_LRB_Name.DataSource;
                    this.Temp_LRB_Name.ValueMember = "Name";
                    this.Temp_LRB_No.ValueMember = "Code";
                    this.Temp_LRB_ID.ValueMember = "TemplateID";

                    this.Temp_LRJ_Name.DataSource = dblist7;
                    this.Temp_LRJ_No.DataSource = this.Temp_LRJ_Name.DataSource;
                    this.Temp_LRJ_ID.DataSource = this.Temp_LRJ_Name.DataSource;
                    this.Temp_LRJ_Name.ValueMember = "Name";
                    this.Temp_LRJ_No.ValueMember = "Code";
                    this.Temp_LRJ_ID.ValueMember = "TemplateID";


                    this.Temp_BaseRep_Name.SelectedIndex = 1;
                    this.Temp_CDB_Name.SelectedIndex = 18;
                    this.Temp_CDJ_Name.SelectedIndex = 23;
                    this.Temp_ZCFZB_Name.SelectedIndex = 20;
                    this.Temp_ZCFZJ_Name.SelectedIndex = 25;
                    this.Temp_LRB_Name.SelectedIndex = 19;
                    this.Temp_LRJ_Name.SelectedIndex = 24;


                    setZTCount();
                }
                catch
                {

                }
                //this.Temp_BaseRep_ZTName_CellName.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTName_Col.Value) + this.Temp_BaseRep_ZTName_Row.Value.ToString();
                Temp_BaseRep_ZTInfo_ColRow_indexChanged(new object(), new EventArgs());
                Temp_BaseRep_DataInfo_ColRow_indexChanged(new object(), new EventArgs());
                Temp_CDWorkRep_ColRow_indexChanged(new object(), new EventArgs());
                Temp_ZCFZWorkRep_ColRow_indexChanged(new object(), new EventArgs());
                Temp_LRWorkRep_ColRow_indexChanged(new object(), new EventArgs());
            }
        }

        private void GetXNL_Click(object sender, EventArgs e)
        {
            BaseRepXML=GetRepXMLData(this.Temp_BaseRep_ID.Text);

            CDBXML=GetRepXMLData(this.Temp_CDB_ID.Text);
            CDJXML=GetRepXMLData(this.Temp_CDJ_ID.Text);
            
            ZCFZBXML=GetRepXMLData(this.Temp_ZCFZB_ID.Text);
            ZCFZJXML=GetRepXMLData(this.Temp_ZCFZJ_ID.Text);
            
            LRBXML=GetRepXMLData(this.Temp_LRB_ID.Text);
            LRJXML=GetRepXMLData(this.Temp_LRJ_ID.Text);

            WorkSetting.SelectTab(BaseSetting);
            ReadBaseRepSet.PerformClick();
            WorkSetting.SelectTab(CDSetting);
            BTN_ReadCDValue.PerformClick();
            WorkSetting.SelectTab(ZCFZSetting);
            ReadZCFZValue.PerformClick();
            WorkSetting.SelectTab(LRSetting);
            ReadLRValue.PerformClick();
        }
        public RepXML GetRepXMLData(string RepID)
        {
            string sql = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", RepID);
            SqlCommand cmd = new SqlCommand(sql, MainSQL);
            SqlDataReader sdr = cmd.ExecuteReader();
            RepXML rx = new RepXML();
            while (sdr.Read())
            {
                string RepXMLStr = (string)sdr.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");
                rx.Doc= new XmlDocument();
                rx.Doc.LoadXml(RepXMLStr);
            }
            sdr.Close();
            cmd.Dispose();
            Console.WriteLine(rx.RootNode.InnerXml.ToString());
            return rx;
        }

        string GetXMLValueByRC(RepXML RX, decimal rowno, decimal colno,string VType="snix")
        {
            string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", rowno, colno);
            string rStr = "";
            switch (VType)
            {
                case "snix":
                    rStr = RX.RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
                case "snafv":
                    rStr = RX.RootNode.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    break;
                default:
                    rStr = RX.RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
            }
            return rStr;
        }

        string GetXMLValueByRC(RepXML RX, NumericUpDown rowno, NumericUpDown colno, string VType = "snix")
        {
            string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", rowno.Value, colno.Value);
            string rStr = "";
            switch (VType)
            {
                case "snix":
                    rStr= RX.RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
                case "snafv":
                    rStr= RX.RootNode.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    break;
                default:
                    rStr= RX.RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
            }
            return rStr;
        }

        void setZTCount()
        {
            decimal B_ZT_Child_Count = BCD_CEnd_C.Value - BCD_CStart_C.Value;
            ZCFZ_M_B_CEnd_C.Value = ZCFZ_M_B_CStart_C.Value + B_ZT_Child_Count;
            ZCFZ_C_B_CEnd_C.Value = ZCFZ_C_B_CStart_C.Value + B_ZT_Child_Count;
            LR_M_B_CEnd_C.Value = LR_M_B_CStart_C.Value + B_ZT_Child_Count;
            LR_C_B_CEnd_C.Value = LR_C_B_CStart_C.Value + B_ZT_Child_Count;

            decimal J_ZT_Child_Count = JCD_CEnd_C.Value - JCD_CStart_C.Value;
            ZCFZ_M_J_CEnd_C.Value = ZCFZ_M_J_CStart_C.Value + J_ZT_Child_Count;
            ZCFZ_C_J_CEnd_C.Value = ZCFZ_C_J_CStart_C.Value + J_ZT_Child_Count;
            LR_M_J_CEnd_C.Value = LR_M_J_CStart_C.Value + J_ZT_Child_Count;
            LR_C_J_CEnd_C.Value = LR_C_J_CStart_C.Value + J_ZT_Child_Count;
        }

        private void SaveSetting_Click(object sender, EventArgs e)
        {

        }

        private void Temp_BaseRep_ZTInfo_ColRow_indexChanged(object sender, EventArgs e)
        {
            this.Temp_BaseRep_ZTName_CellName.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTName_Col.Value) + this.Temp_BaseRep_ZTName_Row.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZBCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_ZBC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZJCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_ZJC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LBCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_LBC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LJCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_LJC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
        }

        private void Temp_BaseRep_DataInfo_ColRow_indexChanged(object sender, EventArgs e)
        {
            this.BZMCS.Text = IndexToASCiiTitle(this.BZMC.Value);
            this.BZMCCS.Text = IndexToASCiiTitle(this.BZMCC.Value);

            this.BZQCS.Text = IndexToASCiiTitle(this.BZQC.Value);
            this.BZQCCS.Text = IndexToASCiiTitle(this.BZQCC.Value);

            this.JZMCS.Text = IndexToASCiiTitle(this.JZMC.Value);
            this.JZMCCS.Text = IndexToASCiiTitle(this.JZMCC.Value);

            this.JZQCS.Text = IndexToASCiiTitle(this.JZQC.Value);
            this.JZQCCS.Text = IndexToASCiiTitle(this.JZQCC.Value);

            this.BLNCS.Text = IndexToASCiiTitle(this.BLNC.Value);
            this.BLNCCS.Text = IndexToASCiiTitle(this.BLNCC.Value);

            this.BLYCS.Text = IndexToASCiiTitle(this.BLYC.Value);
            this.BLYCCS.Text = IndexToASCiiTitle(this.BLYCC.Value);

            this.JLNCS.Text = IndexToASCiiTitle(this.JLNC.Value);
            this.JLNCCS.Text = IndexToASCiiTitle(this.JLNCC.Value);

            this.JLYCS.Text = IndexToASCiiTitle(this.JLYC.Value);
            this.JLYCCS.Text = IndexToASCiiTitle(this.JLYCC.Value);


        }

        string IndexToASCiiTitle(decimal Index)
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

        private void ReadBaseRepSet_Click(object sender, EventArgs e)
        {
            try
            {
                string ak = GetXMLValueByRC(BaseRepXML, this.Temp_BaseRep_ZTName_Row, Temp_BaseRep_ZTName_Col);
                ZTInfoBox.Text = "账套信息:" + ak;
                ZTDataBox.Text = "数据列:" + ak;

                Temp_BaseRep_ZTNo_ZBCV.Text = GetXMLValueByRC(BaseRepXML, Temp_BaseRep_ZTNo_R, Temp_BaseRep_ZTNo_ZBC); //.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_ZJCV.Text = GetXMLValueByRC(BaseRepXML, Temp_BaseRep_ZTNo_R, Temp_BaseRep_ZTNo_ZJC); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_LBCV.Text = GetXMLValueByRC(BaseRepXML, Temp_BaseRep_ZTNo_R, Temp_BaseRep_ZTNo_LBC); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_LJCV.Text = GetXMLValueByRC(BaseRepXML, Temp_BaseRep_ZTNo_R, Temp_BaseRep_ZTNo_LJC); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                decimal[] bzcol = new decimal[8];
                bzcol[0] = this.BZMC.Value;
                bzcol[1] = this.BZQC.Value;
                bzcol[2] = this.BZMCC.Value;
                bzcol[3] = this.BZQCC.Value;
                bzcol[4] = this.BLNC.Value;
                bzcol[5] = this.BLYC.Value;
                bzcol[6] = this.BLNCC.Value;
                bzcol[7] = this.BLYCC.Value;

                getztname(BaseRepXML, bzcol, this.Temp_BaseRep_BZTNo, Temp_BaseRep_BZTNo_New, Temp_BaseRep_BZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

                decimal[] jdcol = new decimal[8];
                jdcol[0] = this.JZMC.Value;
                jdcol[1] = this.JZQC.Value;
                jdcol[2] = this.JZMCC.Value;
                jdcol[3] = this.JZQCC.Value;
                jdcol[4] = this.JLNC.Value;
                jdcol[5] = this.JLYC.Value;
                jdcol[6] = this.JLNCC.Value;
                jdcol[7] = this.JLYCC.Value;

                getztname(BaseRepXML, jdcol, this.Temp_BaseRep_JZTNo, Temp_BaseRep_JZTNo_New, Temp_BaseRep_JZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        void getztname(RepXML TempXDRoot, decimal[] _Colno, TextBox v1, TextBox v2, TextBox v3, TextBox nmsg)
        {
            try
            {
                int RowCount = int.Parse(TempXDRoot.RootNode.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value);
                List<string> ztname = new List<string>();
                foreach (decimal ColNo in _Colno)
                {
                    for (int i = 0; i <= RowCount; i++)
                    {
                        string Cellpath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", i, ColNo);
                        if (TempXDRoot.RootNode.SelectSingleNode(Cellpath).Attributes["Type"].Value == "1")
                        {
                            if (TempXDRoot.RootNode.SelectSingleNode(Cellpath).Attributes["IsFormula"].Value == "1")
                            {
                                string FormulaText = TempXDRoot.RootNode.SelectSingleNode(Cellpath).Attributes["FormulaText"].Value;
                                string pattern = @"[^(]~\d{6}~";
                                foreach (Match match in Regex.Matches(FormulaText, pattern))
                                {
                                    if (ztname.IndexOf(match.Value.Replace("~", "").Replace(",", "")) < 0)
                                    {
                                        ztname.Add(match.Value.Replace("~", "").Replace(",", ""));
                                    }
                                }
                            }
                        }
                    }
                }
                if (ztname.Count > 0)
                {
                    v1.Text = ztname[0];
                    v2.Text = ztname[0];
                    v3.Text = ztname[0];
                }
                if (ztname.Count <= 0)
                {
                    nmsg.Text += "未匹配到有效的账套编码；";
                }
                else if (ztname.Count > 1)
                {
                    nmsg.Text += "匹配到多个账套编码；";
                    v3.Text = "";
                    foreach (string na in ztname)
                    {
                        v3.Text += "," + na;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        //冲抵
        private void Temp_CDWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            BCD_Basic_C_Cell.Text = IndexToASCiiTitle(this.BCD_Basic_C.Value);
            BCD_CCount_C_Cell.Text = IndexToASCiiTitle(this.BCD_CCount_C.Value);
            BCD_CStart_C_Cell.Text = IndexToASCiiTitle(this.BCD_CStart_C.Value);
            BCD_CEnd_C_Cell.Text = IndexToASCiiTitle(this.BCD_CEnd_C.Value);

            JCD_Basic_C_Cell.Text = IndexToASCiiTitle(this.JCD_Basic_C.Value);
            JCD_CCount_C_Cell.Text = IndexToASCiiTitle(this.JCD_CCount_C.Value);
            JCD_CStart_C_Cell.Text = IndexToASCiiTitle(this.JCD_CStart_C.Value);
            JCD_CEnd_C_Cell.Text = IndexToASCiiTitle(this.JCD_CEnd_C.Value);

            setZTCount();
        }

        private void BTN_ReadCDValue_Click(object sender, EventArgs e)
        {
            try
            {
                BCD_Basic_C_No.Text = GetXMLValueByRC(CDBXML, BCD_NO_R, BCD_Basic_C); 
                BCD_CStart_C_No.Text = GetXMLValueByRC(CDBXML, BCD_NO_R, BCD_CStart_C); 
                BCD_CEnd_C_No.Text = GetXMLValueByRC(CDBXML, BCD_NO_R, BCD_CEnd_C); 

                BCD_Basic_C_Name.Text = GetXMLValueByRC(CDBXML, BCD_Name_R, BCD_Basic_C); 
                BCD_CStart_C_Name.Text = GetXMLValueByRC(CDBXML, BCD_Name_R, BCD_CStart_C); 
                BCD_CEnd_C_Name.Text = GetXMLValueByRC(CDBXML, BCD_Name_R, BCD_CEnd_C); 
                
                BCD_CCount_C_F1.Text = GetXMLValueByRC(CDBXML, BCD_CCount_R1, BCD_CCount_C, "snafv"); 
                BCD_CCount_C_F2.Text = GetXMLValueByRC(CDBXML, BCD_CCount_R2, BCD_CCount_C, "snafv"); 
            }
            catch
            {

            }

            try
            {
                JCD_Basic_C_No.Text = GetXMLValueByRC(CDJXML, JCD_NO_R, JCD_Basic_C); 
                JCD_CStart_C_No.Text = GetXMLValueByRC(CDJXML, JCD_NO_R, JCD_CStart_C); 
                JCD_CEnd_C_No.Text = GetXMLValueByRC(CDJXML, JCD_NO_R, JCD_CEnd_C); 

                JCD_Basic_C_Name.Text = GetXMLValueByRC(CDJXML, JCD_Name_R, JCD_Basic_C); 
                JCD_CStart_C_Name.Text = GetXMLValueByRC(CDJXML, JCD_Name_R, JCD_CStart_C); 
                JCD_CEnd_C_Name.Text = GetXMLValueByRC(CDJXML, JCD_Name_R, JCD_CEnd_C); 

                JCD_CCount_C_F1.Text = GetXMLValueByRC(CDJXML, JCD_CCount_R1, JCD_CCount_C, "snafv");
                JCD_CCount_C_F2.Text = GetXMLValueByRC(CDJXML, JCD_CCount_R2, JCD_CCount_C, "snafv");
            }
            catch
            {

            }
        }

        //资产负债工作表
        private void Temp_ZCFZWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            ZCFZ_M_B_Z_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_B_Z_C.Value);
            ZCFZ_M_B_Count_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_B_Count_C.Value);
            ZCFZ_M_B_CStart_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_B_CStart_C.Value);
            ZCFZ_M_B_CEnd_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_B_CEnd_C.Value);

            ZCFZ_C_B_Z_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_B_Z_C.Value);
            ZCFZ_C_B_Count_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_B_Count_C.Value);
            ZCFZ_C_B_CStart_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_B_CStart_C.Value);
            ZCFZ_C_B_CEnd_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_B_CEnd_C.Value);
            //尽调
            ZCFZ_M_J_Z_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_J_Z_C.Value);
            ZCFZ_M_J_Count_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_J_Count_C.Value);
            ZCFZ_M_J_CStart_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_J_CStart_C.Value);
            ZCFZ_M_J_CEnd_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_M_J_CEnd_C.Value);

            ZCFZ_C_J_Z_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_J_Z_C.Value);
            ZCFZ_C_J_Count_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_J_Count_C.Value);
            ZCFZ_C_J_CStart_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_J_CStart_C.Value);
            ZCFZ_C_J_CEnd_C_Cell.Text = IndexToASCiiTitle(this.ZCFZ_C_J_CEnd_C.Value);
        }

        private void ReadZCFZValue_Click(object sender, EventArgs e)
        {
            try
            {
                ZCFZ_M_B_Z_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_M_B_Z_C); 
                ZCFZ_M_B_CStart_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_M_B_CStart_C); 
                ZCFZ_M_B_CEnd_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_M_B_CEnd_C); 

                ZCFZ_M_B_Z_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_M_B_Z_C); 
                ZCFZ_M_B_CStart_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_M_B_CStart_C); 
                ZCFZ_M_B_CEnd_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_M_B_CEnd_C); 

                ZCFZ_M_B_Count_C_F1.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_F1_R, ZCFZ_M_B_Count_C, "snafv");
                ZCFZ_M_B_Count_C_F2.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_F1_R, ZCFZ_M_B_Count_C, "snafv"); 

                //
                ZCFZ_C_B_Z_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_C_B_Z_C); 
                ZCFZ_C_B_CStart_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_C_B_CStart_C); 
                ZCFZ_C_B_CEnd_C_No.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_No_R, ZCFZ_C_B_CEnd_C); 

                ZCFZ_C_B_Z_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_C_B_Z_C); 
                ZCFZ_C_B_CStart_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_C_B_CStart_C); 
                ZCFZ_C_B_CEnd_C_Name.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_Name_R, ZCFZ_C_B_CEnd_C); 

                ZCFZ_C_B_Count_C_F1.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_F1_R, ZCFZ_C_B_Count_C, "snafv"); 
                ZCFZ_C_B_Count_C_F2.Text = GetXMLValueByRC(ZCFZBXML, ZCFZ_B_F2_R, ZCFZ_C_B_Count_C, "snafv"); 
            }
            catch
            {

            }
            try
            {
                ZCFZ_M_J_Z_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_M_J_Z_C); 
                ZCFZ_M_J_CStart_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_M_J_CStart_C); 
                ZCFZ_M_J_CEnd_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_M_J_CEnd_C); 

                ZCFZ_M_J_Z_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_M_J_Z_C); 
                ZCFZ_M_J_CStart_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_M_J_CStart_C);
                ZCFZ_M_J_CEnd_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_M_J_CEnd_C);

                ZCFZ_M_J_Count_C_F1.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_F1_R, ZCFZ_M_J_Count_C,"snafv"); 
                ZCFZ_M_J_Count_C_F2.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_F2_R, ZCFZ_M_J_Count_C, "snafv");
                //
                ZCFZ_C_J_Z_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_C_J_Z_C);
                ZCFZ_C_J_CStart_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_C_J_CStart_C); 
                ZCFZ_C_J_CEnd_C_No.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_No_R, ZCFZ_C_J_CEnd_C); 

                ZCFZ_C_J_Z_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_C_J_Z_C); 
                ZCFZ_C_J_CStart_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_C_J_CStart_C);
                ZCFZ_C_J_CEnd_C_Name.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_J_Name_R, ZCFZ_C_J_CEnd_C); 

                ZCFZ_C_J_Count_C_F1.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_B_F1_R, ZCFZ_C_J_Count_C, "snafv"); 
                ZCFZ_C_J_Count_C_F2.Text = GetXMLValueByRC(ZCFZJXML, ZCFZ_B_F2_R, ZCFZ_C_J_Count_C, "snafv");
            }
            catch
            {

            }
        }

        //利润表

        private void Temp_LRWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            LR_M_B_Z_C_Cell.Text = IndexToASCiiTitle(this.LR_M_B_Z_C.Value);
            LR_M_B_Count_C_Cell.Text = IndexToASCiiTitle(this.LR_M_B_Count_C.Value);
            LR_M_B_CStart_C_Cell.Text = IndexToASCiiTitle(this.LR_M_B_CStart_C.Value);
            LR_M_B_CEnd_C_Cell.Text = IndexToASCiiTitle(this.LR_M_B_CEnd_C.Value);

            LR_C_B_Z_C_Cell.Text = IndexToASCiiTitle(this.LR_C_B_Z_C.Value);
            LR_C_B_Count_C_Cell.Text = IndexToASCiiTitle(this.LR_C_B_Count_C.Value);
            LR_C_B_CStart_C_Cell.Text = IndexToASCiiTitle(this.LR_C_B_CStart_C.Value);
            LR_C_B_CEnd_C_Cell.Text = IndexToASCiiTitle(this.LR_C_B_CEnd_C.Value);
            //尽调
            LR_M_J_Z_C_Cell.Text = IndexToASCiiTitle(this.LR_M_J_Z_C.Value);
            LR_M_J_Count_C_Cell.Text = IndexToASCiiTitle(this.LR_M_J_Count_C.Value);
            LR_M_J_CStart_C_Cell.Text = IndexToASCiiTitle(this.LR_M_J_CStart_C.Value);
            LR_M_J_CEnd_C_Cell.Text = IndexToASCiiTitle(this.LR_M_J_CEnd_C.Value);

            LR_C_J_Z_C_Cell.Text = IndexToASCiiTitle(this.LR_C_J_Z_C.Value);
            LR_C_J_Count_C_Cell.Text = IndexToASCiiTitle(this.LR_C_J_Count_C.Value);
            LR_C_J_CStart_C_Cell.Text = IndexToASCiiTitle(this.LR_C_J_CStart_C.Value);
            LR_C_J_CEnd_C_Cell.Text = IndexToASCiiTitle(this.LR_C_J_CEnd_C.Value);
        }
        private void ReadLRValue_Click(object sender, EventArgs e)
        {
            try
            {
                LR_M_B_Z_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_M_B_Z_C); 
                LR_M_B_CStart_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_M_B_CStart_C); 
                LR_M_B_CEnd_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_M_B_CEnd_C); 

                LR_M_B_Z_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_M_B_Z_C); 
                LR_M_B_CStart_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_M_B_CStart_C); 
                LR_M_B_CEnd_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_M_B_CEnd_C); 

                LR_M_B_Count_C_F1.Text = GetXMLValueByRC(LRBXML, LR_B_F1_R, LR_M_B_Count_C, "snafv"); 
                LR_M_B_Count_C_F2.Text = GetXMLValueByRC(LRBXML, LR_B_F2_R, LR_M_B_Count_C, "snafv"); 
                //
                LR_C_B_Z_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_C_B_Z_C); 
                LR_C_B_CStart_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_C_B_CStart_C); 
                LR_C_B_CEnd_C_No.Text = GetXMLValueByRC(LRBXML, LR_B_No_R, LR_C_B_CEnd_C); 

                LR_C_B_Z_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_C_B_Z_C); 
                LR_C_B_CStart_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_C_B_CStart_C); 
                LR_C_B_CEnd_C_Name.Text = GetXMLValueByRC(LRBXML, LR_B_Name_R, LR_C_B_CEnd_C); 

                LR_C_B_Count_C_F1.Text = GetXMLValueByRC(LRBXML, LR_B_F1_R, LR_C_B_Count_C, "snafv"); 
                LR_C_B_Count_C_F2.Text = GetXMLValueByRC(LRBXML, LR_B_F2_R, LR_C_B_Count_C, "snafv"); 
            }
            catch
            {

            }

            try
            {
                LR_M_J_Z_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_M_J_Z_C); 
                LR_M_J_CStart_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_M_J_CStart_C); 
                LR_M_J_CEnd_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_M_J_CEnd_C); 

                LR_M_J_Z_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_Name_R, LR_M_J_Z_C); 
                LR_M_J_CStart_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_M_J_CStart_C);
                LR_M_J_CEnd_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_M_J_CEnd_C); 

                LR_M_J_Count_C_F1.Text = GetXMLValueByRC(LRJXML, LR_J_F1_R, LR_M_J_Count_C, "snafv"); 
                LR_M_J_Count_C_F2.Text = GetXMLValueByRC(LRJXML, LR_J_F2_R, LR_M_J_Count_C, "snafv");

                LR_C_J_Z_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_C_J_Z_C); 
                LR_C_J_CStart_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_C_J_CStart_C); 
                LR_C_J_CEnd_C_No.Text = GetXMLValueByRC(LRJXML, LR_J_No_R, LR_C_J_CEnd_C); 

                LR_C_J_Z_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_Name_R, LR_C_J_Z_C);
                LR_C_J_CStart_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_Name_R, LR_C_J_CStart_C); 
                LR_C_J_CEnd_C_Name.Text = GetXMLValueByRC(LRJXML, LR_J_Name_R, LR_C_J_CEnd_C); 

                LR_C_J_Count_C_F1.Text = GetXMLValueByRC(LRJXML, LR_B_F1_R, LR_C_J_Count_C, "snafv"); 
                LR_C_J_Count_C_F2.Text = GetXMLValueByRC(LRJXML, LR_B_F2_R, LR_C_J_Count_C, "snafv"); 
            }
            catch
            {

            }
        }

        private void WriteZTListInfo_Click(object sender, EventArgs e)
        {
            zDT.Clear();

            for (decimal i = BCD_CStart_C.Value; i <= BCD_CEnd_C.Value; i++)
            {
                string bno = GetXMLValueByRC(CDBXML, this.BCD_NO_R.Value, i);

                string bname = GetXMLValueByRC(CDBXML, this.BCD_Name_R.Value, i); 

                DataRow NEWZ = zDT.NewRow();

                while (bno.Length < 6)
                {
                    bno = "0" + bno;
                }

                NEWZ["RepNo"]= BaseRepNoFront.Text + bname;
                foreach (DataRow rDTdr in rDT.Rows)
                {
                    if (rDTdr["Code"].ToString() == NEWZ["RepNo"].ToString())
                    {
                        NEWZ["RepName"] = rDTdr["Name"];
                        NEWZ["RepID"] = rDTdr["TemplateID"];
                    }
                }
                NEWZ["zName"] =bname;
                NEWZ["BzNo"] = bno;
                NEWZ["BzCDCol"] = i;
                NEWZ["BzZCFZCol1"] = ZCFZ_M_B_CStart_C.Value + i - BCD_CStart_C.Value;
                NEWZ["BzZCFZCol2"] = ZCFZ_C_B_CStart_C.Value + i - BCD_CStart_C.Value;
                NEWZ["BzLRCol1"] = LR_M_B_CStart_C.Value + i - BCD_CStart_C.Value;
                NEWZ["BzLRCol2"] = LR_C_B_CStart_C.Value + i - BCD_CStart_C.Value;

                zDT.Rows.Add(NEWZ);
            }

            for (decimal i = JCD_CStart_C.Value; i <= JCD_CEnd_C.Value; i++)
            {
                string bno = GetXMLValueByRC(CDJXML, this.JCD_NO_R.Value, i); 
                string bname = GetXMLValueByRC(CDJXML, this.JCD_Name_R.Value, i); 

                while (bno.Length < 6)
                {
                    bno = "0" + bno;
                }

                foreach (DataRow ZDTdr in zDT.Rows)
                {
                    if (ZDTdr["RepNo"].ToString() == BaseRepNoFront.Text + bname)
                    {
                        ZDTdr["JzNo"] = bno;
                        ZDTdr["JzCDCol"] = i;
                        ZDTdr["JzZCFZCol1"] = ZCFZ_M_J_CStart_C.Value + i - JCD_CStart_C.Value;
                        ZDTdr["JzZCFZCol2"] = ZCFZ_C_J_CStart_C.Value + i - JCD_CStart_C.Value;
                        ZDTdr["JzLRCol1"] = LR_M_J_CStart_C.Value + i - JCD_CStart_C.Value;
                        ZDTdr["JzLRCol2"] = LR_C_J_CStart_C.Value + i - JCD_CStart_C.Value;
                    }
                }
            }
            ZTListShow.DataSource = zDT;

            MainWork.SelectTab(WorkPage);
        }

        private void StartWorking_Click(object sender, EventArgs e)
        {
            BackgroundWorker wk = new BackgroundWorker();
            wk.WorkerReportsProgress = true;
            wk.DoWork += Main_DoWork;
        }
        private void Main_DoWork(object sender, DoWorkEventArgs e)
        {

        }
    }

    public class RepXML
    {
        public XmlDocument Doc;
        public XmlNode RootNode { get { return Doc.DocumentElement; } }
    }
}
