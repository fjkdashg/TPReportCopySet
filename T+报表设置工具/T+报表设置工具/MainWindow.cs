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

        XmlDocument BaseRepXML;
        XmlDocument CDBXML;
        XmlDocument CDJXML;
        XmlDocument ZCFZBXML;
        XmlDocument ZCFZJXML;
        XmlDocument LRBXML;
        XmlDocument LRJXML;

        XmlNode BaseRepXML_root;
        XmlNode CDBXML_root;
        XmlNode CDJXML_root;
        XmlNode ZCFZBXML_root;
        XmlNode ZCFZJXML_root;
        XmlNode LRBXML_root;
        XmlNode LRJXML_root;

        public MainWindow()
        {
            InitializeComponent();
            
            //GetWorkDB_Click(new object(), new EventArgs());
        }

        protected override void OnLoad(EventArgs e)
        {
            GetWorkDB.PerformClick();
            SaveSQLConn.PerformClick();
            ReadBaseRepSet.PerformClick();
            BTN_ReadCDValue.PerformClick();
            ReadZCFZValue.PerformClick();
            ReadLRValue.PerformClick();
        }

        private void GetWorkDB_Click(object sender, EventArgs e)
        {
            string ConnStr = String.Format("data source={0},{1};initial catalog=Master;user id={2};pwd={3}", SQL_Host.Text.Trim(),SQL_Port.Text.Trim(),SQL_User.Text.Trim(),SQL_Pwd.Text.Trim());
            Console.WriteLine(ConnStr);
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
                this.MainWorkTabC.Enabled = true;
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
                Temp_BaseRep_ZTInfo_ColRow_indexChanged(new object(),new EventArgs());
                Temp_BaseRep_DataInfo_ColRow_indexChanged(new object(), new EventArgs());
                Temp_CDWorkRep_ColRow_indexChanged(new object(), new EventArgs());
                Temp_ZCFZWorkRep_ColRow_indexChanged(new object(), new EventArgs());
                Temp_LRWorkRep_ColRow_indexChanged(new object(), new EventArgs());
            }
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
            this.Temp_BaseRep_ZTName_CellName.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTName_Col.Value)+ this.Temp_BaseRep_ZTName_Row.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZBCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_ZBC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZJCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_ZJC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LBCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_LBC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LJCS.Text = IndexToASCiiTitle(this.Temp_BaseRep_ZTNo_LJC.Value) + this.Temp_BaseRep_ZTNo_R.Value.ToString();
        }

        private void Temp_BaseRep_DataInfo_ColRow_indexChanged(object sender, EventArgs e)
        {
            this.BZMCS.Text = IndexToASCiiTitle(this.BZMC.Value) ;
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
            string sql = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_BaseRep_ID.Text);
            SqlCommand cmd = new SqlCommand(sql, MainSQL);
            SqlDataReader sdr = cmd.ExecuteReader();
            while (sdr.Read())
            {
                string RepXMLStr =(string) sdr.GetSqlString(0);
                RepXMLStr=RepXMLStr.Replace("=\"", " ='").Replace("\" ","' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                Console.WriteLine("++++++++++++++++++++++++++++++++++++++++++++");

                BaseRepXML= new XmlDocument();
                BaseRepXML.LoadXml(RepXMLStr);
                BaseRepXML_root = BaseRepXML.DocumentElement;

                //this.richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.Temp_BaseRep_ZTName_Row.Value, Temp_BaseRep_ZTName_Col.Value);

                    string ak = BaseRepXML_root.SelectSingleNode(namepath).InnerXml;
                    ZTInfoBox.Text = "账套信息:" + ak;
                    ZTDataBox.Text= "数据列:" + ak;
                    //Console.WriteLine(ak);

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.Temp_BaseRep_ZTNo_R.Value, Temp_BaseRep_ZTNo_ZBC.Value);
                    Temp_BaseRep_ZTNo_ZBCV.Text = BaseRepXML_root.SelectSingleNode(namepath).InnerXml;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.Temp_BaseRep_ZTNo_R.Value, Temp_BaseRep_ZTNo_ZJC.Value);
                    Temp_BaseRep_ZTNo_ZJCV.Text = BaseRepXML_root.SelectSingleNode(namepath).InnerXml;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.Temp_BaseRep_ZTNo_R.Value, Temp_BaseRep_ZTNo_LBC.Value);
                    Temp_BaseRep_ZTNo_LBCV.Text = BaseRepXML_root.SelectSingleNode(namepath).InnerXml;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.Temp_BaseRep_ZTNo_R.Value, Temp_BaseRep_ZTNo_LJC.Value);
                    Temp_BaseRep_ZTNo_LJCV.Text = BaseRepXML_root.SelectSingleNode(namepath).InnerXml;

                    decimal[] bzcol = new decimal[8];
                    bzcol[0] = this.BZMC.Value;
                    bzcol[1] = this.BZQC.Value;
                    bzcol[2] = this.BZMCC.Value;
                    bzcol[3] = this.BZQCC.Value;
                    bzcol[4] = this.BLNC.Value;
                    bzcol[5] = this.BLYC.Value;
                    bzcol[6] = this.BLNCC.Value;
                    bzcol[7] = this.BLYCC.Value;

                    getztname(BaseRepXML_root, bzcol, this.Temp_BaseRep_BZTNo, Temp_BaseRep_BZTNo_New, Temp_BaseRep_BZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

                    decimal[] jdcol = new decimal[8];
                    jdcol[0] = this.JZMC.Value;
                    jdcol[1] = this.JZQC.Value;
                    jdcol[2] = this.JZMCC.Value;
                    jdcol[3] = this.JZQCC.Value;
                    jdcol[4] = this.JLNC.Value;
                    jdcol[5] = this.JLYC.Value;
                    jdcol[6] = this.JLNCC.Value;
                    jdcol[7] = this.JLYCC.Value;

                    getztname(BaseRepXML_root, jdcol, this.Temp_BaseRep_JZTNo, Temp_BaseRep_JZTNo_New, Temp_BaseRep_JZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            sdr.Close();
            cmd.Dispose();
        }

        void getztname(XmlNode TempXDRoot, decimal[] _Colno, TextBox v1, TextBox v2, TextBox v3, TextBox nmsg)
        {
            try
            {
                int RowCount = int.Parse(TempXDRoot.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value);
                List<string> ztname = new List<string>();
                foreach (decimal ColNo in _Colno)
                {
                    for (int i = 0; i <= RowCount; i++)
                    {
                        string Cellpath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", i, ColNo);
                        if (TempXDRoot.SelectSingleNode(Cellpath).Attributes["Type"].Value == "1")
                        {
                            if (TempXDRoot.SelectSingleNode(Cellpath).Attributes["IsFormula"].Value == "1")
                            {
                                string FormulaText = TempXDRoot.SelectSingleNode(Cellpath).Attributes["FormulaText"].Value;
                                string pattern = @"[^(]~\d{6}~";
                                foreach (Match match in Regex.Matches(FormulaText, pattern))
                                {
                                    Console.WriteLine(match.Value);

                                    if (ztname.IndexOf(match.Value.Replace("~", "").Replace(",","")) < 0)
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
                    nmsg.Text +="未匹配到有效的账套编码；";
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
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
        }

        //冲抵
        private void Temp_CDWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            BCD_Basic_C_Cell.Text= IndexToASCiiTitle(this.BCD_Basic_C.Value);
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
            string sql = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_CDB_ID.Text);
            SqlCommand cmd = new SqlCommand(sql, MainSQL);
            SqlDataReader sdr = cmd.ExecuteReader();
            while (sdr.Read())
            {
                string RepXMLStr = (string)sdr.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                CDBXML = new XmlDocument();
                CDBXML.LoadXml(RepXMLStr);
                CDBXML_root= CDBXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_NO_R.Value, BCD_Basic_C.Value);
                    BCD_Basic_C_No.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_NO_R.Value, BCD_CStart_C.Value);
                    BCD_CStart_C_No.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_NO_R.Value, BCD_CEnd_C.Value);
                    BCD_CEnd_C_No.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;


                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_Name_R.Value, BCD_Basic_C.Value);
                    BCD_Basic_C_Name.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_Name_R.Value, BCD_CStart_C.Value);
                    BCD_CStart_C_Name.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_Name_R.Value, BCD_CEnd_C.Value);
                    BCD_CEnd_C_Name.Text = CDBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_CCount_R1.Value, BCD_CCount_C.Value);
                    BCD_CCount_C_F1.Text = CDBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.BCD_CCount_R2.Value, BCD_CCount_C.Value);
                    BCD_CCount_C_F2.Text = CDBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }
            }
            sdr.Close();
            cmd.Dispose();

            string sql2 = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_CDJ_ID.Text);
            SqlCommand cmd2 = new SqlCommand(sql2, MainSQL);
            SqlDataReader sdr2 = cmd2.ExecuteReader();
            while (sdr2.Read())
            {
                string RepXMLStr = (string)sdr2.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                CDJXML = new XmlDocument();
                CDJXML.LoadXml(RepXMLStr);
                CDJXML_root = CDJXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_NO_R.Value, JCD_Basic_C.Value);
                    JCD_Basic_C_No.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_NO_R.Value, JCD_CStart_C.Value);
                    JCD_CStart_C_No.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_NO_R.Value, JCD_CEnd_C.Value);
                    JCD_CEnd_C_No.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;


                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_Name_R.Value, JCD_Basic_C.Value);
                    JCD_Basic_C_Name.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_Name_R.Value, JCD_CStart_C.Value);
                    JCD_CStart_C_Name.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_Name_R.Value, JCD_CEnd_C.Value);
                    JCD_CEnd_C_Name.Text = CDJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_CCount_R1.Value, JCD_CCount_C.Value);
                    JCD_CCount_C_F1.Text = CDJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.JCD_CCount_R2.Value, JCD_CCount_C.Value);
                    JCD_CCount_C_F2.Text = CDJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }


            }
            sdr2.Close();
            cmd2.Dispose();


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
            string sql1 = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_ZCFZB_ID.Text);
            SqlCommand cmd1 = new SqlCommand(sql1, MainSQL);
            SqlDataReader sdr1 = cmd1.ExecuteReader();
            while (sdr1.Read())
            {
                string RepXMLStr = (string)sdr1.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                ZCFZBXML = new XmlDocument();
                ZCFZBXML.LoadXml(RepXMLStr);
                ZCFZBXML_root = ZCFZBXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_M_B_Z_C.Value);
                    ZCFZ_M_B_Z_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_M_B_CStart_C.Value);
                    ZCFZ_M_B_CStart_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_M_B_CEnd_C.Value);
                    ZCFZ_M_B_CEnd_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_M_B_Z_C.Value);
                    ZCFZ_M_B_Z_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_M_B_CStart_C.Value);
                    ZCFZ_M_B_CStart_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_M_B_CEnd_C.Value);
                    ZCFZ_M_B_CEnd_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F1_R.Value, ZCFZ_M_B_Count_C.Value);
                    ZCFZ_M_B_Count_C_F1.Text = ZCFZBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F2_R.Value, ZCFZ_M_B_Count_C.Value);
                    ZCFZ_M_B_Count_C_F2.Text = ZCFZBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    //
                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_C_B_Z_C.Value);
                    ZCFZ_C_B_Z_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_C_B_CStart_C.Value);
                    ZCFZ_C_B_CStart_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_No_R.Value, ZCFZ_C_B_CEnd_C.Value);
                    ZCFZ_C_B_CEnd_C_No.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_C_B_Z_C.Value);
                    ZCFZ_C_B_Z_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_C_B_CStart_C.Value);
                    ZCFZ_C_B_CStart_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_Name_R.Value, ZCFZ_C_B_CEnd_C.Value);
                    ZCFZ_C_B_CEnd_C_Name.Text = ZCFZBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F1_R.Value, ZCFZ_C_B_Count_C.Value);
                    ZCFZ_C_B_Count_C_F1.Text = ZCFZBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F2_R.Value, ZCFZ_C_B_Count_C.Value);
                    ZCFZ_C_B_Count_C_F2.Text = ZCFZBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }
            }
            sdr1.Close();
            cmd1.Dispose();

            //尽调

            string sql2 = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_ZCFZJ_ID.Text);
            SqlCommand cmd2 = new SqlCommand(sql2, MainSQL);
            SqlDataReader sdr2 = cmd2.ExecuteReader();
            while (sdr2.Read())
            {
                string RepXMLStr = (string)sdr2.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                ZCFZJXML = new XmlDocument();
                ZCFZJXML.LoadXml(RepXMLStr);
                ZCFZJXML_root = ZCFZJXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_M_J_Z_C.Value);
                    ZCFZ_M_J_Z_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_M_J_CStart_C.Value);
                    ZCFZ_M_J_CStart_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_M_J_CEnd_C.Value);
                    ZCFZ_M_J_CEnd_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_M_J_Z_C.Value);
                    ZCFZ_M_J_Z_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_M_J_CStart_C.Value);
                    ZCFZ_M_J_CStart_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_M_J_CEnd_C.Value);
                    ZCFZ_M_J_CEnd_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_F1_R.Value, ZCFZ_M_J_Count_C.Value);
                    ZCFZ_M_J_Count_C_F1.Text = ZCFZJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_F2_R.Value, ZCFZ_M_J_Count_C.Value);
                    ZCFZ_M_J_Count_C_F2.Text = ZCFZJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    //
                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_C_J_Z_C.Value);
                    ZCFZ_C_J_Z_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_C_J_CStart_C.Value);
                    ZCFZ_C_J_CStart_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_No_R.Value, ZCFZ_C_J_CEnd_C.Value);
                    ZCFZ_C_J_CEnd_C_No.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_C_J_Z_C.Value);
                    ZCFZ_C_J_Z_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_C_J_CStart_C.Value);
                    ZCFZ_C_J_CStart_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_J_Name_R.Value, ZCFZ_C_J_CEnd_C.Value);
                    ZCFZ_C_J_CEnd_C_Name.Text = ZCFZJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F1_R.Value, ZCFZ_C_J_Count_C.Value);
                    ZCFZ_C_J_Count_C_F1.Text = ZCFZJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.ZCFZ_B_F2_R.Value, ZCFZ_C_J_Count_C.Value);
                    ZCFZ_C_J_Count_C_F2.Text = ZCFZJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }
            }
            sdr2.Close();
            cmd2.Dispose();
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
            string sql1 = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_LRB_ID.Text);
            SqlCommand cmd1 = new SqlCommand(sql1, MainSQL);
            SqlDataReader sdr1 = cmd1.ExecuteReader();
            while (sdr1.Read())
            {
                string RepXMLStr = (string)sdr1.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                LRBXML = new XmlDocument();
                LRBXML.LoadXml(RepXMLStr);
                LRBXML_root = LRBXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_M_B_Z_C.Value);
                    LR_M_B_Z_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_M_B_CStart_C.Value);
                    LR_M_B_CStart_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_M_B_CEnd_C.Value);
                    LR_M_B_CEnd_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_M_B_Z_C.Value);
                    LR_M_B_Z_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_M_B_CStart_C.Value);
                    LR_M_B_CStart_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_M_B_CEnd_C.Value);
                    LR_M_B_CEnd_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F1_R.Value, LR_M_B_Count_C.Value);
                    LR_M_B_Count_C_F1.Text = LRBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F2_R.Value, LR_M_B_Count_C.Value);
                    LR_M_B_Count_C_F2.Text = LRBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    //
                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_C_B_Z_C.Value);
                    LR_C_B_Z_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_C_B_CStart_C.Value);
                    LR_C_B_CStart_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_No_R.Value, LR_C_B_CEnd_C.Value);
                    LR_C_B_CEnd_C_No.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_C_B_Z_C.Value);
                    LR_C_B_Z_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_C_B_CStart_C.Value);
                    LR_C_B_CStart_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_Name_R.Value, LR_C_B_CEnd_C.Value);
                    LR_C_B_CEnd_C_Name.Text = LRBXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F1_R.Value, LR_C_B_Count_C.Value);
                    LR_C_B_Count_C_F1.Text = LRBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F2_R.Value, LR_C_B_Count_C.Value);
                    LR_C_B_Count_C_F2.Text = LRBXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }
            }
            sdr1.Close();
            cmd1.Dispose();

            //尽调

            string sql2 = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", this.Temp_LRJ_ID.Text);
            SqlCommand cmd2 = new SqlCommand(sql2, MainSQL);
            SqlDataReader sdr2 = cmd2.ExecuteReader();
            while (sdr2.Read())
            {
                string RepXMLStr = (string)sdr2.GetSqlString(0);
                RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");

                LRJXML = new XmlDocument();
                LRJXML.LoadXml(RepXMLStr);
                LRJXML_root = LRJXML.DocumentElement;

                //richTextBox3.Text = TempXDRoot.InnerXml;
                try
                {
                    string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_M_J_Z_C.Value);
                    LR_M_J_Z_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_M_J_CStart_C.Value);
                    LR_M_J_CStart_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_M_J_CEnd_C.Value);
                    LR_M_J_CEnd_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_M_J_Z_C.Value);
                    LR_M_J_Z_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_M_J_CStart_C.Value);
                    LR_M_J_CStart_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_M_J_CEnd_C.Value);
                    LR_M_J_CEnd_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_F1_R.Value, LR_M_J_Count_C.Value);
                    LR_M_J_Count_C_F1.Text = LRJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_F2_R.Value, LR_M_J_Count_C.Value);
                    LR_M_J_Count_C_F2.Text = LRJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    //
                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_C_J_Z_C.Value);
                    LR_C_J_Z_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_C_J_CStart_C.Value);
                    LR_C_J_CStart_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_No_R.Value, LR_C_J_CEnd_C.Value);
                    LR_C_J_CEnd_C_No.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_C_J_Z_C.Value);
                    LR_C_J_Z_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_C_J_CStart_C.Value);
                    LR_C_J_CStart_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_J_Name_R.Value, LR_C_J_CEnd_C.Value);
                    LR_C_J_CEnd_C_Name.Text = LRJXML_root.SelectSingleNode(namepath).InnerText;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F1_R.Value, LR_C_J_Count_C.Value);
                    LR_C_J_Count_C_F1.Text = LRJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;

                    namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", this.LR_B_F2_R.Value, LR_C_J_Count_C.Value);
                    LR_C_J_Count_C_F2.Text = LRJXML_root.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                }
                catch
                {

                }
            }
            sdr2.Close();
            cmd2.Dispose();
        }
    }

    class ztinfos
    {
        public ztinfo B = new ztinfo();
        public ztinfo J = new ztinfo();
    }

    class ztinfo
    {
        public string no = "";
        public string name = "";
        public string BaseRepNo = "";
        public string BaseRepName = "";
        public string BaseRepID = "";

        public Boolean isMainCompany = false;

        public string CDColNo = "";

        public string ZCFZCol1 = "";
        public string ZCFZCol2 = "";

        public string LRCol1 = "";
        public string LRCol2 = "";
    }
}
