using Newtonsoft.Json;
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
        public SqlConnection MainSQL;

        public RepWork RW;

        RepXML ZTListRepXML;

        RepXML BaseRepXML;
        RepXML CDBXML;
        RepXML CDJXML;
        RepXML ZCFZBXML;
        RepXML ZCFZJXML;
        RepXML LRBXML;
        RepXML LRJXML;

        int ZTListRep_RowCount = 0;

        DataTable rDT = new DataTable();
        DataTable zDT = new DataTable();

        systemSetting ms;
        public MainWindow()
        {
            InitializeComponent();

            //GetWorkDB_Click(new object(), new EventArgs());
            initZDB();
        }

        protected override void OnLoad(EventArgs e)
        {
            ms = new systemSetting();
            this.SQL_Host.Text = ms.sqlurl;
            this.SQL_Port.Text = ms.sqlport;
            this.SQL_User.Text = ms.sqluser;
            this.SQL_Pwd.Text = ms.sqlpwd;
        }

        void initZDB()
        {
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

            this.SQL_WorkDB.SelectedValue = ms.workdb;

            if (!String.IsNullOrWhiteSpace(this.SQL_WorkDB.Text))
            {
                this.ReadSetting.PerformClick();
            }
        }

        private void ReadSetting_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(this.SQL_WorkDB.Text))
            {
                MainSQL.Close();
                string ConnStr = String.Format("data source={0},{1};initial catalog={2};user id={3};pwd={4}", SQL_Host.Text.Trim(), SQL_Port.Text.Trim(), this.SQL_WorkDB.Text.Trim(), SQL_User.Text.Trim(), SQL_Pwd.Text.Trim());
                //Console.WriteLine(ConnStr);
                MainSQL = new SqlConnection(ConnStr);
                MainSQL.Open();
                RW = new RepWork(MainSQL);
                Console.WriteLine("readSetting : "+RW.mainSQL.State);
                try
                {
                    SqlDataAdapter SDA = new SqlDataAdapter("select  * from   TUFO_ReportTemplateBasic where    isSysTemplate <> 1 and Creator <> 'admin' ", MainSQL);
                    SDA.Fill(rDT);

                    DataTable dblist0 = new DataTable();
                    DataTable dblist1 = new DataTable();
                    DataTable dblist2 = new DataTable();
                    DataTable dblist3 = new DataTable();
                    DataTable dblist4 = new DataTable();
                    DataTable dblist5 = new DataTable();
                    DataTable dblist6 = new DataTable();
                    DataTable dblist7 = new DataTable();
                    SDA.Fill(dblist0);
                    SDA.Fill(dblist1);
                    SDA.Fill(dblist2);
                    SDA.Fill(dblist3);
                    SDA.Fill(dblist4);
                    SDA.Fill(dblist5);
                    SDA.Fill(dblist6);
                    SDA.Fill(dblist7);

                    this.ZtListRep.DataSource = dblist0;
                    this.ZtListRepID.DataSource = this.ZtListRep.DataSource;
                    this.ZtListRep.ValueMember = "Code";
                    this.ZtListRepID.ValueMember = "TemplateID";
                    this.ZtListRepID.SelectedValue = ms.ZtListRepID;


                    this.BaseRep.DataSource = dblist1;
                    this.BaseRep_No.DataSource = this.BaseRep.DataSource;
                    this.BaseRep_ID.DataSource = this.BaseRep.DataSource;
                    this.BaseRep.ValueMember = "Name";
                    this.BaseRep_No.ValueMember = "Code";
                    this.BaseRep_ID.ValueMember = "TemplateID";
                    //this.BaseRep.SelectedValue = ms.BaseRep;


                    this.BCDRep.DataSource = dblist2;
                    this.BCDRep_No.DataSource = this.BCDRep.DataSource;
                    this.BCDRep_ID.DataSource = this.BCDRep.DataSource;
                    this.BCDRep.ValueMember = "Name";
                    this.BCDRep_No.ValueMember = "Code";
                    this.BCDRep_ID.ValueMember = "TemplateID";
                    //this.BCDRep.SelectedValue = ms.BCDRep;


                    this.JCDRep.DataSource = dblist3;
                    this.JCDRep_No.DataSource = this.JCDRep.DataSource;
                    this.JCDRep_ID.DataSource = this.JCDRep.DataSource;
                    this.JCDRep.ValueMember = "Name";
                    this.JCDRep_No.ValueMember = "Code";
                    this.JCDRep_ID.ValueMember = "TemplateID";
                    //this.JCDRep.SelectedValue = ms.JCDRep;

                    this.BZCFZRep.DataSource = dblist4;
                    this.BZCFZRep_No.DataSource = this.BZCFZRep.DataSource;
                    this.BZCFZRep_ID.DataSource = this.BZCFZRep.DataSource;
                    this.BZCFZRep.ValueMember = "Name";
                    this.BZCFZRep_No.ValueMember = "Code";
                    this.BZCFZRep_ID.ValueMember = "TemplateID";
                    //this.BZCFZRep.SelectedValue = ms.BZCFZRep;

                    this.JZCFZRep.DataSource = dblist5;
                    this.JZCFZRep_No.DataSource = this.JZCFZRep.DataSource;
                    this.JZCFZRep_ID.DataSource = this.JZCFZRep.DataSource;
                    this.JZCFZRep.ValueMember = "Name";
                    this.JZCFZRep_No.ValueMember = "Code";
                    this.JZCFZRep_ID.ValueMember = "TemplateID";
                    //this.JZCFZRep.SelectedValue = ms.JZCFZRep;

                    this.BLRRep.DataSource = dblist6;
                    this.BLRRep_No.DataSource = this.BLRRep.DataSource;
                    this.BLRRep_ID.DataSource = this.BLRRep.DataSource;
                    this.BLRRep.ValueMember = "Name";
                    this.BLRRep_No.ValueMember = "Code";
                    this.BLRRep_ID.ValueMember = "TemplateID";
                    //this.BLRRep.SelectedValue = ms.BLRRep;

                    this.JLRRep.DataSource = dblist7;
                    this.JLRRep_No.DataSource = this.JLRRep.DataSource;
                    this.JLRRep_ID.DataSource = this.JLRRep.DataSource;
                    this.JLRRep.ValueMember = "Name";
                    this.JLRRep_No.ValueMember = "Code";
                    this.JLRRep_ID.ValueMember = "TemplateID";
                    //this.JLRRep.SelectedValue = ms.JLRRep;

                    
                    ZTListRepXML = RW.GetRepXMLData(this.ZtListRepID.Text);
                    SetSetting();

                    BaseRepXML = RW.GetRepXMLData(this.BaseRep_ID.Text);

                    CDBXML = RW.GetRepXMLData(this.BCDRep_ID.Text);
                    CDJXML = RW.GetRepXMLData(this.JCDRep_ID.Text);

                    ZCFZBXML = RW.GetRepXMLData(this.BZCFZRep_ID.Text);
                    ZCFZJXML = RW.GetRepXMLData(this.JZCFZRep_ID.Text);

                    LRBXML = RW.GetRepXMLData(this.BLRRep_ID.Text);
                    LRJXML = RW.GetRepXMLData(this.JLRRep_ID.Text);
                }
                catch
                {

                }
                
            }
        }

        void SetSetting()
        {
            
            //this.ZtListRep.SelectedValue = ms.ZtListRep;
            this.BaseRep_ID.SelectedValue = ms.BaseRepID;
            this.BCDRep_ID.SelectedValue = ms.BCDRepID;
            this.JCDRep_ID.SelectedValue = ms.JCDRepID;
            this.BZCFZRep_ID.SelectedValue = ms.BZCFZRepID;
            this.JZCFZRep_ID.SelectedValue = ms.JZCFZRepID;
            this.BLRRep_ID.SelectedValue = ms.BLRRepID;
            this.JLRRep_ID.SelectedValue = ms.JLRRepID;
            
            this.BaseRep_ZTName_C.Value = ms.BaseRep_ZTName_C;
            this.BaseRep_ZTName_R.Value = ms.BaseRep_ZTName_R;

            this.BaseRep_ZTTitle_R.Value = ms.BaseRep_ZTTitle_R;
            this.BaseRep_ZTTitle_BZCFZ_C.Value = ms.BaseRep_ZTTitle_BZCFZ_C;
            this.BaseRep_ZTTitle_JZCFZ_C.Value = ms.BaseRep_ZTTitle_JZCFZ_C;
            this.BaseRep_ZTTitle_BLR_C.Value = ms.BaseRep_ZTTitle_BLR_C;
            this.BaseRep_ZTTitle_JLR_C.Value = ms.BaseRep_ZTTitle_JLR_C;

            this.BaseRep_BZCFZ_QM_C.Value = ms.BaseRep_BZCFZ_QM_C;
            this.BaseRep_BZCFZ_QC_C.Value = ms.BaseRep_BZCFZ_QC_C;
            this.BaseRep_BZCFZ_QM_CD_C.Value = ms.BaseRep_BZCFZ_QM_CD_C;
            this.BaseRep_BZCFZ_QC_CD_C.Value = ms.BaseRep_BZCFZ_QC_CD_C;

            this.BaseRep_JZCFZ_QM_C.Value = ms.BaseRep_JZCFZ_QM_C;
            this.BaseRep_JZCFZ_QC_C.Value = ms.BaseRep_JZCFZ_QC_C;
            this.BaseRep_JZCFZ_QM_CD_C.Value = ms.BaseRep_JZCFZ_QM_CD_C;
            this.BaseRep_JZCFZ_QC_CD_C.Value = ms.BaseRep_JZCFZ_QC_CD_C;

            this.BaseRep_BLR_LFS_C.Value = ms.BaseRep_BLR_LFS_C;
            this.BaseRep_BLR_FS_C.Value = ms.BaseRep_BLR_FS_C;
            this.BaseRep_BLR_LFS_CD_C.Value = ms.BaseRep_BLR_LFS_CD_C;
            this.BaseRep_BLR_FS_CD_C.Value = ms.BaseRep_BLR_FS_CD_C;

            this.BaseRep_JLR_LFS_C.Value = ms.BaseRep_JLR_LFS_C;
            this.BaseRep_JLR_FS_C.Value = ms.BaseRep_JLR_FS_C;
            this.BaseRep_JLR_LFS_CD_C.Value = ms.BaseRep_JLR_LFS_CD_C;
            this.BaseRep_JLR_FS_CD_C.Value = ms.BaseRep_JLR_FS_CD_C;

            this.BCDRep_ZTNo_R.Value = ms.BCDRep_ZTNo_R;
            this.BCDRep_ZTName_R.Value = ms.BCDRep_ZTName_R;
            this.BCDRep_MC_C.Value = ms.BCDRep_MC_C;
            this.BCDREP_CC_C.Value = ms.BCDREP_CC_C;
            this.BCDRep_F_R.Value = ms.BCDRep_F_R;
            this.BCDRep_F_C.Value = ms.BCDRep_F_C;

            this.JCDRep_ZTNo_R.Value = ms.JCDRep_ZTNo_R;
            this.JCDRep_ZTName_R.Value = ms.JCDRep_ZTName_R;
            this.JCDRep_MC_C.Value = ms.JCDRep_MC_C;
            this.JCDREP_CC_C.Value = ms.JCDREP_CC_C;
            this.JCDRep_F_R.Value = ms.JCDRep_F_R;
            this.JCDRep_F_C.Value = ms.JCDRep_F_C;

            this.BZCFZRep_ZTNo_R.Value = ms.BZCFZRep_ZTNo_R;
            this.BZCFZRep_ZTName_R.Value = ms.BZCFZRep_ZTName_R;
            this.BZCFZRep_QM_MC_C.Value = ms.BZCFZRep_QM_MC_C;
            this.BZCFZRep_QM_CC_C.Value = ms.BZCFZRep_QM_CC_C;
            this.BZCFZRep_F_R.Value = ms.BZCFZRep_F_R;
            this.BZCFZRep_QM_F_C.Value = ms.BZCFZRep_QM_F_C;
            this.BZCFZRep_QC_MC_C.Value = ms.BZCFZRep_QC_MC_C;
            this.BZCFZRep_QC_CC_C.Value = ms.BZCFZRep_QC_CC_C;
            this.BZCFZRep_QC_F_C.Value = ms.BZCFZRep_QC_F_C;

            this.JZCFZRep_ZTNo_R.Value = ms.JZCFZRep_ZTNo_R;
            this.JZCFZRep_ZTName_R.Value = ms.JZCFZRep_ZTName_R;
            this.JZCFZRep_QM_MC_C.Value = ms.JZCFZRep_QM_MC_C;
            this.JZCFZRep_QM_CC_C.Value = ms.JZCFZRep_QM_CC_C;
            this.JZCFZRep_F_R.Value = ms.JZCFZRep_F_R;
            this.JZCFZRep_QM_F_C.Value = ms.JZCFZRep_QM_F_C;
            this.JZCFZRep_QC_MC_C.Value = ms.JZCFZRep_QC_MC_C;
            this.JZCFZRep_QC_CC_C.Value = ms.JZCFZRep_QC_CC_C;
            this.JZCFZRep_QC_F_C.Value = ms.JZCFZRep_QC_F_C;

            this.BLRRep_ZTNo_R.Value = ms.BLRRep_ZTNo_R;
            this.BLRRep_ZTName_R.Value = ms.BLRRep_ZTName_R;
            this.BLRRep_F_R.Value = ms.BLRRep_F_R;
            this.BLRRep_LFS_MC_C.Value = ms.BLRRep_LFS_MC_C;
            this.BLRRep_LFS_CC_C.Value = ms.BLRRep_LFS_CC_C;
            this.BLRRep_LFS_F_C.Value = ms.BLRRep_LFS_F_C;
            this.BLRRep_FS_MC_C.Value = ms.BLRRep_FS_MC_C;
            this.BLRRep_FS_CC_C.Value = ms.BLRRep_FS_CC_C;
            this.BLRRep_FS_F_C.Value = ms.BLRRep_FS_F_C;

            this.JLRRep_ZTNo_R.Value = ms.JLRRep_ZTNo_R;
            this.JLRRep_ZTName_R.Value = ms.JLRRep_ZTName_R;
            this.JLRRep_F_R.Value = ms.JLRRep_F_R;
            this.JLRRep_LFS_MC_C.Value = ms.JLRRep_LFS_MC_C;
            this.JLRRep_LFS_CC_C.Value = ms.JLRRep_LFS_CC_C;
            this.JLRRep_LFS_F_C.Value = ms.JLRRep_LFS_F_C;
            this.JLRRep_FS_MC_C.Value = ms.JLRRep_FS_MC_C;
            this.JLRRep_FS_CC_C.Value = ms.JLRRep_FS_CC_C;
            this.JLRRep_FS_F_C.Value = ms.JLRRep_FS_F_C;

            ZTListRep_RowCount = -1;
            for (decimal i = 4; i < int.Parse(ZTListRepXML.RTM_RootNode.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value); i++)
            {
                if (!String.IsNullOrWhiteSpace(RW.GetXMLValueByRC(ZTListRepXML.RTM_RootNode, i, 3)))
                {
                    ZTListRep_RowCount++;
                    Console.WriteLine(ZTListRep_RowCount + "      " + RW.GetXMLValueByRC(ZTListRepXML.RTM_RootNode, i, 3));
                }
            }
            setZTCount(ZTListRep_RowCount);

            Temp_BaseRep_ZTInfo_ColRow_indexChanged(new object(), new EventArgs());
            Temp_BaseRep_DataInfo_ColRow_indexChanged(new object(), new EventArgs());
            Temp_CDWorkRep_ColRow_indexChanged(new object(), new EventArgs());
            Temp_ZCFZWorkRep_ColRow_indexChanged(new object(), new EventArgs());
            Temp_LRWorkRep_ColRow_indexChanged(new object(), new EventArgs());
        }

        private void ViewSetting_Click(object sender, EventArgs e)
        {
            this.WorkSetting.Enabled = true;

            try
            {
                string ak = RW.GetXMLValueByRC(BaseRepXML, this.BaseRep_ZTName_R, BaseRep_ZTName_C);
                ZTInfoBox.Text = "账套信息:" + ak;
                ZTDataBox.Text = "数据列:" + ak;

                Temp_BaseRep_ZTNo_ZBCV.Text = RW.GetXMLValueByRC(BaseRepXML, BaseRep_ZTTitle_R, BaseRep_ZTTitle_BZCFZ_C); //.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_ZJCV.Text = RW.GetXMLValueByRC(BaseRepXML, BaseRep_ZTTitle_R, BaseRep_ZTTitle_JZCFZ_C); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_LBCV.Text = RW.GetXMLValueByRC(BaseRepXML, BaseRep_ZTTitle_R, BaseRep_ZTTitle_BLR_C); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                Temp_BaseRep_ZTNo_LJCV.Text = RW.GetXMLValueByRC(BaseRepXML, BaseRep_ZTTitle_R, BaseRep_ZTTitle_JLR_C); //BaseRepXML.RootNode.SelectSingleNode(namepath).InnerXml;

                decimal[] bzcol = new decimal[8];
                bzcol[0] = this.BaseRep_BZCFZ_QM_C.Value;
                bzcol[1] = this.BaseRep_BZCFZ_QC_C.Value;
                bzcol[2] = this.BaseRep_BZCFZ_QM_CD_C.Value;
                bzcol[3] = this.BaseRep_BZCFZ_QC_CD_C.Value;
                bzcol[4] = this.BaseRep_BLR_LFS_C.Value;
                bzcol[5] = this.BaseRep_BLR_FS_C.Value;
                bzcol[6] = this.BaseRep_BLR_LFS_CD_C.Value;
                bzcol[7] = this.BaseRep_BLR_FS_CD_C.Value;

                SetZTName(BaseRepXML, bzcol, this.Temp_BaseRep_BZTNo, Temp_BaseRep_BZTNo_New, Temp_BaseRep_BZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

                decimal[] jdcol = new decimal[8];
                jdcol[0] = this.BaseRep_JZCFZ_QM_C.Value;
                jdcol[1] = this.BaseRep_JZCFZ_QC_C.Value;
                jdcol[2] = this.BaseRep_JZCFZ_QM_CD_C.Value;
                jdcol[3] = this.BaseRep_JZCFZ_QC_CD_C.Value;
                jdcol[4] = this.BaseRep_JLR_LFS_C.Value;
                jdcol[5] = this.BaseRep_JLR_FS_C.Value;
                jdcol[6] = this.BaseRep_JLR_LFS_CD_C.Value;
                jdcol[7] = this.BaseRep_JLR_FS_CD_C.Value;

                SetZTName(BaseRepXML, jdcol, this.Temp_BaseRep_JZTNo, Temp_BaseRep_JZTNo_New, Temp_BaseRep_JZTNo_ListShow, Temp_BaseRep_ZTNoMsg);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            try
            {
                BCD_Basic_C_No.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTNo_R, BCDRep_MC_C);
                BCD_CStart_C_No.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTNo_R, BCDREP_CC_C);
                BCD_CEnd_C_No.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTNo_R, BCD_CEnd_C);

                BCD_Basic_C_Name.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTName_R, BCDRep_MC_C);
                BCD_CStart_C_Name.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTName_R, BCDREP_CC_C);
                BCD_CEnd_C_Name.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_ZTName_R, BCD_CEnd_C);

                BCD_CCount_C_F1.Text = RW.GetXMLValueByRC(CDBXML, BCDRep_F_R, BCDRep_F_C, "snafv");
                BCD_CCount_C_F2.Text = RW.GetXMLValueByRC(CDBXML, BCD_CCount_R2, BCDRep_F_C, "snafv");
            }
            catch
            {

            }

            try
            {
                JCD_Basic_C_No.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTNo_R, JCDRep_MC_C);
                JCD_CStart_C_No.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTNo_R, JCDREP_CC_C);
                JCD_CEnd_C_No.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTNo_R, JCD_CEnd_C);

                JCD_Basic_C_Name.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTName_R, JCDRep_MC_C);
                JCD_CStart_C_Name.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTName_R, JCDREP_CC_C);
                JCD_CEnd_C_Name.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_ZTName_R, JCD_CEnd_C);

                JCD_CCount_C_F1.Text = RW.GetXMLValueByRC(CDJXML, JCDRep_F_R, JCDRep_F_C, "snafv");
                JCD_CCount_C_F2.Text = RW.GetXMLValueByRC(CDJXML, JCD_CCount_R2, JCDRep_F_C, "snafv");
            }
            catch
            {

            }

            try
            {
                ZCFZ_M_B_Z_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, BZCFZRep_QM_MC_C);
                ZCFZ_M_B_CStart_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, BZCFZRep_QM_CC_C);
                ZCFZ_M_B_CEnd_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, ZCFZ_M_B_CEnd_C);

                ZCFZ_M_B_Z_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, BZCFZRep_QM_MC_C);
                ZCFZ_M_B_CStart_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, BZCFZRep_QM_CC_C);
                ZCFZ_M_B_CEnd_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, ZCFZ_M_B_CEnd_C);

                ZCFZ_M_B_Count_C_F1.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_F_R, BZCFZRep_QM_F_C, "snafv");
                ZCFZ_M_B_Count_C_F2.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_F_R, BZCFZRep_QM_F_C, "snafv");

                //
                ZCFZ_C_B_Z_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, BZCFZRep_QC_MC_C);
                ZCFZ_C_B_CStart_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, BZCFZRep_QC_CC_C);
                ZCFZ_C_B_CEnd_C_No.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTNo_R, ZCFZ_C_B_CEnd_C);

                ZCFZ_C_B_Z_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, BZCFZRep_QC_MC_C);
                ZCFZ_C_B_CStart_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, BZCFZRep_QC_CC_C);
                ZCFZ_C_B_CEnd_C_Name.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_ZTName_R, ZCFZ_C_B_CEnd_C);

                ZCFZ_C_B_Count_C_F1.Text = RW.GetXMLValueByRC(ZCFZBXML, BZCFZRep_F_R, BZCFZRep_QC_F_C, "snafv");
                ZCFZ_C_B_Count_C_F2.Text = RW.GetXMLValueByRC(ZCFZBXML, ZCFZ_B_F2_R, BZCFZRep_QC_F_C, "snafv");
            }
            catch
            {

            }
            try
            {
                ZCFZ_M_J_Z_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, JZCFZRep_QM_MC_C);
                ZCFZ_M_J_CStart_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, JZCFZRep_QM_CC_C);
                ZCFZ_M_J_CEnd_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, ZCFZ_M_J_CEnd_C);

                ZCFZ_M_J_Z_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, JZCFZRep_QM_MC_C);
                ZCFZ_M_J_CStart_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, JZCFZRep_QM_CC_C);
                ZCFZ_M_J_CEnd_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, ZCFZ_M_J_CEnd_C);

                ZCFZ_M_J_Count_C_F1.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_F_R, JZCFZRep_QM_F_C, "snafv");
                ZCFZ_M_J_Count_C_F2.Text = RW.GetXMLValueByRC(ZCFZJXML, ZCFZ_J_F2_R, JZCFZRep_QM_F_C, "snafv");
                //
                ZCFZ_C_J_Z_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, JZCFZRep_QC_MC_C);
                ZCFZ_C_J_CStart_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, JZCFZRep_QC_CC_C);
                ZCFZ_C_J_CEnd_C_No.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTNo_R, ZCFZ_C_J_CEnd_C);

                ZCFZ_C_J_Z_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, JZCFZRep_QC_MC_C);
                ZCFZ_C_J_CStart_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, JZCFZRep_QC_CC_C);
                ZCFZ_C_J_CEnd_C_Name.Text = RW.GetXMLValueByRC(ZCFZJXML, JZCFZRep_ZTName_R, ZCFZ_C_J_CEnd_C);

                ZCFZ_C_J_Count_C_F1.Text = RW.GetXMLValueByRC(ZCFZJXML, BZCFZRep_F_R, JZCFZRep_QC_F_C, "snafv");
                ZCFZ_C_J_Count_C_F2.Text = RW.GetXMLValueByRC(ZCFZJXML, ZCFZ_B_F2_R, JZCFZRep_QC_F_C, "snafv");
            }
            catch
            {

            }

            try
            {
                LR_M_B_Z_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, BLRRep_LFS_MC_C);
                LR_M_B_CStart_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, BLRRep_LFS_CC_C);
                LR_M_B_CEnd_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, LR_M_B_CEnd_C);

                LR_M_B_Z_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, BLRRep_LFS_MC_C);
                LR_M_B_CStart_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, BLRRep_LFS_CC_C);
                LR_M_B_CEnd_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, LR_M_B_CEnd_C);

                LR_M_B_Count_C_F1.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_F_R, BLRRep_LFS_F_C, "snafv");
                LR_M_B_Count_C_F2.Text = RW.GetXMLValueByRC(LRBXML, LR_B_F2_R, BLRRep_LFS_F_C, "snafv");
                //
                LR_C_B_Z_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, BLRRep_FS_MC_C);
                LR_C_B_CStart_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, BLRRep_FS_CC_C);
                LR_C_B_CEnd_C_No.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTNo_R, LR_C_B_CEnd_C);

                LR_C_B_Z_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, BLRRep_FS_MC_C);
                LR_C_B_CStart_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, BLRRep_FS_CC_C);
                LR_C_B_CEnd_C_Name.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_ZTName_R, LR_C_B_CEnd_C);

                LR_C_B_Count_C_F1.Text = RW.GetXMLValueByRC(LRBXML, BLRRep_F_R, BLRRep_FS_F_C, "snafv");
                LR_C_B_Count_C_F2.Text = RW.GetXMLValueByRC(LRBXML, LR_B_F2_R, BLRRep_FS_F_C, "snafv");
            }
            catch
            {

            }

            try
            {
                LR_M_J_Z_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, JLRRep_LFS_MC_C);
                LR_M_J_CStart_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, JLRRep_LFS_CC_C);
                LR_M_J_CEnd_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, LR_M_J_CEnd_C);

                LR_M_J_Z_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTName_R, JLRRep_LFS_MC_C);
                LR_M_J_CStart_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, JLRRep_LFS_CC_C);
                LR_M_J_CEnd_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, LR_M_J_CEnd_C);

                LR_M_J_Count_C_F1.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_F_R, JLRRep_LFS_F_C, "snafv");
                LR_M_J_Count_C_F2.Text = RW.GetXMLValueByRC(LRJXML, LR_J_F2_R, JLRRep_LFS_F_C, "snafv");

                LR_C_J_Z_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, JLRRep_FS_MC_C);
                LR_C_J_CStart_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, JLRRep_FS_CC_C);
                LR_C_J_CEnd_C_No.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTNo_R, LR_C_J_CEnd_C);

                LR_C_J_Z_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTName_R, JLRRep_FS_MC_C);
                LR_C_J_CStart_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTName_R, JLRRep_FS_CC_C);
                LR_C_J_CEnd_C_Name.Text = RW.GetXMLValueByRC(LRJXML, JLRRep_ZTName_R, LR_C_J_CEnd_C);

                LR_C_J_Count_C_F1.Text = RW.GetXMLValueByRC(LRJXML, BLRRep_F_R, JLRRep_FS_F_C, "snafv");
                LR_C_J_Count_C_F2.Text = RW.GetXMLValueByRC(LRJXML, LR_B_F2_R, JLRRep_FS_F_C, "snafv");
            }
            catch
            {

            }

            //WorkSetting.SelectTab(BaseSetting);
            //WorkSetting.SelectTab(CDSetting);
            //WorkSetting.SelectTab(ZCFZSetting);
            //WorkSetting.SelectTab(LRSetting);
        }
        
        void setZTCount(decimal ChildZTCount)
        {
            BCD_CEnd_C.Value = BCDREP_CC_C.Value + ChildZTCount;
            ZCFZ_M_B_CEnd_C.Value = BZCFZRep_QM_CC_C.Value + ChildZTCount;
            ZCFZ_C_B_CEnd_C.Value = BZCFZRep_QC_CC_C.Value + ChildZTCount;
            LR_M_B_CEnd_C.Value = BLRRep_LFS_CC_C.Value + ChildZTCount;
            LR_C_B_CEnd_C.Value = BLRRep_FS_CC_C.Value + ChildZTCount;

            JCD_CEnd_C.Value = JCDREP_CC_C.Value + ChildZTCount;
            ZCFZ_M_J_CEnd_C.Value = JZCFZRep_QM_CC_C.Value + ChildZTCount;
            ZCFZ_C_J_CEnd_C.Value = JZCFZRep_QC_CC_C.Value + ChildZTCount;
            LR_M_J_CEnd_C.Value = JLRRep_LFS_CC_C.Value + ChildZTCount;
            LR_C_J_CEnd_C.Value = JLRRep_FS_CC_C.Value + ChildZTCount;
        }


        private void Temp_BaseRep_ZTInfo_ColRow_indexChanged(object sender, EventArgs e)
        {
            this.Temp_BaseRep_ZTName_CellName.Text = RW.IndexToASCiiTitle(this.BaseRep_ZTName_C.Value) + this.BaseRep_ZTName_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZBCS.Text = RW.IndexToASCiiTitle(this.BaseRep_ZTTitle_BZCFZ_C.Value) + this.BaseRep_ZTTitle_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_ZJCS.Text = RW.IndexToASCiiTitle(this.BaseRep_ZTTitle_JZCFZ_C.Value) + this.BaseRep_ZTTitle_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LBCS.Text = RW.IndexToASCiiTitle(this.BaseRep_ZTTitle_BLR_C.Value) + this.BaseRep_ZTTitle_R.Value.ToString();
            this.Temp_BaseRep_ZTNo_LJCS.Text = RW.IndexToASCiiTitle(this.BaseRep_ZTTitle_JLR_C.Value) + this.BaseRep_ZTTitle_R.Value.ToString();
        }

        private void Temp_BaseRep_DataInfo_ColRow_indexChanged(object sender, EventArgs e)
        {
            this.BZMCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BZCFZ_QM_C.Value);
            this.BZMCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BZCFZ_QM_CD_C.Value);

            this.BZQCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BZCFZ_QC_C.Value);
            this.BZQCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BZCFZ_QC_CD_C.Value);

            this.JZMCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JZCFZ_QM_C.Value);
            this.JZMCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JZCFZ_QM_CD_C.Value);

            this.JZQCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JZCFZ_QC_C.Value);
            this.JZQCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JZCFZ_QC_CD_C.Value);

            this.BLNCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BLR_LFS_C.Value);
            this.BLNCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BLR_LFS_CD_C.Value);

            this.BLYCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BLR_FS_C.Value);
            this.BLYCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_BLR_FS_CD_C.Value);

            this.JLNCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JLR_LFS_C.Value);
            this.JLNCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JLR_LFS_CD_C.Value);

            this.JLYCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JLR_FS_C.Value);
            this.JLYCCS.Text = RW.IndexToASCiiTitle(this.BaseRep_JLR_FS_CD_C.Value);
        }
        //冲抵
        private void Temp_CDWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            BCD_Basic_C_Cell.Text = RW.IndexToASCiiTitle(this.BCDRep_MC_C.Value);
            BCD_CCount_C_Cell.Text = RW.IndexToASCiiTitle(this.BCDRep_F_C.Value);
            BCD_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.BCDREP_CC_C.Value);
            BCD_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.BCD_CEnd_C.Value);

            JCD_Basic_C_Cell.Text = RW.IndexToASCiiTitle(this.JCDRep_MC_C.Value);
            JCD_CCount_C_Cell.Text = RW.IndexToASCiiTitle(this.JCDRep_F_C.Value);
            JCD_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.JCDREP_CC_C.Value);
            JCD_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.JCD_CEnd_C.Value);

            setZTCount(ZTListRep_RowCount);
        }
        //资产负债工作表
        private void Temp_ZCFZWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            ZCFZ_M_B_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QM_MC_C.Value);
            ZCFZ_M_B_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QM_F_C.Value);
            ZCFZ_M_B_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QM_CC_C.Value);
            ZCFZ_M_B_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.ZCFZ_M_B_CEnd_C.Value);

            ZCFZ_C_B_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QC_MC_C.Value);
            ZCFZ_C_B_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QC_F_C.Value);
            ZCFZ_C_B_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.BZCFZRep_QC_CC_C.Value);
            ZCFZ_C_B_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.ZCFZ_C_B_CEnd_C.Value);
            //尽调
            ZCFZ_M_J_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QM_MC_C.Value);
            ZCFZ_M_J_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QM_F_C.Value);
            ZCFZ_M_J_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QM_CC_C.Value);
            ZCFZ_M_J_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.ZCFZ_M_J_CEnd_C.Value);

            ZCFZ_C_J_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QC_MC_C.Value);
            ZCFZ_C_J_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QC_F_C.Value);
            ZCFZ_C_J_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.JZCFZRep_QC_CC_C.Value);
            ZCFZ_C_J_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.ZCFZ_C_J_CEnd_C.Value);
        }
        //利润表

        private void Temp_LRWorkRep_ColRow_indexChanged(object sender, EventArgs e)
        {
            LR_M_B_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_LFS_MC_C.Value);
            LR_M_B_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_LFS_F_C.Value);
            LR_M_B_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_LFS_CC_C.Value);
            LR_M_B_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.LR_M_B_CEnd_C.Value);

            LR_C_B_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_FS_MC_C.Value);
            LR_C_B_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_FS_F_C.Value);
            LR_C_B_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.BLRRep_FS_CC_C.Value);
            LR_C_B_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.LR_C_B_CEnd_C.Value);
            //尽调
            LR_M_J_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_LFS_MC_C.Value);
            LR_M_J_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_LFS_F_C.Value);
            LR_M_J_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_LFS_CC_C.Value);
            LR_M_J_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.LR_M_J_CEnd_C.Value);

            LR_C_J_Z_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_FS_MC_C.Value);
            LR_C_J_Count_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_FS_F_C.Value);
            LR_C_J_CStart_C_Cell.Text = RW.IndexToASCiiTitle(this.JLRRep_FS_CC_C.Value);
            LR_C_J_CEnd_C_Cell.Text = RW.IndexToASCiiTitle(this.LR_C_J_CEnd_C.Value);
        }




        void SetZTName(RepXML TempXDRoot, decimal[] _Colno, TextBox v1, TextBox v2, TextBox v3, TextBox nmsg)
        {
            try
            {
                int RowCount = int.Parse(TempXDRoot.RTM_RootNode.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value);
                List<string> ztname = new List<string>();
                foreach (decimal ColNo in _Colno)
                {
                    for (int i = 0; i <= RowCount; i++)
                    {
                        string Cellpath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", i, ColNo);
                        if (TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["Type"].Value == "1")
                        {
                            if (TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["IsFormula"].Value == "1")
                            {
                                string FormulaText = TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["FormulaText"].Value;
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


        List<string> GetZTNo(RepXML TempXDRoot, decimal[] _Colno)
        {
            List<string> ztname = new List<string>();
            try
            {
                int RowCount = int.Parse(TempXDRoot.RTM_RootNode.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value);
                
                foreach (decimal ColNo in _Colno)
                {
                    for (int i = 0; i <= RowCount; i++)
                    {
                        string Cellpath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", i, ColNo);
                        if (TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["Type"].Value == "1")
                        {
                            if (TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["IsFormula"].Value == "1")
                            {
                                string FormulaText = TempXDRoot.RTM_RootNode.SelectSingleNode(Cellpath).Attributes["FormulaText"].Value;
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
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return ztname;

        }

        private void WriteZTListInfo_Click(object sender, EventArgs e)
        {
            zDT.Clear();



            for (decimal i = BCDREP_CC_C.Value; i <= BCD_CEnd_C.Value; i++)
            {
                string bno = RW.GetXMLValueByRC(CDBXML.RTM_RootNode, this.BCDRep_ZTNo_R.Value, i);

                string bname = RW.GetXMLValueByRC(CDBXML.RTM_RootNode, this.BCDRep_ZTName_R.Value, i); 

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
                NEWZ["BzZCFZCol1"] = BZCFZRep_QM_CC_C.Value + i - BCDREP_CC_C.Value;
                NEWZ["BzZCFZCol2"] = BZCFZRep_QC_CC_C.Value + i - BCDREP_CC_C.Value;
                NEWZ["BzLRCol1"] = BLRRep_LFS_CC_C.Value + i - BCDREP_CC_C.Value;
                NEWZ["BzLRCol2"] = BLRRep_FS_CC_C.Value + i - BCDREP_CC_C.Value;

                zDT.Rows.Add(NEWZ);
            }

            for (decimal i = JCDREP_CC_C.Value; i <= JCD_CEnd_C.Value; i++)
            {
                string bno = RW.GetXMLValueByRC(CDJXML.RTM_RootNode, this.JCDRep_ZTNo_R.Value, i); 
                string bname = RW.GetXMLValueByRC(CDJXML.RTM_RootNode, this.JCDRep_ZTName_R.Value, i); 

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
                        ZDTdr["JzZCFZCol1"] = JZCFZRep_QM_CC_C.Value + i - JCDREP_CC_C.Value;
                        ZDTdr["JzZCFZCol2"] = JZCFZRep_QC_CC_C.Value + i - JCDREP_CC_C.Value;
                        ZDTdr["JzLRCol1"] = JLRRep_LFS_CC_C.Value + i - JCDREP_CC_C.Value;
                        ZDTdr["JzLRCol2"] = JLRRep_FS_CC_C.Value + i - JCDREP_CC_C.Value;
                    }
                }
            }
            ZTListShow.DataSource = zDT;

            MainWork.SelectTab(WorkPage);
        }
        private void SaveSetting_Click(object sender, EventArgs e)
        {

        }
        private void StartWorking_Click(object sender, EventArgs e)
        {
            BackgroundWorker wk = new BackgroundWorker();
            wk.WorkerReportsProgress = true;
            wk.DoWork += Main_DoWork;
            wk.ProgressChanged += Main_ProgressWork;
            wk.RunWorkerAsync(ms);
        }
        private void Main_DoWork(object sender, DoWorkEventArgs e)
        {
            systemSetting ss=(systemSetting) e.Argument;
            BackgroundWorker BW = (BackgroundWorker)sender;


            WorkProgress WP = new WorkProgress();
            WP.StepCount = ZTListShow.Rows.Count * 23 + 2 + 10;
            WorkProgressMSG(BW, WP, "开始任务");

            RepXML _BaseRep_XML = RW.GetRepXMLData(ss.BaseRepID);
            WorkProgressMSG(BW, WP, "读取基础数据模板成功");


            List<string> _BaseRepZTNameList_B = GetZTNo(_BaseRep_XML,ss.BColFNoList);
            List<string> _BaseRepZTNameList_J = GetZTNo(_BaseRep_XML, ss.JColFNoList);

            WorkProgressMSG(BW, WP, "读取基础数据模板账编码成功");

            int BRowCount = int.Parse(_BaseRep_XML.RTM_RootNode.SelectSingleNode("Sheet/Total[1]").Attributes["rows"].Value);
            WP.StepCount = ZTListShow.Rows.Count * (ss.BColFNoList.Length+ ss.JColFNoList.Length)*((BRowCount*3)+5) +WP.step;
            WorkProgressMSG(BW, WP, "开始修订基础数据模板");
            int step = 2;
            foreach (DataGridViewRow dgvr in ZTListShow.Rows)
            {
                ChildInfo CI = new ChildInfo(dgvr);
                RepXML _WorkBaseRep_XML = RW.GetRepXMLData(ss.BaseRepID);
                WorkProgressMSG(BW, WP, "【" + CI.RepName + "】读取账套基础数据报表模板");

                RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, ss.BaseRep_ZTName_R, ss.BaseRep_ZTName_C, CI.RepName);
                RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, ss.BaseRep_ZTTitle_R, ss.BaseRep_ZTTitle_BZCFZ_C, "标准账套-"+CI.B_zNo);
                RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, ss.BaseRep_ZTTitle_R, ss.BaseRep_ZTTitle_JZCFZ_C, "尽调账套-" + CI.J_zNo);
                RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, ss.BaseRep_ZTTitle_R, ss.BaseRep_ZTTitle_BLR_C, "标准账套-" + CI.B_zNo);
                RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, ss.BaseRep_ZTTitle_R, ss.BaseRep_ZTTitle_JLR_C, "尽调账套-" + CI.J_zNo);
                WorkProgressMSG(BW, WP, "【" + CI.RepName + "】设置账套名称及编号");
                foreach (decimal fc in ss.BColFNoList)
                {
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  开始写入第"+fc.ToString()+"列公式");
                    for (decimal r = ss.BaseRep_ZTTitle_R + 1; r <= BRowCount; r++)
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  RTM开始写入" + fc.ToString() + "列"+r.ToString()+"行公式");
                        string RTMF = RW.GetXMLValueByRC(_BaseRep_XML.RTM_RootNode, r, fc, NodeValType.Attributes);
                        if (!string.IsNullOrWhiteSpace(RTMF))
                        {
                            foreach (string old in _BaseRepZTNameList_B)
                            {
                                RTMF = RTMF.Replace(old, CI.B_zNo);
                            }
                            RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, r, fc, RTMF, NodeValType.Attributes);
                            //WorkProgressMSG(BW, WP, "       --      完成");
                        }
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  RTF开始写入" + fc.ToString() + "列" + r.ToString() + "行公式");
                        string RTFF = RW.GetXMLValueByRC(_BaseRep_XML.RTF_RootNode, r, fc, NodeValType.Attributes, XMP_TempType.Formulas);
                        if (!string.IsNullOrWhiteSpace(RTFF))
                        {
                            //WorkProgressMSG(BW, WP, "       --      开始  "+RTFF);
                            foreach (string old in _BaseRepZTNameList_B)
                            {
                                RTFF = RTFF.Replace(old, CI.B_zNo);
                            }
                            RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTF_RootNode, r, fc, RTFF, NodeValType.Attributes, XMP_TempType.Formulas);
                            //WorkProgressMSG(BW, WP, "       --      完成");
                        }
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  " + fc.ToString() + "列" + r.ToString() + "行公式设置完成");
                    }
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  第" + fc.ToString() + "列公式写入完成");
                }
                WorkProgressMSG(BW, WP, "【" + CI.RepName + "】B  标准表公式设置完成");

                foreach (decimal fc in ss.JColFNoList)
                {
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】J  开始写入第" + fc.ToString() + "列公式");
                    for (decimal r = ss.BaseRep_ZTTitle_R + 5; r <= BRowCount; r++)
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】J  RTM开始写入" + fc.ToString() + "列" + r.ToString() + "行公式");
                        string RTM = RW.GetXMLValueByRC(_BaseRep_XML.RTM_RootNode, r, fc, NodeValType.Attributes);
                        if (!string.IsNullOrWhiteSpace(RTM))
                        {
                            foreach (string old in _BaseRepZTNameList_J)
                            {
                                RTM = RTM.Replace(old, CI.J_zNo);
                            }
                            RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTM_RootNode, r, fc, RTM, NodeValType.Attributes);
                        }
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】J  RTF开始写入" + fc.ToString() + "列" + r.ToString() + "行公式");
                        string RTF = RW.GetXMLValueByRC(_BaseRep_XML.RTF_RootNode, r, fc, NodeValType.Attributes, XMP_TempType.Formulas);
                        if (!string.IsNullOrWhiteSpace(RTF))
                        {
                            foreach (string old in _BaseRepZTNameList_J)
                            {
                                RTF = RTF.Replace(old, CI.J_zNo);
                            }
                            RW.UpdateXMLValueByRC(_WorkBaseRep_XML.RTF_RootNode, r, fc, RTF, NodeValType.Attributes, XMP_TempType.Formulas);
                        }
                    }
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】J  第" + fc.ToString() + "列公式写入完成");
                }
                WorkProgressMSG(BW, WP, "【" + CI.RepName + "】J  尽调表公式设置完成");

                string sql1 = String.Format(" DECLARE @ptrval binary(16) \n declare @len int; \n SELECT @ptrval = TEXTPTR({1}), @len = DATALENGTH({1}) / 2 \n  FROM {0} WHERE TemplateID ={2} \n UPDATETEXT {0}.{1} @ptrval 0 @len '{3}'", "TUFO_ReportTemplateModel", "TemplateStyle", CI.RepID, "<ReportTemplateModel>" + _WorkBaseRep_XML.RTM_RootNode.InnerXml.ToString().Replace("~", "\"") + "</ReportTemplateModel>");
                //string sql = String.Format("update TUFO_ReportTemplateModel set TemplateStyle='<ReportTemplateModel>" + _WorkBaseRep_XML.RTM_RootNode.InnerXml.ToString().Replace("~","\"")+ "</ReportTemplateModel>'  where TemplateID={0}", CI.RepID);
                Console.WriteLine(sql1);
                try
                {
                    SqlCommand cmd = new SqlCommand(sql1, MainSQL);
                    int uc = cmd.ExecuteNonQuery();
                    if (uc >= 0)
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Model成功");
                    }
                    else
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Model失败");
                    }
                    cmd.Dispose();
                }
                catch (Exception ex)
                {
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Model失败");
                    Console.WriteLine(ex.Message);
                }
                step += 1;

                string sql2 = String.Format(" DECLARE @ptrval binary(16) \n declare @len int; \n SELECT @ptrval = TEXTPTR({1}), @len = DATALENGTH({1}) / 2 \n  FROM {0} WHERE TemplateID ={2} \n UPDATETEXT {0}.{1} @ptrval 0 @len '{3}'", "TUFO_ReportTemplateFormulas", "FormulaText", CI.RepID, "<Formulas>" + _WorkBaseRep_XML.RTF_RootNode.InnerXml.ToString().Replace("~", "\"") + "</Formulas>");
                //string sql2 = String.Format("update TUFO_ReportTemplateFormulas set FormulaText='<Formulas>" + _WorkBaseRep_XML.RTF_RootNode.InnerXml.ToString().Replace("~", "\"") + "</Formulas>'  where TemplateID={0}", CI.RepID);
                Console.WriteLine(sql2);
                try
                {
                    SqlCommand cmd = new SqlCommand(sql2, MainSQL);
                    int uc = cmd.ExecuteNonQuery();
                    if (uc >= 0)
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Formulas成功");
                    }
                    else
                    {
                        WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Formulas失败");
                    }
                    cmd.Dispose();
                }
                catch (Exception ex)
                {
                    WorkProgressMSG(BW, WP, "【" + CI.RepName + "】写入Formulas失败");
                    Console.WriteLine(ex.Message);
                }
                step += 1;
            }

        }

        public void WorkProgressMSG(BackgroundWorker BW, WorkProgress WP,string msg,Boolean Stete=false)
        {
            WP.step++;
            WP.stepmsg = WP.step.ToString()+"/"+WP.StepCount.ToString()+ msg;
            BW.ReportProgress(WP.step, WP);
            if (Stete)
            {
                Console.Write(WP.stepmsg);
            }
            else
            {
                Console.WriteLine(WP.stepmsg);
            }
            
        }

        private void Main_ProgressWork(object sender, ProgressChangedEventArgs e)
        {
            this.progressBar1.Value = e.ProgressPercentage;
            WorkProgress WP = (WorkProgress)e.UserState;
            if (this.progressBar1.Maximum != WP.StepCount)
            {
                this.progressBar1.Maximum = WP.StepCount;
            }
            ProgressMSG.Text = WP.stepmsg;
        }

        private void ZCFZSetting_Click(object sender, EventArgs e)
        {

        }
    }

    public class RepXML
    {
        public XmlDocument RTM_Doc;
        public XmlNode RTM_RootNode { get { return RTM_Doc.DocumentElement; } }

        public XmlDocument RTF_Doc;
        public XmlNode RTF_RootNode { get { return RTF_Doc.DocumentElement; } }

    }



    public class WorkProgress
    {
        public int step = 0;
        public int StepCount = 100;
        public string stepmsg = "";
        public Boolean State = false;
    }

    public class ChildInfo
    {
        public string RepID;
        public string RepNo;
        public string RepName;

        public string B_zNo;
        public string B_CD_Col;
        public string B_ZCFZ_Col1;
        public string B_ZCFZ_Col2;
        public string B_LR_Col1;
        public string B_LR_Col2;

        public string J_zNo;
        public string J_CD_Col;
        public string J_ZCFZ_Col1;
        public string J_ZCFZ_Col2;
        public string J_LR_Col1;
        public string J_LR_Col2;

        public ChildInfo()
        { }

        public ChildInfo(DataGridViewRow sr)
        {
            RepID = sr.Cells["RepID"].Value.ToString();
            RepNo = sr.Cells["RepName"].Value.ToString();
            RepName = sr.Cells["zName"].Value.ToString();

            B_zNo = sr.Cells["BzNo"].Value.ToString();
            B_CD_Col = sr.Cells["BzCDCol"].Value.ToString();
            B_ZCFZ_Col1 = sr.Cells["BzZCFZCol1"].Value.ToString();
            B_ZCFZ_Col2 = sr.Cells["BzZCFZCol2"].Value.ToString();
            B_LR_Col1 = sr.Cells["BzLRCol1"].Value.ToString();
            B_LR_Col2 = sr.Cells["BzLRCol2"].Value.ToString();

            J_zNo = sr.Cells["JzNo"].Value.ToString();
            J_CD_Col = sr.Cells["JzCDCol"].Value.ToString();
            J_ZCFZ_Col1 = sr.Cells["JzZCFZCol1"].Value.ToString();
            J_ZCFZ_Col2 = sr.Cells["JzZCFZCol2"].Value.ToString();
            J_LR_Col1 = sr.Cells["JzLRCol1"].Value.ToString();
            J_LR_Col2 = sr.Cells["JzLRCol1"].Value.ToString();
        }
    }
    public enum NodeValType { innerXML,innerText, Attributes }
    public enum XMP_TempType { Model , Formulas }
    public class systemSetting
    {
        public string sqlurl="192.168.100.115";
        public string sqlport="1433";
        public string sqluser="sa";
        public string sqlpwd="Hc@3232327";

        public string workdb= "UFTData320976_999999";
        public string ZtListRepID = "234";


        public string BaseRepID = "183";
        public string BCDRepID = "200";
        public string JCDRepID = "205";
        public string BZCFZRepID = "202";
        public string JZCFZRepID = "207";
        public string BLRRepID = "201";
        public string JLRRepID = "206";

        public decimal BaseRep_ZTName_R=2;
        public decimal BaseRep_ZTName_C=1;
        public decimal BaseRep_ZTTitle_R=4;
        public decimal BaseRep_ZTTitle_BZCFZ_C=3;
        public decimal BaseRep_ZTTitle_JZCFZ_C=7;
        public decimal BaseRep_ZTTitle_BLR_C=14;
        public decimal BaseRep_ZTTitle_JLR_C=18;

        public decimal BaseRep_BZCFZ_QM_C=3;
        public decimal BaseRep_BZCFZ_QC_C=4;
        public decimal BaseRep_BZCFZ_QM_CD_C=5;
        public decimal BaseRep_BZCFZ_QC_CD_C=6;

        public decimal BaseRep_JZCFZ_QM_C=7;
        public decimal BaseRep_JZCFZ_QC_C=8;
        public decimal BaseRep_JZCFZ_QM_CD_C=9;
        public decimal BaseRep_JZCFZ_QC_CD_C=10;

        public decimal BaseRep_BLR_LFS_C=14;
        public decimal BaseRep_BLR_FS_C=15;
        public decimal BaseRep_BLR_LFS_CD_C=16;
        public decimal BaseRep_BLR_FS_CD_C=17;

        public decimal BaseRep_JLR_LFS_C=18;
        public decimal BaseRep_JLR_FS_C=19;
        public decimal BaseRep_JLR_LFS_CD_C=20;
        public decimal BaseRep_JLR_FS_CD_C=21;

        public decimal[] BColFNoList { get {
                decimal[] bzcol = new decimal[8];
                bzcol[0] = BaseRep_BZCFZ_QM_C;
                bzcol[1] = BaseRep_BZCFZ_QC_C;
                bzcol[2] = BaseRep_BZCFZ_QM_CD_C;
                bzcol[3] = BaseRep_BZCFZ_QC_CD_C;
                bzcol[4] = BaseRep_BLR_LFS_C;
                bzcol[5] = BaseRep_BLR_FS_C;
                bzcol[6] = BaseRep_BLR_LFS_CD_C;
                bzcol[7] = BaseRep_BLR_FS_CD_C;
                return bzcol;
            } }

        public decimal[] JColFNoList { get {
                decimal[] jdcol = new decimal[8];
                jdcol[0] = BaseRep_JZCFZ_QM_C;
                jdcol[1] = BaseRep_JZCFZ_QC_C;
                jdcol[2] = BaseRep_JZCFZ_QM_CD_C;
                jdcol[3] = BaseRep_JZCFZ_QC_CD_C;
                jdcol[4] = BaseRep_JLR_LFS_C;
                jdcol[5] = BaseRep_JLR_FS_C;
                jdcol[6] = BaseRep_JLR_LFS_CD_C;
                jdcol[7] = BaseRep_JLR_FS_CD_C;
                return jdcol;
            } }

        public decimal BCDRep_ZTNo_R=4;
        public decimal BCDRep_ZTName_R=5;
        public decimal BCDRep_MC_C=6;
        public decimal BCDREP_CC_C=10;
        public decimal BCDRep_F_R=6;
        public decimal BCDRep_F_C=8;

        public decimal JCDRep_ZTNo_R=4;
        public decimal JCDRep_ZTName_R=5;
        public decimal JCDRep_MC_C=6;
        public decimal JCDREP_CC_C=10;
        public decimal JCDRep_F_R=6;
        public decimal JCDRep_F_C=8;

        public decimal BZCFZRep_ZTNo_R=3;
        public decimal BZCFZRep_ZTName_R=4;
        public decimal BZCFZRep_QM_MC_C=16;
        public decimal BZCFZRep_QM_CC_C=20;
        public decimal BZCFZRep_F_R=7;
        public decimal BZCFZRep_QM_F_C=18;
        public decimal BZCFZRep_QC_MC_C=47;
        public decimal BZCFZRep_QC_CC_C=51;
        public decimal BZCFZRep_QC_F_C=49;

        public decimal JZCFZRep_ZTNo_R=3;
        public decimal JZCFZRep_ZTName_R=4;
        public decimal JZCFZRep_QM_MC_C=16;
        public decimal JZCFZRep_QM_CC_C=20;
        public decimal JZCFZRep_F_R=7;
        public decimal JZCFZRep_QM_F_C=18;
        public decimal JZCFZRep_QC_MC_C=47;
        public decimal JZCFZRep_QC_CC_C=51;
        public decimal JZCFZRep_QC_F_C=49;

        public decimal BLRRep_ZTNo_R=3;
        public decimal BLRRep_ZTName_R=4;
        public decimal BLRRep_F_R=7;
        public decimal BLRRep_LFS_MC_C=16;
        public decimal BLRRep_LFS_CC_C=20;
        public decimal BLRRep_LFS_F_C=18;
        public decimal BLRRep_FS_MC_C=47;
        public decimal BLRRep_FS_CC_C=51;
        public decimal BLRRep_FS_F_C=49;

        public decimal JLRRep_ZTNo_R=3;
        public decimal JLRRep_ZTName_R=4;
        public decimal JLRRep_F_R=7;
        public decimal JLRRep_LFS_MC_C=16;
        public decimal JLRRep_LFS_CC_C=20;
        public decimal JLRRep_LFS_F_C=18;
        public decimal JLRRep_FS_MC_C=47;
        public decimal JLRRep_FS_CC_C=51;
        public decimal JLRRep_FS_F_C=49;
    }
}
