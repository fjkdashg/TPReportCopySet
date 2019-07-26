using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace T_报表设置工具
{
    public class RepWork
    {
        public SqlConnection mainSQL;

        public RepWork(SqlConnection _MainSQL)
        {
            mainSQL = _MainSQL;
            //MainSQL.Open();
            Console.WriteLine(mainSQL.State);
        }

        public  RepXML GetRepXMLData(string RepID)
        {
            RepXML rx = new RepXML();

            string sql = String.Format("select TemplateStyle from TUFO_ReportTemplateModel where TemplateID={0}", RepID);
            //Console.WriteLine(sql);
            try
            {
                SqlCommand cmd = new SqlCommand(sql, mainSQL);
                SqlDataReader sdr = cmd.ExecuteReader();
                while (sdr.Read())
                {
                    string RepXMLStr = (string)sdr.GetSqlString(0);
                    RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");
                    //Console.WriteLine("GetRepXMLData:   " + RepXMLStr);
                    rx.RTM_Doc = new XmlDocument();
                    rx.RTM_Doc.LoadXml(RepXMLStr);
                }
                sdr.Close();
                cmd.Dispose();
                //Console.WriteLine(rx.RootNode.InnerXml.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("GetRepXMLData:   " + ex.Message);
            }

            string sql2 = String.Format("select FormulaText from TUFO_ReportTemplateFormulas where TemplateID={0}", RepID);
            //Console.WriteLine(sql);
            try
            {
                SqlCommand cmd = new SqlCommand(sql2, mainSQL);
                SqlDataReader sdr = cmd.ExecuteReader();
                while (sdr.Read())
                {
                    string RepXMLStr = (string)sdr.GetSqlString(0);
                    //RepXMLStr = RepXMLStr.Replace("=\"", " ='").Replace("\" ", "' ").Replace("类 ='", "类 =\"").Replace("\"", "~").Replace("'", "\"");
                    //Console.WriteLine("GetRepXMLData:   " + RepXMLStr);
                    rx.RTF_Doc = new XmlDocument();
                    rx.RTF_Doc.LoadXml(RepXMLStr);
                }
                sdr.Close();
                cmd.Dispose();
                //Console.WriteLine(rx.RootNode.InnerXml.ToString());
            }
            catch (Exception ex)
            {
                Console.WriteLine("GetRepXMLData:   " + ex.Message);
            }

            return rx;
        }

        public string GetXMLValueByRC(RepXML RX, NumericUpDown rowno, NumericUpDown colno, string VType = "snix")
        {
            string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", rowno.Value, colno.Value);
            string rStr = "";
            switch (VType)
            {
                case "snix":
                    rStr = RX.RTM_RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
                case "snafv":
                    rStr = RX.RTM_RootNode.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                    break;
                default:
                    rStr = RX.RTM_RootNode.SelectSingleNode(namepath).InnerXml;
                    break;
            }
            return rStr;
        }

        public string GetXMLValueByRC(XmlNode RN, decimal rowno, decimal colno, NodeValType VType = NodeValType.innerXML, XMP_TempType TP = XMP_TempType.Model)
        {
            string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", rowno, colno);
            if (TP == XMP_TempType.Formulas)
            {
                namepath = String.Format("Sheet/Formula/cell[@row='{0}' and @col='{1}']", rowno, colno);
            }
            //Console.WriteLine("GetXMLValueByRC: " + namepath);
            //Console.WriteLine("GetXMLValueByRC XML: " + RN.OwnerDocument.InnerXml);
            
            string rStr = "";
            try
            {
                switch (VType)
                {
                    case NodeValType.innerXML:
                        rStr = RN.SelectSingleNode(namepath).InnerXml;
                        break;
                    case NodeValType.Attributes:
                        rStr = RN.SelectSingleNode(namepath).Attributes["FormulaText"].Value;
                        break;
                    default:
                        rStr = RN.SelectSingleNode(namepath).InnerText;
                        break;
                }
            }
            catch
            {

            }
            
            return rStr;
        }


        public void UpdateXMLValueByRC(XmlNode RN, decimal rowno, decimal colno, string newvalue, NodeValType VType = NodeValType.innerXML, XMP_TempType TP = XMP_TempType.Model)
        {
            string namepath = String.Format("Sheet/cell[@row='{0}' and @col='{1}']", rowno, colno);
            if (TP == XMP_TempType.Formulas)
            {
                namepath = String.Format("Sheet/Formula/cell[@row='{0}' and @col='{1}']", rowno, colno);
            }
            //Console.WriteLine("UpdateXMLValueByRC:  "+namepath+"        "+newvalue);
            //Console.WriteLine("GetXMLValueByRC XML: " + RN.OwnerDocument.InnerXml);
            switch (VType)
            {
                case NodeValType.innerXML:
                    RN.SelectSingleNode(namepath).InnerXml = newvalue;
                    break;
                case NodeValType.Attributes:
                    RN.SelectSingleNode(namepath).Attributes["FormulaText"].Value = newvalue;
                    break;
                default:
                    RN.SelectSingleNode(namepath).InnerText = newvalue;
                    break;
            }
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
    }

    
}
