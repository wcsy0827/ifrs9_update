using System;
using System.Data;
using System.Linq;
using System.Xml.Linq;
using Transfer.Utility;
using Transfer.ViewModels;
using System.Data.SqlClient;
using static Transfer.Enum.Ref;
using System.Collections.Generic;
using Transfer.Models.Repository;

namespace Transfer.Report.Data
{
    public class D64 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            DataSet resultsTable = new DataSet();
            D6Repository d6Repository = new D6Repository();

            string sql = string.Empty;
            string ReferenceNbr = parms.Where(x => x.key == "ReferenceNbr").FirstOrDefault()?.value ?? string.Empty;
            string ClassName = parms.Where(x => x.key == "ClassName").FirstOrDefault()?.value ?? string.Empty;
            string ReportDate = parms.Where(x => x.key == "ReportDate").FirstOrDefault()?.value ?? string.Empty;
            string Version = parms.Where(x => x.key == "Version").FirstOrDefault()?.value ?? string.Empty;
            int ver = 0;
            Int32.TryParse(Version, out ver);

            DateTime reportDate = DateTime.Parse(ReportDate);

            List<D63ViewModel> queryData = d6Repository.getD63(reportDate, "", Evaluation_Status_Type.All, "All", false, "",ver);
            List<D64ViewModel> _queryData = d6Repository.D63GetD64(queryData);
            DataTable dt = _queryData.ToDataTable();
            resultsTable.Tables.Add(dt);

            //XDocument rdlcXml = XDocument.Load(AppDomain.CurrentDomain.BaseDirectory + "/Report/Rdlc/" + ClassName + ".rdlc");
            //XNamespace xmlns = rdlcXml.Root.FirstAttribute.Value;

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    if (i != 0) { sql += "UNION ALL" + "\n"; }
            //    sql += "SELECT" + "\n";
            //    foreach (XElement element in rdlcXml.Descendants(xmlns + "Field"))
            //    {
            //        string Field = element.FirstAttribute.Value;
            //        sql += "'" + dt.Rows[i].Field<string>(Field) + "' AS " + Field + ",\n";
            //    }
            //    sql = sql.Trim().Trim(',') + "\n";
            //}
            //if (dt.Rows.Count == 0)
            //{
            //    sql += "SELECT" + "\n";
            //    foreach (XElement element in rdlcXml.Descendants(xmlns + "Field"))
            //    {
            //        string Field = element.FirstAttribute.Value;
            //        sql += "'' AS " + Field + ",\n";
            //    }
            //}
            //sql = sql.Trim().Trim(',');

            //using (SqlConnection conn = new SqlConnection(defaultConnection))
            //{
            //    using (SqlCommand cmd = new SqlCommand(sql, conn))
            //    {
            //        Extension.NlogSet(cmd.CommandText);
            //        conn.Open();
            //        SqlDataAdapter adapter = new SqlDataAdapter(cmd);
            //        adapter.Fill(resultsTable);
            //    }
            //}
            return resultsTable;
        }
    }
}