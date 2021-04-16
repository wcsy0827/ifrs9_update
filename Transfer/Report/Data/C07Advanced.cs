using System;
using System.Data;
using System.Linq;
using System.Xml.Linq;
using Transfer.Utility;
using Transfer.ViewModels;
using System.Data.SqlClient;
using System.Collections.Generic;
using Transfer.Models.Repository;
using System.Text.RegularExpressions;
using Transfer.Models;


namespace Transfer.Report.Data
{
    public class C07Advanced : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            DataSet resultsTable = new DataSet();
            C0Repository c0Repository = new C0Repository();

            string sql = string.Empty;
            string ReferenceNbr = parms.Where(x => x.key == "ReferenceNbr").FirstOrDefault()?.value ?? string.Empty;
            string ClassName = parms.Where(x => x.key == "ClassName").FirstOrDefault()?.value ?? string.Empty;
            string ReportDate = parms.Where(x => x.key == "ReportDate").FirstOrDefault()?.value ?? string.Empty;
            string Version = parms.Where(x => x.key == "Version").FirstOrDefault()?.value ?? string.Empty;
            string GroupProductName = parms.Where(x => x.key == "GroupProductName").FirstOrDefault()?.value ?? string.Empty;
            string ProductCode = parms.Where(x => x.key == "ProductCode").FirstOrDefault()?.value ?? string.Empty;

            ReportDate = ReportDate.Replace('/', '-');
            string GroupProductCode = Regex.Replace(GroupProductName, "[^0-9]", "");

            Tuple<string, List<C07AdvancedViewModel>> queryData = c0Repository.getC07Advanced(GroupProductCode, ProductCode, ReportDate, Version, "", "",true);
            Tuple<string, List<C07AdvancedViewModel>> queryDataSum = c0Repository.getC07Advanced(GroupProductCode, ProductCode, ReportDate, Version, "", "");
            DataTable dt = queryData.Item2.ToDataTable();            
            resultsTable.Tables.Add(dt);

            var sum_Exposure_EL = new reportParm()
            {
                key = "sum_Exposure_EL",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Exposure_EL))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Exposure_EL))).Normalize().ToString().formateThousand()

            };
            var sum_Exposure_Ex=new reportParm()
            {
                key = "sum_Exposure_Ex",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Exposure_Ex))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Exposure_Ex))).Normalize().ToString().formateThousand()

            };
            var sum_Y1_EL=new reportParm()
            {
                key = "sum_Y1_EL",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Y1_EL))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Y1_EL))).Normalize().ToString().formateThousand()
            };
            var sum_Y1_EL_Ex=new reportParm()
            {
                key = "sum_Y1_EL_Ex",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Y1_EL_Ex))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Y1_EL_Ex))).Normalize().ToString().formateThousand()
            };
            var sum_Lifetime_EL = new reportParm()
            {
                key = "sum_Lifetime_EL",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Lifetime_EL))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Lifetime_EL))).Normalize().ToString().formateThousand()
            };
            var sum_Lifetime_EL_Ex = new reportParm()
            {
                key = "sum_Lifetime_EL_Ex",
                //value = Math.Round(queryData.Item2.Sum(x => double.Parse(x.Lifetime_EL_Ex))).ToString("N0")
                value = Math.Round(queryDataSum.Item2.Sum(x => TypeTransfer.stringToDecimal(x.Lifetime_EL_Ex))).Normalize().ToString().formateThousand()
            };
            extensionParms.Add(sum_Exposure_EL);
            extensionParms.Add(sum_Exposure_Ex);
            extensionParms.Add(sum_Y1_EL);
            extensionParms.Add(sum_Y1_EL_Ex);
            extensionParms.Add(sum_Lifetime_EL);
            extensionParms.Add(sum_Lifetime_EL_Ex);



            return resultsTable;
        }
    }
}