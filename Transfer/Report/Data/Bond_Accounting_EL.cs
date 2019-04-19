using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class Bond_Accounting_EL : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;

                sql += $@"
select 
IAS39_CATEGORY,
SUM(CASE WHEN Impairment_Stage = '1' AND Risk_Level = '低風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_1_1,
SUM(CASE WHEN Impairment_Stage = '1' AND Risk_Level = '中風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_1_2,
SUM(CASE WHEN Impairment_Stage = '1' AND Risk_Level = '高風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_1_3,
SUM(CASE WHEN Impairment_Stage = '2' AND Risk_Level = '低風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_2_1,
SUM(CASE WHEN Impairment_Stage = '2' AND Risk_Level = '中風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_2_2,
SUM(CASE WHEN Impairment_Stage = '2' AND Risk_Level = '高風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_2_3,
SUM(CASE WHEN Impairment_Stage = '3' AND Risk_Level = '高風險'
     THEN ISNULL(Accounting_EL,0)
	 ELSE 0
	 END) AS Risk_Level_3_3,
SUM(ISNULL(Principal_EL_Ex,0)) AS Principal_EL_Ex
 from
(select Reference_Nbr, IAS39_CATEGORY,Impairment_Stage,Risk_Level,Accounting_EL 
from Bond_Accounting_EL 
where Report_Date = @Report_Date
and IAS39_CATEGORY not like 'FVPL%' ) AS BAE
left join
(
select D52.Reference_Nbr,D52.Principal_EL_Ex from 
(select Reference_Nbr,Principal_EL_Ex,PRJID,FLOWID  from IFRS9_Bond_Report 
where Report_Date = @Report_Date) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
) AS ILR
on BAE.Reference_Nbr = ILR.Reference_Nbr
group by IAS39_CATEGORY ;
";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    string reportDate = parms.Where(x => x.key == "Report_Date").FirstOrDefault()?.value ?? string.Empty;
                    string Group_Product_Code = parms.Where(x => x.key == "Group_Product_Code").FirstOrDefault()?.value ?? string.Empty;
                    cmd.Parameters.Add(new SqlParameter("Report_Date", reportDate));
                    cmd.Parameters.Add(new SqlParameter("Group_Product_Code", Group_Product_Code));
                    conn.Open();

                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);
                }
            }

            return resultsTable;
        }
    }
}