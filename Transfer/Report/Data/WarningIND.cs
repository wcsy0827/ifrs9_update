using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using Transfer.Models;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class WarningIND : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"
WITH TempD62 AS
(
select TOP 1  Version  from Bond_Basic_Assessment where Report_Date = @ReportDate and Version=@Version ORDER BY Version DESC
),
TempD62All AS
(
select 
A41.Assessment_Sub_Kind AS Assessment_Sub_Kind, --產業別 (A41#48)
D62.Bond_Type AS Bond_Type, --債券種類 (D62#5)
D62.Bond_Number AS Bond_Number, --債券編號 (D62#8)
D62.Lots AS Lots, --Lots (D62#9)
D62.Portfolio_Name AS Portfolio_Name, -- Portfolio_英文 (D62#11)
D62.Portfolio AS Portfolio, --Portfolio_中文 (D62#10)
A41.Segment_Name, --債券(資產)名稱 (A41#4)
convert(varchar, A41.Origination_Date, 111) AS Origination_Date, --債券購入(認列)日期 (62#12)
D62.Cost_Value, --攤銷後之成本數(原幣) (D62#20)
D62.Market_Value_Ori, --市價(原幣) (D62#21)
D62.Value_Change_Ratio, --未實現累計損失率_本月 (D62#22)
D62.Chg_In_Spread, --信用利差_本月 (D62#30)
CASE WHEN D62.Original_External_Rating is null
     THEN 'N'
	 ELSE 'Y'
END AS Original_External_Rating_Flag, --是否有購買日信評 (判斷 D62#13)
D62.Original_External_Rating, --購買日信評(主標尺轉換值)  (D62#13)
D62.Current_External_Rating,  --報導日信評(主標尺轉換值)  (D62#14)
D62.Rating_Ori_Good_IND, --原始購買信評是否為低信用風險等級  (D62#15)
D62.Rating_Curr_Good_Ind, --報導日信評是否為低信用風險等級 (D62#16)
D62.Curr_Ori_Rating_Diff, --評等下降數  (D62#17)
D62.Accumulation_Loss_This_Month, --未實現累計損失率逾越30%(含)連續月數 (D62#24)
D62.Chg_In_Spread_This_Month,--信用利差拓寬300基點(含)以上連續月數 (D62#33)
D62.Map_Rule_Id_D71  AS MapRule, --對應規則參數編號 (D62#29)
D62.Version
from 
(SELECT _D62.* 
 FROM Bond_Basic_Assessment _D62 , TempD62
 WHERE _D62.Report_Date = @ReportDate
 AND   _D62.Version = TempD62.Version
 AND   _D62.Warning_IND = 'Y' --預警名單
 ) AS D62 
JOIN 
(
  SELECT A41.* 
  FROM Bond_Account_Info A41,TempD62
  WHERE A41.Report_Date = @ReportDate
  AND A41.Version = TempD62.Version
)
AS A41
ON D62.Reference_Nbr = A41.Reference_Nbr)
Select * from TempD62All
union 
select
null AS Assessment_Sub_Kind,
null AS Bond_Type,
null AS Bond_Number,
null AS Lots,
null AS Portfolio_Name,
null AS Portfolio,
null AS Segment_Name,
null AS Origination_Date,
null AS Cost_Value,
null AS Market_Value_Ori,
null AS Value_Change_Ratio,
null AS Chg_In_Spread,
null AS Original_External_Rating_Flag,
null AS Original_External_Rating,
null AS Current_External_Rating,
null AS Rating_Ori_Good_IND, 
null AS Rating_Curr_Good_Ind, 
null AS Curr_Ori_Rating_Diff, 
null AS Accumulation_Loss_This_Month, 
null AS Chg_In_Spread_This_Month,
null AS MapRule, 
TempD62.Version AS Version
from TempD62
";
                using (var cmd = new SqlCommand(sql, conn))
                {
                    string reportDate = parms.Where(x => x.key == "ReportDate").FirstOrDefault()?.value ?? string.Empty;
                    string version = parms.Where(x => x.key == "Version").FirstOrDefault()?.value ?? "0";
                    cmd.Parameters.Add(new SqlParameter("ReportDate", reportDate));
                    cmd.Parameters.Add(new SqlParameter("Version", version));
                    Extension.NlogSet(cmd.CommandText);
                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);

                    var Version = new reportParm()
                    {
                        key = "Version",
                        value = version
                    };
                     
                    //var dt = resultsTable.Tables[0].AsEnumerable();
                    //if (dt.Any())
                    //{
                    //    var first = dt.First();
                    //    Version.value = first.Field<int>("Version").ToString();
                    //}
                    extensionParms.Add(Version);
                }
            }
            return resultsTable;
        }

    }
}