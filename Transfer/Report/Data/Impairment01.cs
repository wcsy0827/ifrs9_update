using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using Transfer.Models;
using Transfer.Utility;
using static Transfer.Enum.Ref;

namespace Transfer.Report.Data
{
    public class Impairment01 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"
SELECT 
Security_Name, --Security_Name
Bond_Number, --債券編號
Lots, --Lots
Segment_Name , --債券(資產)名稱
convert(varchar, Origination_Date, 111) AS Origination_Date, --債券購入(認列)日期
convert(varchar, Maturity_Date, 111) AS Maturity_Date , --債券到期日
ISSUER_AREA , --Issuer所屬區域 
Industry_Sector, --對手產業別 
IAS39_CATEGORY, --公報分類 
Bond_Aera, --國內\國外 
Asset_Type, --金融商品分類 
Currency_Code, --幣別 
ASSET_SEG, --資產區隔 
IH_OS, --自操\委外 
Portfolio_Name, --Portfolio 
Portfolio, --Portfolio中文 
PRODUCT, --SMF 
Ori_Amount, --帳列面額(原幣) 
Ori_Amount_Ori_Ex, -- 帳列面額(成本匯率台幣) 
Principal, --攤銷後成本(原幣) 
Ori_Ex_rate, --成本匯率 
Amort_Amt_Ori_Tw, --攤銷後成本(成本匯率台幣) 
Ex_rate, --報表日匯率 
Amort_Amt_Tw, --攤銷後成本(報表日匯率台幣) 
Market_Value_Ori, --市價(原幣) 
Market_Value_TW, --市價(報表日匯率台幣) 
Interest_Receivable, --應收利息(原幣) 
Interest_Receivable_tw AS Int_Receivable_tw, --應收利息(報表日匯率台幣) 
Y1_EL ,--一年期預期信用損失(原幣)
Lifetime_EL, --存續期間預期信用損失(原幣)
Impairment_Stage, --減損階段
Original_External_Rating, --購買時評等(主標尺)
Current_External_Rating, --最近一次評等(主標尺)
Current_LGD, --LGD
PD, --最近一次評等PD
Lien_position, --債券擔保順位 
Current_Int_Rate, --票面利率 
EIR, --EIR 
Principal_Payment_Method_Code, --現金流類型 
EL, --累計減損(原幣)
EL_Ori_Ex, ---累計減損(成本匯率台幣)
EL_Ex, --累計減損(報表日匯率台幣)
EL_EX_Diff, --累計減損匯兌損益(台幣)
Principal_EL_Ori, --累計減損-本金(原幣)
Principal_EL_Ori_Ex, --累計減損-本金(成本匯率台幣)
Principal_EL_Ex, --累計減損-本金(報表日匯率台幣)
Principal_Diff_Tw, --累計減損匯兌損益-本金(台幣)
Interest_Receivable_EL_Ori, --累計減損-利息(原幣)
Interest_Receivable_EL_Ex, --累計減損-利息(報表日匯率台幣)
CASE WHEN Maturity_Date is not null
     THEN (
	 CASE WHEN (DATEDIFF(MONTH,Report_Date,Maturity_Date)/CAST(12 AS float)) > 1
	      THEN '非流動'
		  ELSE '流動'
	 END)
ELSE '非流動'
END AS Current_Noncurrent --流動/非流動
From
(select * from IFRS9_Bond_Report
WHERE Report_Date = @reportDate) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
";
                using (var cmd = new SqlCommand(sql, conn))
                {
                    string reportDate = parms.Where(x => x.key == "reportDate").FirstOrDefault()?.value ?? string.Empty;
                    string Group_Product_Code = parms.Where(x => x.key == "Group_Product_Code").FirstOrDefault()?.value ?? string.Empty;
                    cmd.Parameters.Add(new SqlParameter("reportDate", reportDate));
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