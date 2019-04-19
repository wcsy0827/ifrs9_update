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
    public class Impairment02 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                List<string> Product_Codes = new List<string>()
                {
                    Product_Code.B_A.GetDescription(),
                    Product_Code.B_B.GetDescription(),
                    Product_Code.B_P.GetDescription()
                };
                var _Product_Codes = Product_Codes.stringListToInSql();
                string sql = string.Empty;
                sql += $@"
WITH TEMPA AS
(
   SELECT 
   IAS39_CATEGORY AS A, --公報分類
   Bond_Aera AS B, --國內\國外
   Asset_Type AS C, --金融商品分類
   Currency_Code AS D, --幣別
   ASSET_SEG AS E, --資產區隔 
   IH_OS AS F, --自操\委外
   Ori_Amount AS G, --帳列面額(原幣)
   Ori_Amount_Ori_Ex AS H, --帳列面額(成本匯率台幣)
   Principal AS I, --攤銷後成本(原幣) 
   Ori_Ex_rate AS J, --成本匯率
   Amort_Amt_Ori_Tw AS K, --攤銷後成本(成本匯率台幣)
   Ex_rate AS L, --報表日匯率
   Amort_Amt_Tw AS M, --攤銷後成本(報表日匯率台幣)
   Market_Value_Ori AS N, --市價(原幣)
   Market_Value_TW AS O, --市價(報表日匯率台幣)
   Principal_EL_Ori AS P, --累計減損-本金(原幣)
   Principal_EL_Ori_Ex AS Q, --累計減損-本金(成本匯率台幣)
   Principal_EL_Ex AS R, --累計減損-本金(報表日匯率台幣)
   Principal_Diff_Tw AS S, --累計減損匯兌損益(台幣)
   CASE WHEN IAS39_CATEGORY = 'FVOCI債務'
     THEN '透過其他綜合損益按公允價值衡量之金融資產FVOCI'
	 WHEN IAS39_CATEGORY = 'AC'
	 THEN '按攤銷後成本衡量之金融資產AC'
	 ELSE IAS39_CATEGORY
   END AS Title1,
   CASE WHEN ASSET_SEG like '%外幣%'
     THEN '-外幣保單'
   	 WHEN SUBSTRING (Portfolio,0,4) = 'OIU'
   	 THEN '-OIU'
   	 ELSE ' '
   END AS Title2
   FROM 
(select * from IFRS9_Bond_Report
WHERE Report_Date = @reportDate) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
),
TEMPB AS
(
   SELECT
   '應收利息' AS A, --公報分類
   Bond_Aera AS B, --國內\國外
   Asset_Type AS C, --金融商品分類
   Currency_Code AS D, --幣別
   ASSET_SEG AS E, --資產區隔
   IH_OS AS F, --自操\委外
   null  AS G,
   null  AS H,
   Interest_Receivable AS I, --帳列面額(原幣)
   null  AS J, 
   null  AS K,
   Ex_rate AS L, --報表日匯率
   Interest_Receivable_Tw AS M, --帳列金額(報表日匯率台幣)  -------
   null  AS N,
   null  AS O,
   Interest_Receivable_EL_Ori AS P, --累計減損-應收利息(原幣)
   null AS Q,
   Interest_Receivable_EL_Ex AS R, --累計減損-應收利息(報表日匯率台幣)
   null AS S,
   '應收利息' AS Title1,
   CASE WHEN ASSET_SEG like '%外幣%'
        THEN '-外幣保單'
   	    WHEN SUBSTRING (Portfolio,0,4) = 'OIU'
   	    THEN '-OIU'
   	    ELSE ' '
   END AS Title2
   FROM 
(select * from IFRS9_Bond_Report
WHERE Report_Date = @reportDate) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
),
VAL AS
(
SELECT 
Title1 + Title2 AS Title,
A, --公報分類
B, --國內\國外
C, --金融商品分類
D, --幣別
E, --資產區隔
F, --自操\委外
SUM(G) AS G, --帳列面額(原幣)
SUM(H) AS H, --帳列面額(成本匯率台幣)
SUM(I) AS I, --攤銷後成本(原幣)
null AS J, --成本匯率 --成本匯率 不顯示 2018/04/27
SUM(K) AS K, --攤銷後成本(成本匯率台幣)
L AS L, --報表日匯率
SUM(M) AS M, --攤銷後成本(報表日匯率台幣)
SUM(N) AS N, --市價(原幣)
SUM(O) AS O, --市價(報表日匯率台幣)
SUM(P) AS P, --累計減損-本金(原幣)
SUM(Q) AS Q, --累計減損-本金(成本匯率台幣)
SUM(R) AS R, --累計減損-本金(報表日匯率台幣)
SUM(S) AS S, --累計減損匯兌損益-本金(台幣)
Title1,
Title2
FROM TEMPA
WHERE LEFT(TEMPA.Title1,4) != 'FVPL' --FVPL的總表並不需要，能否排除? 但FVPL屬應收利息的部份仍要於應收提列總表納入 2018/02/07
GROUP BY A,B,C,D,E,F,L,Title1,Title2 --成本匯率 不Group 2018/04/27
UNION ALL
SELECT 
Title1 + Title2 AS Title,
A, --公報分類
B, --國內\國外
C, --金融商品分類
D, --幣別
E, --資產區隔
F, --自操\委外
null AS G,
null AS H,
SUM(I) AS I, --帳列金額(原幣)
null AS J,
null AS K,
L AS L, --報表日匯率
SUM(M) AS M, --帳列金額(報表日匯率台幣)
null AS N,
null AS O,
SUM(P) AS P, --累計減損-應收利息(原幣)
null AS Q,
SUM(R) AS R, --累計減損-應收利息(報表日匯率台幣)
null AS S,
Title1,
Title2
FROM TEMPB
GROUP BY A,B,C,D,E,F,L,Title1,Title2
)
SELECT
*
FROM VAL
ORDER BY CASE WHEN A = '應收利息'  
         THEN 1
		 ELSE 0
		 END,  
Title,A,B,C,D,E,F
; ";
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