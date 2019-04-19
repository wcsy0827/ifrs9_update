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
    public class Impairment03 : ReportData
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
WITH Ver AS
(
   SELECT TOP 1 Version 
   FROM Bond_Account_Info
   WHERE Report_Date = @reportDate
   ORDER BY Version DESC
),
ELTEMP AS
(
    SELECT distinct A41.Reference_Nbr,
    CASE WHEN D66.Assessment_Stage = '2'
         THEN CASE WHEN D66.Qualitative_Pass_Stage2 = 'Y'
              THEN '1'
              ELSE '2'
         END
         WHEN D66.Assessment_Stage = '3'
 	     THEN CASE WHEN D66.Qualitative_Pass_Stage3 = 'Y'
              THEN '1'
              ELSE '3'
         END
	ELSE C07.Impairment_Stage
	END AS Impairment_Stage,
	CASE WHEN D66.Assessment_Stage = '2'
	     THEN C07.Lifetime_EL
		 WHEN D66.Assessment_Stage = '3' and C07.PD <> 0
		 THEN C07.Y1_EL / C07.PD
		 WHEN D66.Assessment_Stage = '3' and C07.PD = 0
		 THEN 0
	ELSE C07.EL
	END AS EL,
    C07.PD AS PD
	FROM (SELECT _A41.Reference_Nbr,
                 _A41.Report_Date,
                 _A41.Version
          FROM   Bond_Account_Info _A41 ,Ver
          WHERE  _A41.Report_Date = @reportDate
          AND    _A41.Version = Ver.Version) AS A41
	LEFT JOIN 
    (
      SELECT C07.*
      FROM EL_Data_Out C07,Ver
      WHERE C07.Report_Date = CONVERT(char(10), @reportDate,126) 
      AND C07.Version = Ver.Version
      AND C07.Product_Code IN ({_Product_Codes})
    )
    AS C07
    ON  A41.Reference_Nbr = C07.Reference_Nbr
	LEFT JOIN 
    (
      SELECT D66.* 
      FROM Bond_Qualitative_Assessment_Result D66,Ver
      WHERE D66.Report_Date = @reportDate
      AND D66.Version = Ver.Version
      AND D66.Assessment_Result_Version = D66.Result_Version_Confirm
    ) 
    AS D66
	ON A41.Reference_Nbr = D66.Reference_Nbr
),
A41 AS
(
SELECT
IAS39_CATEGORY AS A, --公報分類 (42) (A)
Bond_Aera AS B, --國內\國外 (43) (B)
Asset_Type AS C, --金融商品分類 (44) (C)
Currency_Code AS D, --幣別 (45) (D)
ASSET_SEG AS E, --資產區隔 (46) (E)
IH_OS AS F, --自操\委外 (47) (F)
Ori_Amount AS G, --帳列面額(原幣) (15) (G)
Ori_Amount * Ori_Ex_rate AS H, --帳列面額(成本匯率台幣) Ori_Amount_Tw (66) = (17) * (22) (H) 
Principal AS I, --攤銷後成本(原幣) (14) (I)
Ori_Ex_rate AS J, --成本匯率 (22) (J)
Amort_Amt_Ori_Tw AS K, --攤銷後成本(成本匯率台幣) (17) (K)
Ex_rate AS L, --報表日匯率 (21) (L)
Amort_Amt_Tw AS M, --攤銷後成本(報表日匯率台幣) (20) (M)
Market_Value_Ori AS N, --市價(原幣) (50) (N)
Market_Value_TW AS O, --市價(報表日匯率台幣) (51) (O)
-- hide
Reference_Nbr, --帳戶編號
Report_Date, --基準日
Bond_Account_Info.Version, -- 版本
SUBSTRING (Portfolio,0,4) AS OIU, --Portfolio
Interest_Receivable --應收利息(原幣) (16)
from  Bond_Account_Info , Ver
WHERE Report_Date = @reportDate
AND Bond_Account_Info.Version = Ver.version
--AND ASSET_SEG  like '%外幣%'
),
TEMP_Interest AS
(
SELECT 
--A41.A AS A, --公報分類
'應收利息' AS A, --公報分類
A41.B AS B, --國內\國外
A41.C AS C, --金融商品分類
A41.D AS D, --幣別
A41.E AS E, --資產區隔
A41.F AS F, --自操\委外
null  AS G,
null  AS H,
A41.G AS I, --帳列面額
null  AS J, 
null  AS K,
A41.L AS L, --報表日匯率
A41.G * A41.L AS M, --帳列金額(報表日匯率台幣)
null  AS N,
null  AS O,
A41.Interest_Receivable * T.PD * C01.Current_LGD AS P,--累計減損-應收利息(原幣)
null  AS Q,
A41.L * (A41.Interest_Receivable * T.PD * C01.Current_LGD)  AS R, --累計減損-應收利息(報表日匯率台幣)
null  AS S,
'應收利息' AS Title1,
CASE WHEN A41.E like '%外幣%'
     THEN '-外幣保單'
	 WHEN A41.OIU = 'OIU'
	 THEN '-OIU'
	 ELSE ' '
END AS Title2
FROM A41
JOIN ELTEMP T
ON A41.Reference_Nbr = T.Reference_Nbr
LEFT JOIN 
(
   Select C01.* 
   From EL_Data_In C01,Ver
   Where C01.Report_Date = @reportDate
   AND C01.Version = Ver.Version
   AND C01.Product_Code IN ({_Product_Codes})
)
C01
ON C01.Reference_Nbr = A41.Reference_Nbr
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
FROM TEMP_Interest
GROUP BY A,B,C,D,E,F,L,Title1,Title2)
SELECT * FROM VAL
ORDER BY CASE WHEN A = '應收利息'  
         THEN 1
		 ELSE 0
		 END,  
Title,A,B,C,D,E,F
";
                using (var cmd = new SqlCommand(sql, conn))
                {
                    string reportDate = parms.Where(x => x.key == "reportDate").FirstOrDefault()?.value ?? string.Empty;
                    cmd.Parameters.Add(new SqlParameter("reportDate", reportDate));
                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);                  
                }
            }
            return resultsTable;
        }

    }
}