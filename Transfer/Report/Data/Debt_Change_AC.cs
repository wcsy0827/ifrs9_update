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
    public class Debt_Change_AC : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();

            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"

WITH TEMPT AS --期末
(
   Select * from 
   IFRS9_Bond_Report
   WHERE Report_Date = @Final
   AND IAS39_CATEGORY = @IAS39_CATEGORY 
),
TEMPT1 AS --期初
(
   Select * from 
   IFRS9_Bond_Report
   WHERE Report_Date = @Initial
   AND IAS39_CATEGORY = @IAS39_CATEGORY 
),
TEMPI AS --表3
(
   Select 
   T.IAS39_CATEGORY AS IAS39_CATEGORY_T,--期末公報分類(期末E)
   T1.IAS39_CATEGORY AS IAS39_CATEGORY_T1, --期初公報分類(期初E)
   T.Impairment_Stage AS Impairment_Stage_T, --期末減損階段(U)
   T1.Impairment_Stage AS Impairment_Stage_T1, --期初減損階段(T)
   T1.Principal_EL_Ex AS Principal_EL_Ex_T1_3, --期初累計減損_本金(期初報表日匯率台幣)(表3-AG)
   --期初剩餘部位之期末ECL(報表日匯率台幣)(表3-AR) = (表3-AG)-(表3-AN)-(表3-AP)
   CASE WHEN ISNULL(T.Ori_Amount,0) = 0
      THEN T1.Principal_EL_Ex
	  ELSE 0
   END AS Excluding_Monetary_Assets_ECL_Ex_3, --除列之金融資產ECL(期初報表日匯率台幣)(表3-AN)
   CASE WHEN (T.Ori_Amount > 0 AND (T.Ori_Amount - T1.Ori_Amount) < 0)
      THEN  T1.Principal_EL_Ex * ABS( (T.Ori_Amount - T1.Ori_Amount) /  T1.Ori_Amount )
	  ELSE 0
   END AS Section_Excluding_Monetary_Assets_ECL_Ex_3 --部分除列之金融資產ECL(報表日匯率台幣)(表3-AP)
   from
   TEMPT1 AS T1  
   left join
   TEMPT AS T
   ON T1.Bond_Number = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Bond_Number_Old 
	       ELSE T.Bond_Number  
	  END
   AND T1.Lots = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Lots_Old
		   ELSE T.Lots
	  END
   AND T1.Portfolio_Name = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Portfolio_Name_Old
		   ELSE T.Portfolio_Name
	  END
),
TEMPF AS --表4
(
   Select 
   T.IAS39_CATEGORY AS IAS39_CATEGORY_T,--期末公報分類(期末E)
   T1.IAS39_CATEGORY AS IAS39_CATEGORY_T1, --期初公報分類(期初E)
   T.Impairment_Stage AS Impairment_Stage_T, --期末減損階段(U)
   T1.Impairment_Stage AS Impairment_Stage_T1, --期初減損階段(T)
   T.Principal_EL_Ex AS Principal_EL_Ex_T_4, --本期累計減損_本金(報表日匯率台幣)(表4-AK)
   CASE WHEN ISNULL(T1.Ori_Amount,0) = 0
      THEN T.Principal_EL_Ex
	  ELSE 0
   END AS Purchase_Monetary_Assets_ECL_Ex_4, --新購入之金融資產ECL(報表日匯率台幣)(表4-AN)
   CASE WHEN (T1.Ori_Amount > 0 AND (T.Ori_Amount - T1.Ori_Amount) > 0)
      THEN T.Principal_EL_Ex * (T.Ori_Amount - T1.Ori_Amount) / T1.Ori_Amount
	  ELSE 0
   END AS Section_Purchase_Monetary_Assets_ECL_Ex_4, --部分新增之金融資產ECL(報表日匯率台幣)(表4-AP)
   --期初剩餘部位之期末ECL(報表日匯率台幣)(表4-AR) = (表4-AK)-(表4-AN)-(表4-AP)
   T1.Ex_rate AS Ex_rate_T1_4, --期初報表日匯率(表4-AC)
   T.Principal_EL_Ori AS Principal_EL_Ori_T_4, --本期累計減損_本金(原幣)(表4-AI)
   CASE WHEN ISNULL(T1.Ori_Amount,0) = 0
      THEN T.Principal_EL_Ori
	  ELSE 0
   END AS Purchase_Monetary_Assets_ECL_Ori_4, --新購入之金融資產ECL(原幣)(表4-AM)
   CASE WHEN (T1.Ori_Amount > 0 AND (T.Ori_Amount - T1.Ori_Amount) > 0)
      THEN T.Principal_EL_Ori * (T.Ori_Amount - T1.Ori_Amount) / T1.Ori_Amount
	  ELSE 0
   END AS Section_Purchase_Monetary_Assets_ECL_Ori_4 --部分新增之金融資產ECL(原幣)(表4-AO)
   --期初剩餘部位之期末ECL(原幣)(表4-AQ) = (表4-AI)-(表4-AM)-(表4-AO)
   --期初剩餘部位之期末ECL(期初報表日匯率台幣) (表4-AS) = (表4-AQ)*(表4-AC)
   --匯兌變動影響(表4-AT) = (表4-AR)-(表4-AS)
   from
   TEMPT AS T
   left join
   TEMPT1 AS T1
   ON T1.Bond_Number = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Bond_Number_Old 
	       ELSE T.Bond_Number  
	  END
   AND T1.Lots = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Lots_Old
		   ELSE T.Lots
	  END
   AND T1.Portfolio_Name = 
      CASE WHEN T.ISIN_Changed_Ind = 'Y'
	       THEN T.Portfolio_Name_Old
		   ELSE T.Portfolio_Name
	  END
)
,
TEMP1 AS
( 
Select
IAS39_CATEGORY_T,--期末公報分類(期末E)
IAS39_CATEGORY_T1, --期初公報分類(期初E)
Impairment_Stage_T, --期末減損階段(U)
Impairment_Stage_T1, --期初減損階段(T)
Principal_EL_Ex_T1_3, --期初累計減損_本金(期初報表日匯率台幣)(表3-AG)
Principal_EL_Ex_T1_3 - Excluding_Monetary_Assets_ECL_Ex_3 - Section_Excluding_Monetary_Assets_ECL_Ex_3
AS Final_Remaining_Part_ECL_Ex_3, --期初剩餘部位之期末ECL(報表日匯率台幣)(表3-AR) = (表3-AG)-(表3-AN)-(表3-AP)
Excluding_Monetary_Assets_ECL_Ex_3, --除列之金融資產ECL(期初報表日匯率台幣)(表3-AN)
Section_Excluding_Monetary_Assets_ECL_Ex_3 --部分除列之金融資產ECL(報表日匯率台幣)(表3-AP)
from TEMPI
)
,
TEMP2 AS
( 
Select
IAS39_CATEGORY_T,--期末公報分類(期末E)
IAS39_CATEGORY_T1, --期初公報分類(期初E)
Impairment_Stage_T, --期末減損階段(U)
Impairment_Stage_T1, --期初減損階段(T)
Purchase_Monetary_Assets_ECL_Ex_4, --新購入之金融資產ECL(報表日匯率台幣)(表4-AN)
Section_Purchase_Monetary_Assets_ECL_Ex_4, --部分新增之金融資產ECL(報表日匯率台幣)(表4-AP)
(Principal_EL_Ori_T_4 - Purchase_Monetary_Assets_ECL_Ori_4 - Section_Purchase_Monetary_Assets_ECL_Ori_4) *
--期初剩餘部位之期末ECL(原幣)(表4-AQ) = (表4-AI)-(表4-AM)-(表4-AO)
Ex_rate_T1_4 AS Initial_Remaining_Part_ECL_Ex_Rate_4,
--期初剩餘部位之期末ECL(期初報表日匯率台幣) (表4-AS) = (表4-AQ)*(表4-AC)
(Principal_EL_Ex_T_4 - Purchase_Monetary_Assets_ECL_Ex_4 - Section_Purchase_Monetary_Assets_ECL_Ex_4) -
--期初剩餘部位之期末ECL(報表日匯率台幣)(表4-AR) = (表4-AK)-(表4-AN)-(表4-AP)
((Principal_EL_Ori_T_4 - Purchase_Monetary_Assets_ECL_Ori_4 - Section_Purchase_Monetary_Assets_ECL_Ori_4) * Ex_rate_T1_4)
AS Exchange_Changes_4
 --匯兌變動影響(表4-AT) = (表4-AR)-(表4-AS)
from TEMPF
)
Select
--期初餘額(5)
(Select SUM(TEMP1.Principal_EL_Ex_T1_3) from TEMP1 where TEMP1.Impairment_Stage_T1 = '1') AS E5,
--'(T-1期)Impairment_Stage = 1 期初累計減損_本金(期初報表日匯率台幣)(T - 1期)Principal_EL_Ex 報表3中，E欄 = AC且T欄 = 1之AG欄加總'
(Select SUM(TEMP1.Principal_EL_Ex_T1_3) from TEMP1 where TEMP1.Impairment_Stage_T1 = '2') AS G5,
--'(T-1期)Impairment_Stage=2 期初累計減損_本金(期初報表日匯率台幣) (T-1期)Principal_EL_Ex 報表3中，E欄=AC且T欄=2之AG欄加總'
(Select SUM(TEMP1.Principal_EL_Ex_T1_3) from TEMP1 where TEMP1.Impairment_Stage_T1 = '3') AS H5,
--'(T-1期)Impairment_Stage=3 期初累計減損_本金(期初報表日匯率台幣) (T-1期)Principal_EL_Ex 報表3中，E欄=AC且T欄=3之AG欄加總'
--因期初已認列之金融工具所產生之變動：(6)
-- 轉為存續期間預期信用損失(7)
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '1' and TEMP1.Impairment_Stage_T = '2') AS E7,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=1且T期Impairment_Stage=2條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=1、U欄=2之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T1 IN('1', '3') and TEMP1.Impairment_Stage_T = '2') AS G7,
--'(加項) 加總報表3中符合T-1期Impairment_Stage=1 or 3且T期Impairment_Stage=2條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=1 or 3、U欄=2之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '3' and TEMP1.Impairment_Stage_T = '2') AS H7,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=3且T期Impairment_Stage=2條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=3、U欄=2之AR欄加總'
-- 轉為信用減損金融資產(8)
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '1' and TEMP1.Impairment_Stage_T = '3') AS E8,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=1且T期Impairment_Stage=3條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=1、U欄=3之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '2' and TEMP1.Impairment_Stage_T = '3') AS G8,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=2且T期Impairment_Stage=3條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=2、U欄=3之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T1 IN('1', '2') and TEMP1.Impairment_Stage_T = '3') AS H8,
--'(加項) 加總報表3中符合T-1期Impairment_Stage=1 or 2且T期Impairment_Stage=3條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=1 or 2、U欄=3之AR欄加總'
-- 轉為12個月預期信用損失(9)
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T1 IN('2', '3') and TEMP1.Impairment_Stage_T = '1') AS E9,
--'(加項) 加總報表3中符合T-1期Impairment_Stage=2 or 3且T期Impairment_Stage=1條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=2 or 3、U欄=1之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '2' and TEMP1.Impairment_Stage_T = '1') AS G9,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=2且T期Impairment_Stage=1條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=2、U欄=1之AR欄加總'
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '3' and TEMP1.Impairment_Stage_T = '1') AS H9,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=3且T期Impairment_Stage=1條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄=AC、T欄=3、U欄=1之AR欄加總'
-- 於當期除列之金融資產(10)
(Select(SUM(TEMP1.Excluding_Monetary_Assets_ECL_Ex_3) + SUM(TEMP1.Section_Excluding_Monetary_Assets_ECL_Ex_3)) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '1') AS E10,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=1條件者之除列之金融資產ECL(期初報表日匯率台幣)及部分除列之金融資產ECL(期初報表日匯率台幣) 報表3中，E欄=AC及T欄=1之AN欄及AP欄加總'
(Select(SUM(TEMP1.Excluding_Monetary_Assets_ECL_Ex_3) + SUM(TEMP1.Section_Excluding_Monetary_Assets_ECL_Ex_3)) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '2') AS G10,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=2條件者之除列之金融資產ECL(期初報表日匯率台幣)及部分除列之金融資產ECL(期初報表日匯率台幣) 報表3中，E欄=AC及T欄=2之AN欄及AP欄加總'
(Select(SUM(TEMP1.Excluding_Monetary_Assets_ECL_Ex_3) + SUM(TEMP1.Section_Excluding_Monetary_Assets_ECL_Ex_3)) * -1 from TEMP1 where TEMP1.Impairment_Stage_T1 = '3') AS H10,
--'(減項) 加總報表3中符合T-1期Impairment_Stage=3條件者之除列之金融資產ECL(期初報表日匯率台幣)及部分除列之金融資產ECL(期初報表日匯率台幣) 報表3中，E欄=AC及T欄=3之AN欄及AP欄加總'
--創始或購入之新金融資產(11)
(Select(SUM(TEMP2.Purchase_Monetary_Assets_ECL_Ex_4) + SUM(TEMP2.Section_Purchase_Monetary_Assets_ECL_Ex_4)) from TEMP2 where TEMP2.Impairment_Stage_T = '1') AS E11,
--'(加項) 加總報表4中符合T期Impairment_Stage=1條件者之新購入之金融資產ECL(報表日匯率台幣)及部分新增之金融資產ECL(報表日匯率台幣) 報表4中，E欄=AC、U欄=1之AN欄及AP欄加總'
(Select(SUM(TEMP2.Purchase_Monetary_Assets_ECL_Ex_4) + SUM(TEMP2.Section_Purchase_Monetary_Assets_ECL_Ex_4)) from TEMP2 where TEMP2.Impairment_Stage_T = '2') AS G11,
--'(加項) 加總報表4中符合T期Impairment_Stage=2條件者之新購入之金融資產ECL(報表日匯率台幣)及部分新增之金融資產ECL(報表日匯率台幣) 報表4中，E欄=AC、U欄=2之AN欄及AP欄加總'
(Select(SUM(TEMP2.Purchase_Monetary_Assets_ECL_Ex_4) + SUM(TEMP2.Section_Purchase_Monetary_Assets_ECL_Ex_4)) from TEMP2 where TEMP2.Impairment_Stage_T = '3') AS H11,
--'(加項) 加總報表4中符合T期Impairment_Stage=3條件者之新購入之金融資產ECL(報表日匯率台幣)及部分新增之金融資產ECL(報表日匯率台幣) 報表4中，E欄=AC、U欄=3之AN欄及AP欄加總'
--本期沖銷(12)
--本期收回(13)
--模型 / 風險參數之改變(14)
(Select SUM(TEMP2.Initial_Remaining_Part_ECL_Ex_Rate_4) from TEMP2 where TEMP2.Impairment_Stage_T = '1') -
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T = '1') AS E14,
--'(加項) 加總報表4中符合T期Impairment_Stage=1條件者之期初剩餘部位之期末ECL(期初報表日匯率台幣) 報表4中，E欄=AC、U欄=1之AS欄加總
--(減項) 加總報表3中符合T期Impairment_Stage = 1條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄 = AC、U欄 = 1之AR欄加總'
(Select SUM(TEMP2.Initial_Remaining_Part_ECL_Ex_Rate_4) from TEMP2 where TEMP2.Impairment_Stage_T = '2') -
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T = '2') AS G14,
--'(加項) 加總報表4中符合T期Impairment_Stage=2條件者之期初剩餘部位之期末ECL(期初報表日匯率台幣) 報表4中，E欄=AC、U欄=2之AS欄加總
--(減項) 加總報表3中符合T期Impairment_Stage = 2條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄 = AC、U欄 = 2之AR欄加總'
(Select SUM(TEMP2.Initial_Remaining_Part_ECL_Ex_Rate_4) from TEMP2 where TEMP2.Impairment_Stage_T = '3') -
(Select SUM(TEMP1.Final_Remaining_Part_ECL_Ex_3) from TEMP1 where TEMP1.Impairment_Stage_T = '3') AS H14,
--'(加項) 加總報表4中符合T期Impairment_Stage=3條件者之期初剩餘部位之期末ECL(期初報表日匯率台幣) 報表4中，E欄=AC、U欄=3之AS欄加總
--(減項) 加總報表3中符合T期Impairment_Stage = 3條件者之期末剩餘部位之ECL(期初報表日匯率台幣) 報表3中，E欄 = AC、U欄 = 3之AR欄加總'
--匯兌及其他變動(15)
(Select SUM(TEMP2.Exchange_Changes_4) from TEMP2 where TEMP2.Impairment_Stage_T = '1') AS E15,
--'(加項) 加總報表4中符合T期Impairment_Stage=1條件者之匯兌變動影響 報表4中，E欄=AC、U欄=1之AT欄加總'
(Select SUM(TEMP2.Exchange_Changes_4) from TEMP2 where TEMP2.Impairment_Stage_T = '2') AS G15,
--'(加項) 加總報表4中符合T期Impairment_Stage=2條件者之匯兌變動影響 報表4中，E欄=AC、U欄=2之AT欄加總'
(Select SUM(TEMP2.Exchange_Changes_4) from TEMP2 where TEMP2.Impairment_Stage_T = '3') AS H15
--'(加項) 加總報表4中符合T期Impairment_Stage=3條件者之匯兌變動影響 報表4中，E欄=AC、U欄=3之AT欄加總'
--期末餘額(16)
";
                using (var cmd = new SqlCommand(sql, conn))
                {
                    string _IAS39_CATEGORY = "AC";
                    string Initial = parms.Where(x => x.key == "Initial").FirstOrDefault()?.value ?? string.Empty;
                    string Final = parms.Where(x => x.key == "Final").FirstOrDefault()?.value ?? string.Empty;
                    DateTime Initial_Date = DateTime.Parse(Initial);
                    DateTime Final_Date = DateTime.Parse(Final);
                    var _Initial = new SqlParameter();
                    _Initial.SqlDbType = SqlDbType.VarChar;
                    _Initial.ParameterName = "@Initial";
                    _Initial.Value = Initial;
                    _Initial.Direction = ParameterDirection.Input;
                    var _Final = new SqlParameter();
                    _Final.SqlDbType = SqlDbType.VarChar;
                    _Final.ParameterName = "@Final";
                    _Final.Value = Final;
                    _Final.Direction = ParameterDirection.Input;
                    var IAS39_CATEGORY = new SqlParameter();
                    IAS39_CATEGORY.SqlDbType = SqlDbType.VarChar;
                    IAS39_CATEGORY.ParameterName = "@IAS39_CATEGORY";
                    IAS39_CATEGORY.Value = _IAS39_CATEGORY;
                    IAS39_CATEGORY.Direction = ParameterDirection.Input;
                    cmd.Parameters.Add(_Initial); //期初 
                    cmd.Parameters.Add(_Final); //期末
                    cmd.Parameters.Add(IAS39_CATEGORY);
                    cmd.CommandTimeout = 0; //190624 USER產報表跳出逾時訊息
                    //cmd.Parameters.Add(new SqlParameter("Initial", Initial)); //期初 
                    //cmd.Parameters.Add(new SqlParameter("Final", Final)); //期末
                    //cmd.Parameters.Add(new SqlParameter("IAS39_CATEGORY", _IAS39_CATEGORY));
                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);

                    var DateRange = new reportParm()
                    {
                        key = "DateRange",
                        value = $@"{Initial_Date.ToString("yyyy/MM/dd")}~{Final_Date.ToString("yyyy/MM/dd")}"
                    };
                    extensionParms.Add(DateRange);
                }
            }
            return resultsTable;
        }

    }
}