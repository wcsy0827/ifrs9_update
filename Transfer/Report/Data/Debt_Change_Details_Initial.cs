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
    public class Debt_Change_Details_Initial : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            string Initial = parms.Where(x => x.key == "Initial").FirstOrDefault()?.value ?? string.Empty;
            string Final = parms.Where(x => x.key == "Final").FirstOrDefault()?.value ?? string.Empty;
            string Group_Product_Code = parms.Where(x => x.key == "Group_Product_Code").FirstOrDefault()?.value ?? string.Empty;
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"
WITH T AS --期末
(
SELECT D52.* 
FROM 
(select * from IFRS9_Bond_Report
WHERE Report_Date = @Final) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
),
T1 AS --期初
(
SELECT 
D52.*
FROM 
(select * from IFRS9_Bond_Report
WHERE Report_Date = @Initial) D52
JOIN (select PRJID,FLOWID from Flow_Info
where Group_Product_Code = @Group_Product_Code) FI
on D52.PRJID = FI.PRJID
and D52.FLOWID = FI.FLOWID
),
TEMP AS
(
SELECT
 T1.Reference_Nbr, --Reference_Nbr
 T1.Security_Name, --Security Name(A)
 T1.Bond_Number, --債券編號(B)
 T1.Lots, --Lots(C)
 T1.Segment_Name, --債券(資產)名稱(D)
 T1.IAS39_CATEGORY, --公報分類(E)
 T1.Bond_Aera, --國內\國外(F)
 T1.Asset_Type, --金融商品分類(G)
 T1.Currency_Code, --幣別(H)
 T1.ASSET_SEG, --資產區隔(I)
 T1.IH_OS, --自操\委外(J)
 T1.Portfolio, --Portfolio(K)
 T1.PRODUCT, --SMF(L)
 T1.Ori_Amount AS Ori_Amount_T1, --期初帳列面額(原幣)(M)
 T.Ori_Amount AS Ori_Amount_T, --期末帳列面額(原幣)(N)
 T.Ori_Amount - T1.Ori_Amount AS Ori_Amount, --面額增(減)變動(原幣)(O)
 T1.Principal AS Principal_T1, --期初攤銷後成本(原幣)(P)
 T.Principal AS Principal_T, --期末攤銷後成本(原幣)(Q)
 T1.Interest_Receivable AS Interest_Receivable_T1, --期初應收利息(原幣)(R)
 T.Interest_Receivable AS Interest_Receivable_T, --期末應收利息(原幣)(S)
 T1.Impairment_Stage AS Impairment_Stage_T1, --期初減損階段(T)
 T.Impairment_Stage AS Impairment_Stage_T, --期末減損階段(U)
 T1.PD AS PD_T1, --期初第一年違約率(V)
 T.PD AS PD_T, --期末第一年違約率(W)
 CASE WHEN ISNULL((T1.Principal * T1.Current_LGD) , 0) = 0
      THEN NULL
      ELSE
      T1.Principal_EL_Ori/(T1.Principal * T1.Current_LGD)
 END AS Default_Rate_T1, --期初違約率(X)
 CASE WHEN ISNULL((T.Principal * T.Current_LGD) , 0) = 0
      THEN NULL
      ELSE
      T.Principal_EL_Ori/(T.Principal * T.Current_LGD)
 END AS Default_Rate_T, --期末違約率(Y)
 T1.Current_LGD AS Current_LGD_T1, --期初違約損失率(Z)
 T.Current_LGD AS Current_LGD_T, --期末違約損失率(AA)
 T1.Ori_Ex_rate, --成本匯率(AB)
 T1.Ex_rate AS Ex_rate_T1, --期初報表日匯率(AC)
 T.Ex_rate AS Ex_rate_T, --報表日匯率(AD)
 T1.Principal_EL_Ori AS Principal_EL_Ori_T1, --期初累計減損_本金(原幣)(AE)
 T1.Principal_EL_Ori_Ex AS Principal_EL_Ori_Ex_T1, --期初累計減損_本金(成本匯率台幣)(AF)
 T1.Principal_EL_Ex AS Principal_EL_Ex_T1, --期初累計減損_本金(期初報表日匯率台幣)(AG)
 T1.Principal_Diff_Tw AS Principal_Diff_Tw_T1, --期初累計減損_本金匯兌損益(台幣)(AH)
 T.Principal_EL_Ori AS Principal_EL_Ori_T, --本期累計減損_本金(原幣)(AI)
 T.Principal_EL_Ori_Ex AS Principal_EL_Ori_Ex_T, --本期累計減損_本金(成本匯率台幣)(AJ)
 T.Principal_EL_Ex AS Principal_EL_Ex_T, --本期累計減損_本金(報表日匯率台幣)(AK)
 T.Principal_Diff_Tw AS Principal_Diff_Tw_T, --本期累計減損_本金匯兌損益(台幣)(AL)
 CASE WHEN ISNULL(T.Ori_Amount,0) = 0
      THEN T1.Principal_EL_Ori
	  ELSE 0
 END AS Excluding_Monetary_Assets_ECL_Ori, --除列之金融資產ECL(原幣)(AM)
 CASE WHEN ISNULL(T.Ori_Amount,0) = 0
      THEN T1.Principal_EL_Ex
	  ELSE 0
 END AS Excluding_Monetary_Assets_ECL_Ex, --除列之金融資產ECL(期初報表日匯率台幣)(AN)
 CASE WHEN (T.Ori_Amount > 0 AND (T.Ori_Amount - T1.Ori_Amount) < 0)
      THEN T1.Principal_EL_Ori * ABS( (T.Ori_Amount - T1.Ori_Amount) /  T1.Ori_Amount )
	  ELSE 0
 END AS Section_Excluding_Monetary_Assets_ECL_Ori, --部分除列之金融資產ECL(原幣)(AO)
 CASE WHEN (T.Ori_Amount > 0 AND (T.Ori_Amount - T1.Ori_Amount) < 0)
      THEN  T1.Principal_EL_Ex * ABS( (T.Ori_Amount - T1.Ori_Amount) /  T1.Ori_Amount )
	  ELSE 0
 END AS Section_Excluding_Monetary_Assets_ECL_Ex --部分除列之金融資產ECL(報表日匯率台幣)(AP)
FROM 
 T1 left join T
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
select 
TEMP.*,
TEMP.Principal_EL_Ori_T1 - TEMP.Excluding_Monetary_Assets_ECL_Ori - TEMP.Section_Excluding_Monetary_Assets_ECL_Ori
AS Final_Remaining_Part_ECL_Ori, --期初剩餘部位之期末ECL(原幣)(AQ) = AE欄-AM欄-AO欄
TEMP.Principal_EL_Ex_T1 - TEMP.Excluding_Monetary_Assets_ECL_Ex - TEMP.Section_Excluding_Monetary_Assets_ECL_Ex
AS Final_Remaining_Part_ECL_Ex --期初剩餘部位之期末ECL(報表日匯率台幣)(AR) = AG欄-AN欄-AP欄
from 
TEMP ;
";


                DateTime Initial_Date = DateTime.Parse(Initial);
                DateTime Final_Date = DateTime.Parse(Final);
                //using (IFRS9DBEntities db = new IFRS9DBEntities())
                //{
                //    var _data = db.Database.DynamicSqlQuery(sql
                //        , null,
                //        new object[] {
                //                  new SqlParameter("Initial", Initial),
                //                  new SqlParameter("Final", Final),
                //                  new SqlParameter("Group_Product_Code", Group_Product_Code)
                //        }
                //     );
                //    db.Database.Connection.Close();
                //    resultsTable.Tables.Add(_data.ToDataTable());
                //}

                using (var cmd = new SqlCommand(sql, conn))
                {
                    var _Initial = new SqlParameter();
                    _Initial.SqlDbType = SqlDbType.VarChar;
                    _Initial.ParameterName = "@Initial";
                    _Initial.Value = Initial;
                    var _Final = new SqlParameter();
                    _Final.SqlDbType = SqlDbType.VarChar;
                    _Final.ParameterName = "@Final";
                    _Final.Value = Final;
                    var _Group_Product_Code = new SqlParameter();
                    _Group_Product_Code.SqlDbType = SqlDbType.VarChar;
                    _Group_Product_Code.ParameterName = "@Group_Product_Code";
                    _Group_Product_Code.Value = Group_Product_Code;
                    cmd.Parameters.Add(_Initial); //期初 
                    cmd.Parameters.Add(_Final); //期末
                    cmd.Parameters.Add(_Group_Product_Code);
                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);
                }

                var DateRange = new reportParm()
                {
                    key = "DateRange",
                    value = $@"{Initial_Date.ToString("yyyy/MM/dd")}~{Final_Date.ToString("yyyy/MM/dd")}"
                };
                extensionParms.Add(DateRange);
            }
            return resultsTable;
        }

    }
}