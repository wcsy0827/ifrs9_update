using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_5 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"
select 
convert(varchar,  C01.Report_Date, 111) Report_Date, --評估基準日/報導日
convert(varchar,  C01.Processing_Date, 111) Processing_Date, --資料處理日期
C01.Product_Code, --產品
C01.Reference_Nbr, --案件編號/帳號
C01.Current_Rating_Code, --風險區隔
C01.Ori_Amount, --原始購買金額
C01.Exposure, --曝險額
D02.Principal, --剩餘本金餘額
D02.Interest_Receivable, --應收利息
C01.Actual_Year_To_Maturity, --合約到期年限
C01.Duration_Year, --估計存續期間_年
C01.Remaining_Month, --估計存續期間_月
C01.Current_Int_Rate, --合约利率/產品利率
C01.EIR, --有效利率
A02.Delinquent_Days, --遲延天數
A02.Restructure_Ind, --是否為紓困名單
A02.Collateral_Legal_Action_Ind, --是否為假扣押名單
A02.Obj_Impair, --是否有客觀減損證據
A02.Bad_Debt_Ind, --是否屬呆帳
D02.Loan_Risk_Type, --風險三級分類
D02.Collection_Ind, --是否為催收款
D02.Loan_Type, --金融業務分類
D02.NO34RCV, --回收率的區隔
D02.Interest, --實收利息
C01.Impairment_Stage, --減損階段
D02.Current_LGD, --違約損失率
D02.PD, --第一年違約率
D02.Y1_EL, --一年期預期信用損失
D02.Lifetime_EL, --存續期間預期信用損失
D02.EL, --最終預期信用損失
D02.Principal_EL, --累計減損_本金
D02.Interest_Receivable_EL --累計減損_利息
from (select * from  EL_Data_In
where Report_Date = @Report_Date ) C01
join (select * from Loan_Account_Info
where Report_Date = @Report_Date) A02
on C01.Reference_Nbr = A02.Reference_Nbr
join (select * from IFRS9_Loan_Report
where Report_Date = @Report_Date) D02
on C01.Reference_Nbr = D02.Reference_Nbr
Order by Reference_Nbr
                        ";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    string reportDate = parms.Where(x => x.key == "Report_Date").FirstOrDefault()?.value ?? string.Empty;

                    cmd.Parameters.Add(new SqlParameter("Report_Date", reportDate));

                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);
                }
            }

            return resultsTable;
        }
    }
}