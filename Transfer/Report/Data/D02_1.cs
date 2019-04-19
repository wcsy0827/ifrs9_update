using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_1 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;

                sql += $@"
                            WITH D02 AS
                            (
		                        SELECT A.* 
                                FROM 
                                ( 
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Report_Date
                                ) A
                                JOIN
                                (
                                    SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                                ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                            ),
                            Collection_Ind_Y AS
                            (
	                            SELECT *
	                            FROM D02
	                            WHERE Collection_Ind = 'Y'
                            ),
                            Collection_Ind_NotY_English AS
                            (
	                            SELECT *
	                            FROM D02
	                            WHERE Collection_Ind <> 'Y'
	                                  AND Reference_Nbr LIKE '[A-Za-z]%'
                            ),
                            Collection_Ind_NotY_NotEnglish AS
                            (
	                            SELECT *
	                            FROM D02
	                            WHERE Collection_Ind <> 'Y'
	                                    AND Reference_Nbr NOT LIKE '[A-Za-z]%'
                            ),
                            TEMP AS
                            (
								SELECT '放款轉列之催收款' Collection_Ind_Name,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_1_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_1_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_1_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_2_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_2_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_2_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_Y WHERE Impairment_Stage = '3') Stage3,
										(SELECT ISNULL(SUM(CAST(ISNULL(EL,0) AS decimal)),0)  FROM Collection_Ind_Y) EL,
										(1) SortNumber
								UNION 
								SELECT '貼現及放款' Collection_Ind_Name,
										NULL Loan_Risk_Type_1_1,
										NULL Loan_Risk_Type_1_2,
										NULL Loan_Risk_Type_1_3,
										NULL Loan_Risk_Type_2_1,
										NULL Loan_Risk_Type_2_2,
										NULL Loan_Risk_Type_2_3,
										(0) Stage3,
										(0) EL,
										(2) SortNumber
								UNION
								SELECT '   -個人金融業務' Collection_Ind_Name,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_1_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_1_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_1_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_2_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_2_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_2_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_English WHERE Impairment_Stage = '3') Stage3,
										(SELECT ISNULL(SUM(CAST(ISNULL(EL,0) AS decimal)),0) FROM Collection_Ind_NotY_English) EL,
										(3) SortNumber
								UNION
								SELECT '   -法人金融業務' Collection_Ind_Name,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_1_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_1_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '1' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_1_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '低風險') Loan_Risk_Type_2_1,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '中風險') Loan_Risk_Type_2_2,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '2' AND Loan_Risk_Type = '高風險') Loan_Risk_Type_2_3,
										(SELECT ISNULL(SUM(Principal),0) FROM Collection_Ind_NotY_NotEnglish WHERE Impairment_Stage = '3') Stage3,
										(SELECT ISNULL(SUM(CAST(ISNULL(EL,0) AS decimal)),0) FROM Collection_Ind_NotY_NotEnglish) EL,
										(4) SortNumber
                            )
                            SELECT * FROM TEMP ORDER BY SortNumber
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