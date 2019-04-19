using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_3 : ReportData
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
                            Individual AS
                            (
	                            SELECT Impairment_Stage,
								       ISNULL(SUM(Principal),0) Principal,
								       ISNULL(SUM(CAST(EL AS decimal)),0) EL,
									   ISNULL(SUM(EIR),0) EIR,
									   ISNULL(SUM(Principal*EIR),0) PrincipalEIR,
									   ISNULL(SUM(CAST(Interest AS decimal)),0) Interest
	                            FROM D02
	                            WHERE Reference_Nbr LIKE '[A-Za-z]%'
								GROUP BY Impairment_Stage
                            ),
                            Enterprise AS
                            (
	                            SELECT Impairment_Stage,
								       ISNULL(SUM(Principal),0) Principal,
								       ISNULL(SUM(CAST(EL AS decimal)),0) EL,
									   ISNULL(SUM(EIR),0) EIR,
									   ISNULL(SUM(Principal*EIR),0) PrincipalEIR,
									   ISNULL(SUM(CAST(Interest AS decimal)),0) Interest
	                            FROM D02
	                            WHERE Reference_Nbr NOT LIKE '[A-Za-z]%'
								GROUP BY Impairment_Stage
                            ),
                            TEMP AS
                            (
	                            SELECT A.*,
                                       ISNULL(B.Principal,0) Principal,
                                       ISNULL(B.EL,0) EL,
                                       ISNULL(B.BookValue,0) BookValue,
                                       ISNULL(B.WeightedAverageInterestRate,0) WeightedAverageInterestRate,
                                       ISNULL(CAST((B.BookValue*B.WeightedAverageInterestRate/12) AS decimal),0) InterestIncome,
                                       ISNULL(B.Interest,0) Interest,
                                       ISNULL(CAST(((B.BookValue*B.WeightedAverageInterestRate/12) - B.Interest) AS decimal),0) AllowanceForImpairmentAmortization
								FROM
								(
		                            SELECT '個人-Stage 1' Combination, (1) SortNumber
		                            UNION
		                            SELECT '個人-Stage 2' Combination, (2) SortNumber
		                            UNION
		                            SELECT '個人-Stage 3' Combination, (3) SortNumber
		                            UNION
		                            SELECT '企金-Stage 1' Combination, (4) SortNumber
		                            UNION
		                            SELECT '企金-Stage 2' Combination, (5) SortNumber
		                            UNION
		                            SELECT '企金-Stage 3' Combination, (6) SortNumber
                                ) A
                                LEFT JOIN
                                (
									SELECT ('個人-Stage ' + Impairment_Stage) Combination,
									       *,
										   (Principal - EL) BookValue,
										   (CASE WHEN Principal <> 0 THEN (PrincipalEIR/Principal) ELSE 0 END) WeightedAverageInterestRate,
										   (1) SortNumber
									FROM Individual 
									UNION 
									SELECT ('企金-Stage ' + Impairment_Stage) Combination,
									       *,
										   (Principal - EL) BookValue,
										   (CASE WHEN Principal <> 0 THEN (PrincipalEIR/Principal) ELSE 0 END) WeightedAverageInterestRate,
										   (2) SortNumber
									FROM Enterprise
                                ) B ON A.Combination = B.Combination
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