using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_2 : ReportData
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
                                    AND Reference_Nbr LIKE '[A-Za-z]%'
                                ) A
                                JOIN
                                (
                                    SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                                ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                            ),
                            NorthArea AS
                            (
	                            SELECT ('Stage ' + Impairment_Stage) Impairment_Stage,
		                               ISNULL(SUM(Exposure),0) Exposure,
		                               ISNULL(SUM(CAST(Y1_EL AS decimal)),0) Y1_EL,
		                               ISNULL(SUM(CAST(Lifetime_EL AS decimal)),0) Lifetime_EL,
		                               ISNULL(SUM(CAST(EL AS decimal)),0) EL
	                            FROM D02 WHERE NO34RCV = '北區'
	                            GROUP BY Impairment_Stage
                            ),
                            NotNorthArea AS
                            (
	                            SELECT ('Stage ' + Impairment_Stage) Impairment_Stage,
		                               ISNULL(SUM(Exposure),0) Exposure,
		                               ISNULL(SUM(CAST(Y1_EL AS decimal)),0) Y1_EL,
		                               ISNULL(SUM(CAST(Lifetime_EL AS decimal)),0) Lifetime_EL,
		                               ISNULL(SUM(CAST(EL AS decimal)),0) EL
	                            FROM D02 WHERE NO34RCV <> '北區'
	                            GROUP BY Impairment_Stage
                            ),
                            TEMP AS
                            (
	                            SELECT A.Impairment_Stage,
	                                   A.NO34RCV,
		                               A.SortNumber,
		                               ISNULL(B.Exposure,0) Exposure,
		                               ISNULL(B.Y1_EL,0) Y1_EL,
		                               ISNULL(B.Lifetime_EL,0) Lifetime_EL,
		                               ISNULL(B.EL,0) EL
	                            FROM
	                            (
		                            SELECT 'Stage 1' Impairment_Stage,'北區' NO34RCV, (1) SortNumber
		                            UNION
		                            SELECT 'Stage 2' Impairment_Stage,'北區' NO34RCV, (1) SortNumber
		                            UNION
		                            SELECT 'Stage 3' Impairment_Stage,'北區' NO34RCV, (1) SortNumber
		                            UNION
		                            SELECT 'Stage 1' Impairment_Stage,'非北區' NO34RCV, (2) SortNumber
		                            UNION
		                            SELECT 'Stage 2' Impairment_Stage,'非北區' NO34RCV, (2) SortNumber
		                            UNION
		                            SELECT 'Stage 3' Impairment_Stage,'非北區' NO34RCV, (2) SortNumber
	                            ) A
	                            LEFT JOIN
	                            (
		                            SELECT *,'北區' NO34RCV, (1) SortNumber FROM NorthArea
		                            UNION 
		                            SELECT *,'非北區' NO34RCV, (2) SortNumber FROM NotNorthArea
	                            ) B ON A.Impairment_Stage = B.Impairment_Stage AND A.NO34RCV = B.NO34RCV	
                            )
                            SELECT A.* 
                            FROM 
                            (
                                SELECT * FROM TEMP
                                UNION 
	                            SELECT Impairment_Stage,
                                       '小計' NO34RCV,
                                       (3) SortNumber,
		                               ISNULL(SUM(Exposure),0) Exposure,
		                               ISNULL(SUM(Y1_EL),0) Y1_EL,
		                               ISNULL(SUM(Lifetime_EL),0) Lifetime_EL,
		                               ISNULL(SUM(EL),0) EL
	                            FROM TEMP
	                            GROUP BY Impairment_Stage
                            ) A  ORDER BY Impairment_Stage,SortNumber
                        ";

                string sql2 = string.Empty;
                sql2 += $@"
                            WITH D02 AS
                            (
		                        SELECT A.* 
                                FROM 
                                ( 
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Report_Date
                                    AND Reference_Nbr NOT LIKE '[A-Za-z]%'
                                ) A
                                JOIN
                                (
                                    SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                                ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                            ),
                            StageData AS
                            (
	                            SELECT ('企金-Stage ' + Impairment_Stage) Impairment_Stage,
		                               ISNULL(SUM(Exposure),0) Exposure,
		                               ISNULL(SUM(CAST(Y1_EL AS decimal)),0) Y1_EL,
		                               ISNULL(SUM(CAST(Lifetime_EL AS decimal)),0) Lifetime_EL,
		                               ISNULL(SUM(CAST(EL AS decimal)),0) EL
	                            FROM D02
	                            GROUP BY Impairment_Stage
                            ),
                            TEMP AS
                            (
	                            SELECT A.Impairment_Stage,
	                                   A.NO34RCV,
		                               A.SortNumber,
		                               ISNULL(B.Exposure,0) Exposure,
		                               ISNULL(B.Y1_EL,0) Y1_EL,
		                               ISNULL(B.Lifetime_EL,0) Lifetime_EL,
		                               ISNULL(B.EL,0) EL
	                            FROM
	                            (
		                            SELECT '企金-Stage 1' Impairment_Stage,'' NO34RCV, (1) SortNumber
		                            UNION
		                            SELECT '企金-Stage 2' Impairment_Stage,'' NO34RCV, (1) SortNumber
		                            UNION
		                            SELECT '企金-Stage 3' Impairment_Stage,'' NO34RCV, (1) SortNumber
	                            ) A
	                            LEFT JOIN
	                            (
		                            SELECT *,'' NO34RCV, (1) SortNumber FROM StageData
	                            ) B ON A.Impairment_Stage = B.Impairment_Stage AND A.NO34RCV = B.NO34RCV
                            )
                            SELECT * FROM TEMP ORDER BY Impairment_Stage,SortNumber                        
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

                    cmd.CommandText = sql2;
                    DataTable dt2 = new DataTable();
                    dt2.Load(cmd.ExecuteReader());

                    resultsTable.Tables.Add(dt2);
                }
            }

            return resultsTable;
        }
    }
}