using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_4_2 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;

                sql = @"
                            WITH D02_1 AS
                            (
		                        SELECT A.* 
                                FROM 
                                ( 
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Final
                                ) A
                                JOIN
                                (
                                    SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                                ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                            ),
                            D02_2 AS
                            (
		                        SELECT A.* 
                                FROM 
                                ( 
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Initial
                                ) A
                                JOIN
                                (
                                    SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                                ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                            ),
                            D02_3 AS
                            (
	                            SELECT A.FLOWID,
		                               A.Product_Code,
                                       CASE WHEN LEN(A.Reference_Nbr) >= 8
                                            THEN SUBSTRING(A.Reference_Nbr,0,5) + '***' + SUBSTRING(A.Reference_Nbr,8,LEN(A.Reference_Nbr)-7)
                                            WHEN LEN(A.Reference_Nbr) >= 5
                                            THEN SUBSTRING(A.Reference_Nbr,0,5) + REPLICATE('*' , LEN(A.Reference_Nbr) - 4)
                                            ELSE A.Reference_Nbr
                                       END AS Reference_Nbr, 
                                       A.Collection_Ind,
		                               ROUND(A.Principal,0) Principal_1,
		                               A.Interest_Receivable Interest_Receivable_1,
		                               A.Exposure Exposure_1,
		                               A.Impairment_Stage Impairment_Stage_1,
		                               CAST(A.EL AS decimal(30,10)) EL_1,
		                               CAST(A.Interest_Receivable_EL AS decimal(30,10)) Interest_Receivable_EL_1,
		                               CAST(A.Principal_EL AS decimal(30,10)) Principal_EL_1,
		                               ISNULL(B.Principal,0) Principal_2,
		                               ISNULL(B.Interest_Receivable,0) Interest_Receivable_2,
		                               ROUND(ISNULL(B.Exposure,0),0) Exposure_2,
		                               ISNULL(B.Impairment_Stage,0) Impairment_Stage_2,
		                               CAST(ISNULL(B.EL,0) AS decimal(30,10)) EL_2,
		                               CAST(ISNULL(B.Interest_Receivable_EL,0) AS decimal(30,10)) Interest_Receivable_EL_2,
		                               CAST(ISNULL(B.Principal_EL,0) AS decimal(30,10)) Principal_EL_2,
		                               ROUND((A.Principal - ISNULL(B.Principal,0)),0) EAD
	                            FROM D02_1 A
	                            LEFT JOIN D02_2 B 
                                ON A.Reference_Nbr = B.Reference_Nbr
				              AND A.Product_Code = B.Product_Code 
                            )
                             SELECT *,
                                    CAST((CASE WHEN EAD > 0 THEN (EAD /Principal_1*
                                    (CASE WHEN (Collection_Ind = 'Y')
									     THEN (Principal_EL_1 + Interest_Receivable_EL_1)
										 ELSE Principal_EL_1
									 END)
                                    ) ELSE 0 END) AS decimal(30,10)) This_Phase_EL,
                                CASE WHEN Collection_Ind = 'Y'
                                     THEN
	                               (CASE WHEN (EAD <> Principal_1) AND (EAD>=0) THEN CAST( ((Principal_2+Interest_Receivable_2) / (Principal_1+Interest_Receivable_1) * (Principal_EL_1+Interest_Receivable_EL_1) - (Principal_EL_2+Interest_Receivable_EL_2)) AS decimal(30,10))
	                                     WHEN (EAD <> Principal_1) AND (EAD<0) THEN ((Principal_EL_1+Interest_Receivable_EL_1) - (Principal_1+Interest_Receivable_1) / (Principal_2+Interest_Receivable_2) * (Principal_EL_2+Interest_Receivable_EL_2))
			                            WHEN (EAD = Principal_1) THEN 0 END)
                                     ELSE
	                               (CASE WHEN (EAD <> Principal_1) AND (EAD>=0) THEN CAST( (Principal_2 / Principal_1 * Principal_EL_1 - Principal_EL_2) AS decimal(30,10))
	                                     WHEN (EAD <> Principal_1) AND (EAD<0) THEN (Principal_EL_1 - Principal_1/ Principal_2 * Principal_EL_2)
			                            WHEN (EAD = Principal_1) THEN 0 END)
                                    END AS This_Phase_EL_2
                            FROM D02_3
                      ";

                using (var cmd = new SqlCommand(sql, conn))
                {
                    //string reportDate = parms.Where(x => x.key == "Report_Date").FirstOrDefault()?.value ?? string.Empty;
                    string Group_Product_Code = parms.Where(x => x.key == "Group_Product_Code").FirstOrDefault()?.value ?? string.Empty;
                    string Initial = parms.Where(x => x.key == "Initial").FirstOrDefault()?.value ?? string.Empty;
                    string Final = parms.Where(x => x.key == "Final").FirstOrDefault()?.value ?? string.Empty;
                    //cmd.Parameters.Add(new SqlParameter("TheEndOfLastYearReport_Date", DateTime.Parse(reportDate).AddYears(-1).ToString("yyyy/12/31")));
                    //cmd.Parameters.Add(new SqlParameter("Report_Date", reportDate));
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
                    var _Group_Product_Code = new SqlParameter();
                    _Group_Product_Code.SqlDbType = SqlDbType.VarChar;
                    _Group_Product_Code.ParameterName = "@Group_Product_Code";
                    _Group_Product_Code.Value = Group_Product_Code;
                    _Group_Product_Code.Direction = ParameterDirection.Input;
                    cmd.Parameters.Add(_Initial); //期初 
                    cmd.Parameters.Add(_Final); //期末
                    cmd.Parameters.Add(_Group_Product_Code);
                    //cmd.Parameters.Add(new SqlParameter("Initial", Initial));
                    //cmd.Parameters.Add(new SqlParameter("Final", Final));
                    //cmd.Parameters.Add(new SqlParameter("Group_Product_Code", Group_Product_Code));

                    conn.Open();
                    var adapter = new SqlDataAdapter(cmd);
                    adapter.Fill(resultsTable);
                }
            }

            return resultsTable;
        }
    }
}