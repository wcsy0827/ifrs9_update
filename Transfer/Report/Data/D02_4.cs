using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class D02_4 : ReportData
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
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Initial
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
                                    SELECT * FROM IFRS9_Loan_Report WHERE Report_Date = @Final
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
		                               A.Reference_Nbr,
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
		                               ROUND((ISNULL(B.Principal,0) - A.Principal),0) EAD
	                            FROM D02_1 A
	                            LEFT JOIN D02_2 B 
                                ON A.Reference_Nbr = B.Reference_Nbr 
				              AND A.Product_Code = B.Product_Code 
                            ),
                            D02_4 AS
                            (
                                SELECT *,
                                    CAST((CASE WHEN EAD < 0 
                                    THEN (
                                    (CASE WHEN Collection_Ind = 'Y'
                                          THEN (EAD - Interest_Receivable_1)
                                          ELSE EAD
                                          END
                                    )
                                    / 
                                    (CASE WHEN Collection_Ind = 'Y'
                                          THEN (Principal_1 + Interest_Receivable_1)
                                          ELSE Principal_1
                                          END
                                    )
                                    * 
                                    (CASE WHEN (Collection_Ind = 'Y')
									     THEN (Principal_EL_1 + Interest_Receivable_EL_1)
										 ELSE Principal_EL_1
									 END)
                                    ) ELSE 0 END) AS decimal(30,10)) This_Phase_EL,
	                               (
	                                        CASE WHEN (Impairment_Stage_1 <> Impairment_Stage_2) AND (EAD<=0) THEN CAST(((Principal_2 * -1) / Principal_1  * Principal_EL_1) AS decimal(30,10))	                                             
                                                  WHEN (Impairment_Stage_1 <> Impairment_Stage_2) AND (EAD>0) THEN (Principal_EL_1*-1)
			                                    WHEN (Impairment_Stage_1 = Impairment_Stage_2) THEN 0 END
	                               ) Re_Classify_EL
                                FROM D02_3
                            ),
                            D02_A AS
                            (
                               select 
                               SUM(case when Impairment_Stage_1 = '1' AND Collection_Ind = 'Y'
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '1'
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS A1,
                               SUM(case when Impairment_Stage_1 = '1' AND Impairment_Stage_2 = '2'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS A2,
                               SUM(case when Impairment_Stage_1 = '1' AND Impairment_Stage_2 = '3'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS A3,
                               SUM(case when Impairment_Stage_1 = '1'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS A5,
                               SUM(case when Impairment_Stage_1 = '2' AND Collection_Ind = 'Y'
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '2'
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS B1,     
                               SUM(case when Impairment_Stage_1 = '2' AND Impairment_Stage_2 = '3'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS B3,
                               SUM(case when Impairment_Stage_1 = '2' AND Impairment_Stage_2 = '1'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS B4,       
                               SUM(case when Impairment_Stage_1 = '2'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS B5, 
                               SUM(case when Impairment_Stage_1 = '3' AND Collection_Ind = 'Y'
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '3'
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS C1,  
                               SUM(case when Impairment_Stage_1 = '3' AND Impairment_Stage_2 = '2'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS C2,
                               SUM(case when Impairment_Stage_1 = '3' AND Impairment_Stage_2 = '1'
                                        then CAST(ISNULL(Re_Classify_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS C4,
                               SUM(case when Impairment_Stage_1 = '3'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS C5                                                                                                                                                                                                                                                                  
                               from D02_4                               
                            ),
                            D02_5 AS
                            (
	                            SELECT A.FLOWID,
		                               A.Product_Code,
		                               A.Reference_Nbr,
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
	                            FROM D02_2 A
	                            LEFT JOIN D02_1 B 
                                ON A.Reference_Nbr = B.Reference_Nbr 
				              AND A.Product_Code = B.Product_Code
                            ),
                            D02_6 AS
                            (
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
                                FROM D02_5
                            ),
                            D02_B As
                            (
                                 SELECT  
                                 SUM(case when Impairment_Stage_1 = '1'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS A6,
                                 SUM(case when Impairment_Stage_1 = '1'
                                        then CAST(ISNULL(This_Phase_EL_2,0) AS decimal(30,10))
                                        else 0
                                        end) AS A7, 
                                 SUM(case when Impairment_Stage_1 = '1' AND Collection_Ind = 'Y'
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '1'
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS A8, 
                                 SUM(case when Impairment_Stage_1 = '2'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS B6,  
                                 SUM(case when Impairment_Stage_1 = '2'
                                        then CAST(ISNULL(This_Phase_EL_2,0) AS decimal(30,10))
                                        else 0
                                        end) AS B7,
                                 SUM(case when Impairment_Stage_1 = '2' AND Collection_Ind = 'Y' 
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '2'
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS B8,
                                 SUM(case when Impairment_Stage_1 = '3'
                                        then CAST(ISNULL(This_Phase_EL,0) AS decimal(30,10))
                                        else 0
                                        end) AS C6,  
                                 SUM(case when Impairment_Stage_1 = '3'
                                        then CAST(ISNULL(This_Phase_EL_2,0) AS decimal(30,10))
                                        else 0
                                        end) AS C7,
                                 SUM(case when Impairment_Stage_1 = '3'  AND Collection_Ind = 'Y' 
                                        then CAST(ISNULL(Principal_EL_1,0) + ISNULL(Interest_Receivable_EL_1,0) AS decimal(30,10))
                                        when Impairment_Stage_1 = '3' 
                                        then CAST(ISNULL(Principal_EL_1,0) AS decimal(30,10))
                                        else 0
                                        end) AS C8                                                                                                                                                                                                                                                 
                                 FROM D02_6
                            )
                            SELECT CAST(A1 AS decimal) A1,
                                   CAST(A2 AS decimal) A2,
                                   CAST(A3 AS decimal) A3,
                                   CAST(((B4 + C4)*-1) AS decimal) AS A4,
                                   CAST(A5 AS decimal) A5,
                                   CAST(A6 AS decimal) A6,
                                   CAST(A7 AS decimal) A7,
                                   CAST(A8 AS decimal) A8,
                                   CAST(B1 AS decimal) B1,
                                   CAST(((A2 + C2)*-1) AS decimal) AS B2,
                                   CAST(B3 AS decimal) B3,
                                   CAST(B4 AS decimal) B4,
                                   CAST(B5 AS decimal) B5,
                                   CAST(B6 AS decimal) B6,
                                   CAST(B7 AS decimal) B7,
                                   CAST(B8 AS decimal) B8,
                                   CAST(C1 AS decimal) C1,
                                   CAST(C2 AS decimal) C2,
                                   CAST(((A3 + B3) *-1) AS decimal) AS C3,
                                   CAST(C4 AS decimal) C4,
                                   CAST(C5 AS decimal) C5,
                                   CAST(C6 AS decimal) C6,
                                   CAST(C7 AS decimal) C7,
                                   CAST(C8 AS decimal) C8,
                                   CAST((A1 + B1 + C1) AS decimal) AS D1,
                                   CAST((A2 + ((A2 + C2)*-1) + C2) AS decimal) AS D2,
                                   CAST((A3 + B3 + ((A3 + B3)*-1)) AS decimal) AS D3,
                                   CAST((((B4 + C4)*-1) + B4 + C4) AS decimal) AS D4,
                                   CAST((A5 + B5 + C5) AS decimal) AS D5,
                                   CAST((A6 + B6 + C6) AS decimal) AS D6,
                                   CAST((A7 + B7 + C7) AS decimal) AS D7
                            FROM D02_A,D02_B
";

                #region old sql
                //    string sqlold = @"
                //                WITH D02_1 AS
                //                (
                //              SELECT A.* 
                //                    FROM 
                //                    ( 
                //                        SELECT * FROM IFRS9_Loan_Report WHERE CONVERT(varchar(10),CONVERT(date,Report_Date),111) = @Initial
                //                    ) A
                //                    JOIN
                //                    (
                //                        SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                //                    ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                //                ),
                //                D02_2 AS
                //                (
                //              SELECT A.* 
                //                    FROM 
                //                    ( 
                //                        SELECT * FROM IFRS9_Loan_Report WHERE CONVERT(varchar(10),CONVERT(date,Report_Date),111) = @Final
                //                    ) A
                //                    JOIN
                //                    (
                //                        SELECT * FROM Flow_Info WHERE Group_Product_Code = @Group_Product_Code
                //                    ) B ON A.PRJID = B.PRJID AND A.FLOWID = B.FLOWID
                //                ),
                //                D02_3 AS
                //                (
                //                 SELECT A.FLOWID,
                //                     A.Product_Code,
                //                     A.Reference_Nbr,
                //                     A.Principal Principal_1,
                //                     A.Interest_Receivable Interest_Receivable_1,
                //                     A.Exposure Exposure_1,
                //                     A.Impairment_Stage Impairment_Stage_1,
                //                     CAST(A.EL AS decimal) EL_1,
                //                     CAST(A.Interest_Receivable_EL AS decimal) Interest_Receivable_EL_1,
                //                     CAST(A.Principal_EL AS decimal) Principal_EL_1,
                //                     B.Principal Principal_2,
                //                     B.Interest_Receivable Interest_Receivable_2,
                //                     B.Exposure Exposure_2,
                //                     B.Impairment_Stage Impairment_Stage_2,
                //                     CAST(B.EL AS decimal) EL_2,
                //                     CAST(B.Interest_Receivable_EL AS decimal) Interest_Receivable_EL_2,
                //                     CAST(B.Principal_EL AS decimal) Principal_EL_2,
                //                     (B.Principal - A.Principal) EAD
                //                 FROM D02_1 A
                //                 JOIN D02_2 B ON A.FLOWID = B.FLOWID 
                //                    AND A.Product_Code = B.Product_Code 
                //                    AND A.Reference_Nbr = B.Reference_Nbr
                //                ),
                //                D02_4 AS
                //                (
                //                    SELECT *,
                //                           CAST((CASE WHEN EAD < 0 THEN (CAST(EAD AS float)/Principal_1*Principal_EL_1) ELSE 0 END) AS decimal) This_Phase_EL,
                //                        CAST((
                //                              CASE WHEN (Impairment_Stage_1 <> Impairment_Stage_2) AND (EAD<=0) THEN (Principal_2*-1/CAST(Principal_1 AS float)*Principal_EL_1) 
                //                                   WHEN (Impairment_Stage_1 <> Impairment_Stage_2) AND (EAD>0) THEN (Principal_EL_1*-1)
                //                             WHEN (Impairment_Stage_1 = Impairment_Stage_2) THEN 0 END
                //                            ) AS decimal) Re_Classify_EL
                //                    FROM D02_3
                //                ),
                //                D02_5 AS
                //                (
                //                 SELECT B.FLOWID,
                //                     B.Product_Code,
                //                     B.Reference_Nbr,
                //                     B.Principal Principal_1,
                //                     B.Interest_Receivable Interest_Receivable_1,
                //                     B.Exposure Exposure_1,
                //                     B.Impairment_Stage Impairment_Stage_1,
                //                     CAST(B.EL AS decimal) EL_1,
                //                     CAST(B.Interest_Receivable_EL AS decimal) Interest_Receivable_EL_1,
                //                     CAST(B.Principal_EL AS decimal) Principal_EL_1,
                //                     A.Principal Principal_2,
                //                     A.Interest_Receivable Interest_Receivable_2,
                //                     A.Exposure Exposure_2,
                //                     A.Impairment_Stage Impairment_Stage_2,
                //                     CAST(A.EL AS decimal) EL_2,
                //                     CAST(A.Interest_Receivable_EL AS decimal) Interest_Receivable_EL_2,
                //                     CAST(A.Principal_EL AS decimal) Principal_EL_2,
                //                     (B.Principal - A.Principal) EAD
                //                 FROM D02_1 A
                //                 JOIN D02_2 B ON A.FLOWID = B.FLOWID 
                //                    AND A.Product_Code = B.Product_Code 
                //                    AND A.Reference_Nbr = B.Reference_Nbr
                //                ),
                //                D02_6 AS
                //                (
                //                    SELECT *,
                //                           CAST((CASE WHEN EAD > 0 THEN (CAST(EAD AS float)/Principal_1*Principal_EL_1) ELSE 0 END) AS decimal) This_Phase_EL,
                //                        CAST(
                //                              (CASE WHEN (EAD <> Principal_1) AND (EAD>=0) THEN (CAST(Principal_2 AS float)/Principal_1*Principal_EL_1-Principal_EL_2) 
                //                                   WHEN (EAD <> Principal_1) AND (EAD<0) THEN (Principal_EL_1-Principal_1/CAST(Principal_2 AS float)*Principal_EL_2)
                //                             WHEN (EAD = Principal_1) THEN 0 END)
                //                        AS decimal) This_Phase_EL_2
                //                    FROM D02_5
                //                ),
                //                D02_7 AS
                //                (
                //                    SELECT (SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '1') AS A1,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '1' AND Impairment_Stage_2 = '2') AS A2,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '1' AND Impairment_Stage_2 = '3') AS A3,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '1') AS A5,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '1') AS A6,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL_2 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '1') AS A7,
                //(SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '1') AS A8,
                //                           (SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '2') AS B1,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '2' AND Impairment_Stage_2 = '3') AS B3,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '2' AND Impairment_Stage_2 = '1') AS B4,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '2') AS B5,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '2') AS B6,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL_2 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '2') AS B7,
                //                           (SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '2') AS B8,
                //                           (SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '3') AS C1,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '3' AND Impairment_Stage_2 = '2') AS C2,
                //                           (SELECT ISNULL(SUM(CAST(Re_Classify_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '3' AND Impairment_Stage_2 = '1') AS C4,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_4 WHERE Impairment_Stage_1 = '3') AS C5,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '3') AS C6,
                //                           (SELECT ISNULL(SUM(CAST(This_Phase_EL_2 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '3') AS C7,
                //                           (SELECT ISNULL(SUM(CAST(Principal_EL_1 AS decimal)),0) FROM D02_6 WHERE Impairment_Stage_1 = '3') AS C8
                //                )
                //                SELECT A1,
                //                       A2,
                //                       A3,
                //                       ((B4 + C4)*-1) AS A4,
                //                       A5,
                //                       A6,
                //                       A7,
                //                       A8,
                //                       B1,
                //                       ((A2 + C2)*-1) AS B2,
                //                       B3,
                //                       B4,
                //                       B5,
                //                       B6,
                //                       B7,
                //                       B8,
                //                       C1,
                //                       C2,
                //                       ((A3 + B3) *-1) AS C3,
                //                       C4,
                //                       C5,
                //                       C6,
                //                       C7,
                //                       C8,
                //                       (A1 + B1 + C1) AS D1,
                //                       (A2 + ((A2 + C2)*-1) + C2) AS D2,
                //                       (A3 + B3 + ((A3 + B3)*-1)) AS D3,
                //                       (((B4 + C4)*-1) + B4 + C4) AS D4,
                //                       (A5 + B5 + C5) AS D5,
                //                       (A6 + B6 + C6) AS D6,
                //                       (A7 + B7 + C7) AS D7
                //                FROM D02_7
                //           ";
                #endregion


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