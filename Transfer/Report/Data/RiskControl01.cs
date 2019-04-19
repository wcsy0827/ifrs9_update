using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using Transfer.Models;
using Transfer.Utility;

namespace Transfer.Report.Data
{
    public class RiskControl01 : ReportData
    {
        public override DataSet GetData(List<reportParm> parms)
        {
            var resultsTable = new DataSet();
            using (var conn = new SqlConnection(defaultConnection))
            {
                string sql = string.Empty;
                sql += $@"
SET NOCOUNT ON;

DECLARE @#SomeTable TABLE (ISSUER VARCHAR(50),Change_Memo varchar(30));

with Ver AS
(
   SELECT TOP 1 Version 
   FROM Bond_Rating_Warning
   WHERE Report_Date = @reportDate
   ORDER BY Version DESC
)
INSERT into @#SomeTable 
(
ISSUER,
Change_Memo
)
select Distinct ISSUER,Change_Memo
from Bond_Rating_Warning,Ver
where Report_Date = @reportDate
and Bond_Rating_Warning.Version = Ver.Version
and Wraming_1_Ind = 'Y'
and New_Ind = 'Y'
AND Product_Group_1 is not null

DECLARE @AllString  varchar(2000) = '',
        @New varchar(20) = '為新投資',
		@NewString varchar(3000) = '',
		@Increase varchar(20) = '係因增加投資',
		@IncreaseString varchar(3000) = '',
		@Decrease varchar(20) = '係因降評所致',
		@DecreaseString varchar(3000) = '',
		@Diff_Over varchar(30) = '於六個月內降評三級以上(含)。',
		@Diff_OverString varchar(3000) = '',
        @ISSUER  varchar(50),
		@Change_Memo varchar(30);

    --1. 新投資 => 為新投資
	--2. 增加307 => 係因增加投資
	--3. BB下降 => 係因降評所致
    --本月新增預警，BNP PARIBAS係因增加投資，BECTON DICKINSO為新投資；Mattel Inc於六個月內降評三級以上(含)。
	--本月新增預警，Ensco PLC、Mattel Inc、Pride Intl LLC係因降評所致；Deutsche BK AG、Gaz Capital SA及Vale Overseas係因增加投資所致。

WHILE EXISTS (SELECT * FROM  @#SomeTable)
BEGIN
    SELECT TOP 1 @ISSUER = ISSUER,@Change_Memo = Change_Memo FROM @#SomeTable;

	if( @Change_Memo = '新投資')
	Set @NewString += (@ISSUER + ',');
	if( @Change_Memo like '增加%')
    SET @IncreaseString += (@ISSUER + ',');
	if( @Change_Memo like '%下降')
    SET @DecreaseString += (@ISSUER + ',');

    DELETE  @#SomeTable Where ISSUER = @ISSUER AND Change_Memo = @Change_Memo;
END;

with Ver AS
(
   SELECT TOP 1 Version 
   FROM Bond_Rating_Warning
   WHERE Report_Date = @reportDate
   ORDER BY Version DESC
)
INSERT into @#SomeTable 
(
ISSUER
)
select Distinct ISSUER 
from Bond_Rating_Warning,Ver
where Report_Date = @reportDate
and Bond_Rating_Warning.Version = Ver.Version
and Rating_diff_Over_Ind='Y'
and Wraming_2_Ind ='Y'
AND Product_Group_1 is not null

WHILE EXISTS (SELECT * FROM  @#SomeTable)
BEGIN
    SELECT TOP 1 @ISSUER = ISSUER FROM @#SomeTable;

	Set @Diff_OverString += (@ISSUER + ',');

    DELETE  @#SomeTable Where ISSUER = @ISSUER ;
END;

WITH AllString AS
(
   select CASE WHEN LEN(@NewString) > 0
               THEN SUBSTRING(@NewString, 1, (LEN(@NewString) -1)) + @New + ';'
			   ELSE ''
			   END AS NEW,
          CASE WHEN LEN(@IncreaseString) > 0
               THEN SUBSTRING(@IncreaseString, 1, (LEN(@IncreaseString) -1)) + @Increase + ';'
			   ELSE ''
			   END AS Increase,
		  CASE WHEN LEN(@DecreaseString) > 0
               THEN SUBSTRING(@DecreaseString, 1, (LEN(@DecreaseString) -1)) + @Decrease + ';'
			   ELSE ''
			   END AS Decrease,
		  CASE WHEN LEN(@Diff_OverString) > 0
		       THEN SUBSTRING(@Diff_OverString, 1, (LEN(@Diff_OverString) -1)) + @Diff_Over 
			   ELSE '無六個月內降評三級以上(含)。'
			   END AS Diff_Over
),
textValue as
(
  Select '本月新增:' + 
  CASE WHEN LEN(AllString.NEW + AllString.Increase + AllString.Decrease) > 0
            THEN AllString.NEW + AllString.Increase + AllString.Decrease
			ELSE '無'
	END
    + ' , ' + AllString.Diff_Over AS allstr from AllString
),
Ver AS
(
   SELECT MAX(Version) AS version
   FROM Bond_Rating_Warning
   WHERE Report_Date = @reportDate
)
Select 
CASE WHEN D67.Wraming_1_Ind = 'Y' and D67.Bond_Area = '國外'
     THEN 'BB-'
	 WHEN D67.Wraming_1_Ind = 'Y' and D67.Bond_Area = '國內'
	 THEN 'twBB-'
	 WHEN D67.Wraming_2_Ind = 'Y' and D67.Bond_Area = '國外'
	 THEN 'BB-'
	 WHEN D67.Wraming_2_Ind = 'Y' and D67.Bond_Area = '國內' 
	 THEN 'twBB-'
END + '以下(含)' AS GRADE_Warning, --預警標準
D67.Bond_Number AS Bond_Number, -- CUSIP/ISIN/Price Symbol (債券編號)
D67.ISSUER AS ISSUER, --發行者
ISNULL(D67.Rating_Worse,'Non-rated')  As Rating_Worse, --評等
D67.Amort_Amt_Tw AS Amort_Amt_Tw, --金額
'是' AS Wraming, --是否預警
--D67.Wraming_1_Ind AS Wraming_1_Ind, --是否評等過低預警?
D67.Observation_Month AS Observation_Month, --連續追蹤月數
--D67.Product_Group_1 AS Product_Group_1, --資產群組別(提供信用風險資產減損預警彙總表使用)
--D67.Product_Group_2 AS Product_Group_2, --投資標的資產群組別(提供投資標的信用風險預警彙總表使用)
--D67.Rating_diff_Over_Ind AS Rating_diff_Over_Ind, --是否最近6個月內最近信評與原始信評的等級差異降評三級以上(含) 
REPLACE(REPLACE(D67.Change_Memo,'增加','增'),'下降','↓') AS Change_Memo, --本月變動
CASE WHEN D67.Product_Group_1 = '公債'
     THEN 1
	 WHEN D67.Product_Group_1 = '金融債&可贖回零息金融債&NCD'
	 THEN 2
	 WHEN D67.Product_Group_1 = '公司債'
	 THEN 3
	 WHEN D67.Product_Group_1 = '結構式商品'
	 THEN 4
	 WHEN D67.Product_Group_1 = '證券化商品'
	 THEN 5
ELSE 0 
END AS Status1, --
CASE WHEN D67.New_Ind = 'Y'
     THEN 1 --新增預警統計表
     ELSE 2 --持續監控預警統計表
END AS Status2
From 
Bond_Rating_Warning D67, Ver 
Where Report_Date =  @reportDate
AND D67.Version = Ver.Version
AND D67.Wraming_1_Ind = 'Y'
AND D67.Product_Group_1 is not null
UNION ALL
Select 
CASE WHEN D67.Wraming_1_Ind = 'Y' and D67.Bond_Area = '國外'
     THEN 'BB-'
	 WHEN D67.Wraming_1_Ind = 'Y' and D67.Bond_Area = '國內'
	 THEN 'twBB-'
	 WHEN D67.Wraming_2_Ind = 'Y' and D67.Bond_Area = '國外'
	 THEN 'BB-'
	 WHEN D67.Wraming_2_Ind = 'Y' and D67.Bond_Area = '國內' 
	 THEN 'twBB-'
END + '以下(含)' AS GRADE_Warning, --預警標準
D67.Bond_Number AS Bond_Number, -- CUSIP/ISIN/Price Symbol (債券編號)
D67.ISSUER AS ISSUER, --發行者
ISNULL(D67.Rating_Worse,'Non-rated')  As Rating_Worse, --評等
D67.Amort_Amt_Tw AS Amort_Amt_Tw, --金額
'是' AS Wraming, --是否預警
--D67.Wraming_1_Ind AS Wraming_1_Ind, --是否評等過低預警?
null AS Observation_Month, --連續追蹤月數
--D67.Product_Group_1 AS Product_Group_1, --資產群組別(提供信用風險資產減損預警彙總表使用)
--D67.Product_Group_2 AS Product_Group_2, --投資標的資產群組別(提供投資標的信用風險預警彙總表使用)
--D67.Rating_diff_Over_Ind AS Rating_diff_Over_Ind, --是否最近6個月內最近信評與原始信評的等級差異降評三級以上(含) 
REPLACE(REPLACE(D67.Rating_Change_Memo,'增加','增'),'下降','↓') AS Change_Memo, --本月變動
CASE WHEN D67.Product_Group_1 = '公債'
     THEN 1
	 WHEN D67.Product_Group_1 = '金融債&可贖回零息金融債&NCD'
	 THEN 2
	 WHEN D67.Product_Group_1 = '公司債'
	 THEN 3
	 WHEN D67.Product_Group_1 = '結構式商品'
	 THEN 4
	 WHEN D67.Product_Group_1 = '證券化商品'
	 THEN 5
ELSE 0 
END AS Status1, --
3 AS Status2 --六個月內降評三級以上(含)
From 
Bond_Rating_Warning D67, Ver 
Where Report_Date =  @reportDate
AND D67.Version = Ver.Version
AND D67.Wraming_2_Ind = 'Y'
AND D67.Rating_diff_Over_Ind = 'Y'
AND D67.Product_Group_1 is not null
UNION ALL
select
NULL,
NULL,
NULL,
NULL,
NULL,
NULL,
NULL,
allstr ,
-1,
-1
from textValue
order by Status2,Status1,ISSUER,Bond_Number
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