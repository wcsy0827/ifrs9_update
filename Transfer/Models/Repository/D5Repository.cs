using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class D5Repository : ID5Repository
    {
        #region 其他

        protected Common common
        {
            get;
            private set;
        }

        protected Common.User _UserInfo
        {
            get;
            private set;
        }

        private string C07reportDateFormate = "yyyy-MM-dd";

        public D5Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
        }
        #endregion 其他

        #region Get Data
        /// <summary>
        /// 查詢D53
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public Tuple<string, List<D53ViewModel>> GetD53()
        {
            string resultMessage = string.Empty;
            List<D53ViewModel> data = new List<D53ViewModel>();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    data = db.SMF_Info.AsNoTracking()
                        .AsEnumerable()
                        .Select(x => new D53ViewModel()
                        {
                            SMF_Code = x.SMF_Code,
                            SMF = x.SMF,
                            SMF_Desc = x.SMF_Desc,
                            Asset_type_Code = x.Asset_type_Code,
                            Asset_Type = x.Asset_Type
                        }).ToList();
                    if (!data.Any())
                        resultMessage = Message_Type.not_Find_Any.GetDescription(Table_Type.D53.tableNameGetDescription());
                }
            }
            catch (Exception ex)
            {
                resultMessage = ex.exceptionMessage();
            }
            return new Tuple<string, List<D53ViewModel>>(resultMessage, data);
        }

        #region D54查詢預計調整資料
        /// <summary>
        /// D54查詢預計調整資料
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public List<D54ViewModel> getD54InsertSearch(string dt)
        {
            List<D54ViewModel> results = new List<D54ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                DateTime reportDt = DateTime.MinValue;
                if (!DateTime.TryParse(dt, out reportDt))
                    return results;
                var A41s = db.Bond_Account_Info.AsNoTracking().Where(x => x.Report_Date == reportDt);
                if (!A41s.Any())
                    return results;
                var ver = A41s.Max(x => x.Version).Value;
                var D62s = db.Bond_Basic_Assessment.AsNoTracking().
                    Where(x => x.Report_Date == reportDt && x.Version == ver).ToList();
                var D62s_W = D62s.Where(x => x.Watch_IND == "Y").ToList();
                var D63Refs = db.Bond_Quantitative_Resource.AsNoTracking().
                    Where(x => x.Report_Date == reportDt &&
                               x.Version == ver &&
                               x.Result_Version_Confirm != null &&
                               x.Assessment_Result_Version == x.Result_Version_Confirm &&
                               x.Quantitative_Pass_Confirm == "Y")
                               .Select(x=>x.Reference_Nbr).Distinct().ToList();
                var D65s = db.Bond_Qualitative_Assessment.AsNoTracking()
                             .Where(x => x.Report_Date == reportDt &&
                                         x.Version == ver &&
                                         x.Result_Version_Confirm != null &&
                                         x.Assessment_Result_Version == x.Result_Version_Confirm &&
                                         x.Extra_Case == null);
                var D65Case = db.Bond_Qualitative_Assessment.AsNoTracking()
                             .Where(x => x.Report_Date == reportDt &&
                                         x.Version == ver &&
                                         x.Extra_Case == "Y");
                var D65CaseR = new List<string>();
                var D65CaseCheck = new List<Bond_Qualitative_Assessment>();
                if (D65Case.Any())
                {
                    D65CaseR = D65Case.GroupBy(x => x.Reference_Nbr).Select(x=>x.Key).ToList();
                    D65CaseCheck = D65Case.Where(x => x.Result_Version_Confirm != null &&
                                         x.Assessment_Result_Version == x.Result_Version_Confirm).ToList();
                }
                var D65Refs = D65s.Select(x=>x.Reference_Nbr).Distinct().ToList();
                if (D62s.Any())
                {
                    var Refs = D62s_W.GroupBy(x => x.Reference_Nbr).Select(x => x.Key).ToList();

                    if (Refs.Count != D63Refs.Count + D65Refs.Count) //有Reference_Nbr 沒有最後版本
                    {
                        var _Except = new List<string>();
                        _Except.AddRange(D63Refs);
                        _Except.AddRange(D65Refs);
                        Refs.Except(_Except)
                            .ToList() //抓取沒有最後版本的Reference_Nbr
                            .ForEach(
                            x =>
                            {
                                results.Add(new D54ViewModel()
                                {
                                    Reference_Nbr = x
                                });
                            });
                    }
                    else if (D65CaseR.Count != D65CaseCheck.Count())
                    {
                        D65CaseR.Except(D65CaseCheck.Select(x => x.Reference_Nbr)).ToList()
                            .ForEach(
                            x =>
                            {
                                results.Add(new D54ViewModel()
                                {
                                    Reference_Nbr = x
                                });
                            });
                    }
                    else
                    {
                        D65Refs.AddRange(D65CaseR);                      
                        var A41s2 = A41s.Where(x => x.Version == ver).ToList();
                        var D66s = db.Bond_Qualitative_Assessment_Result.AsNoTracking()
                            .Where(x => D65Refs.Contains(x.Reference_Nbr)).ToList();
                        var _data = D65s.Where(x => x.Quantitative_Pass_Confirm != "Y").ToList();
                        _data.AddRange(D65CaseCheck);
                        _data.ForEach(x =>
                        {
                            var _Assessment_Stage = D66s.First(z => z.Reference_Nbr == x.Reference_Nbr &&
                                                      z.Assessment_Result_Version == x.Assessment_Result_Version)
                                                      .Assessment_Stage;
                            var A41 = A41s2.First(y => y.Reference_Nbr == x.Reference_Nbr);
                            results.Add(new D54ViewModel()
                            {
                                Product_Code = formateProduct_Code(A41.Principal_Payment_Method_Code),
                                Reference_Nbr = x.Reference_Nbr,
                                Bond_Number = A41.Bond_Number,
                                Lots = A41.Lots,
                                Portfolio_Name = A41.Portfolio_Name,
                                Impairment_Stage = _Assessment_Stage == "2" ? "2" : "3"
                            });
                        });
                    }
                }
            }
            return results;
        }
        #endregion

        #region get D54 Group_Product_Code
        /// <summary>
        /// get D54 Group_Product_Code
        /// </summary>
        /// <returns></returns>
        public List<SelectOption> getD54GroupProductCode()
        {
            List<SelectOption> result = new List<SelectOption>();
            List<string> Product_Code = new List<string>() { "Bond_A", "Bond_B", "Bond_P" };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D05s = db.Group_Product_Code_Mapping.AsNoTracking()
                    .Where(x => Product_Code.Contains(x.Product_Code)).AsEnumerable()
                    .Select(x => x.Group_Product_Code).Distinct();
                result = db.Group_Product.AsNoTracking().Where(x => D05s.Contains(x.Group_Product_Code))
                    .AsEnumerable()
                    .Select(x => new SelectOption()
                    {
                        Value = x.Group_Product_Code,
                        Text = $@"{x.Group_Product_Code} ({x.Group_Product_Name})"
                    }).ToList();
            }
            return result;
        }
        #endregion

        #region get D54
        /// <summary>
        /// get D54
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="groupProductCode"></param>
        /// <param name="bondNumber"></param>
        /// <param name="refNumber"></param>
        /// <returns></returns>
        public List<D54ViewModel> getD54(DateTime reportDate, string groupProductCode, string bondNumber, string refNumber)
        {
            List<D54ViewModel> results = new List<D54ViewModel>();
            var dt = reportDate.ToString(C07reportDateFormate);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var D01 = db.Flow_Info.AsNoTracking()
                    .FirstOrDefault(x => x.Group_Product_Code == groupProductCode &&
                    x.Apply_Off_Date == null);
                if (D01 != null)
                {
                    var D54s = db.IFRS9_EL.AsNoTracking().Where(x => x.Report_Date == dt);
                    if (!D54s.Any())
                        return results;
                    var ver = D54s.Max(x => x.Version);
                    results = D54s.Where(x =>
                                 x.Version == ver &&
                                 x.FLOWID == D01.FLOWID &&
                                 x.PRJID == D01.PRJID)
                                 .Where(x => x.Bond_Number == bondNumber, !bondNumber.IsNullOrWhiteSpace())
                                 .Where(x => x.Reference_Nbr == refNumber, !refNumber.IsNullOrWhiteSpace())
                                 .AsEnumerable()
                                  .Select(x => DbToD54ViewModel(x)).ToList();
                }
            }
            return results;
        }
        #endregion

        #endregion Get Data

        #region Save Db

        /// <summary>
        /// Save D53
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD53(List<D53ViewModel> datas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (!datas.Any())
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                StringBuilder sb = new StringBuilder();
                sb.Append($@"
                delete SMF_Info ;  ");
                datas.ForEach(x =>
                {
                    sb.Append($@"
INSERT INTO [SMF_Info]
           ([SMF_Code]
           ,[SMF]
           ,[SMF_Desc]
           ,[Asset_type_Code]
           ,[Asset_Type])
     VALUES
           ({x.SMF_Code.stringToStrSql()}
           ,{x.SMF.stringToStrSql()}
           ,{x.SMF_Desc.stringToStrSql()}
           ,{x.Asset_type_Code.stringToStrSql()}
           ,{x.Asset_Type.stringToStrSql()}) ; ");
                });
                try
                {
                    db.Database.ExecuteSqlCommand(sb.ToString());
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        #region 減損階段確認
        /// <summary>
        /// 減損階段確認
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public MSGReturnModel SaveD54(string dt)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<string> Product_Codes = new List<string>()
                {
                    Product_Code.B_A.GetDescription(),
                    Product_Code.B_B.GetDescription(),
                    Product_Code.B_P.GetDescription()
                };
                var _Product_Codes = Product_Codes.stringListToInSql();
                DateTime reportDt = DateTime.MinValue;
                DateTime startDr = DateTime.Now;
                if (!DateTime.TryParse(dt, out reportDt))
                {
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    return result;
                }
                var A41s = db.Bond_Account_Info.AsNoTracking().Where(x => x.Report_Date == reportDt);
                if (!A41s.Any())
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A41.tableNameGetDescription());
                    return result;
                }
                var ver = A41s.Max(x => x.Version).Value;
                var rdt = reportDt.ToString(C07reportDateFormate);
                var _message = $@"基準日:{rdt} 版本:{ver}";
                if (!db.EL_Data_Out.AsNoTracking().Any(x => x.Report_Date == rdt && x.Version == ver))
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.C07.tableNameGetDescription(), _message);
                    return result;
                }
                if (!db.EL_Data_In_Update.AsNoTracking().Any(x => x.Report_Date == rdt && x.Version == ver))
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.C09.tableNameGetDescription(), _message);
                    return result;
                }
                if (db.IFRS9_EL.AsNoTracking().Any(x => x.Report_Date == rdt && x.Version == ver))
                {
                    result.DESCRIPTION = Message_Type.already_Save.GetDescription(Table_Type.D54.tableNameGetDescription(), _message);
                    return result;
                }
                string sql = $@"
WITH C07S AS
(
     SELECT C07.PRJID AS PRJID, --專案名稱
            C07.FLOWID AS FLOWID, --流程名稱
            C07.Report_Date AS Report_Date, --評估基準日/報導日
            C07.Version AS Version, --資料版本
            C07.Processing_Date AS Processing_Date, --資料處理日期
            C07.Product_Code AS Product_Code, --產品
            C07.Reference_Nbr AS Reference_Nbr, --案件編號/帳號
            A41.Bond_Number AS Bond_Number, --債券編號
            A41.Lots AS Lots, --Lots
            A41.Portfolio_Name AS Portfolio_Name, --Portfolio英文
            A41.Segment_Name AS Segment_Name, --債券(資產)名稱
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
	        END AS Impairment_Stage, --減損階段
            C07.PD AS PD, --第一年違約率
            C07.Lifetime_EL AS Lifetime_EL, --存續期間預期信用損失(原幣)
            C07.Y1_EL AS Y1_EL, --一年期預期信用損失(原幣)
	        CASE WHEN D66.Assessment_Stage = '2'
	             THEN C07.Lifetime_EL
	        	 WHEN D66.Assessment_Stage = '3' and C07.PD <> 0
	        	 THEN C07.Y1_EL / C07.PD
	        	 WHEN D66.Assessment_Stage = '3' and C07.PD = 0
	        	 THEN 0
	        ELSE C07.EL
	        END AS EL, --累計減損(原幣)(最終預期信用損失)
            A41.Ex_rate AS Ex_rate, --月底匯率(報表日匯率)
            A41.Ori_Ex_rate AS Ori_Ex_rate --成本匯率
     FROM (SELECT * FROM Bond_Account_Info
     where Report_Date =  @Report_Date 
     and Version = @Version
     ) AS A41 
     JOIN( 
     SELECT * 
     FROM EL_Data_Out 
     WHERE Report_Date = @Report_Date 
     AND Version = @Version 
     AND Product_Code IN ('Bond_A','Bond_B','Bond_P') ) AS C07
     ON C07.Reference_Nbr = A41.Reference_Nbr
     JOIN
     (select * from Flow_Info
     where Apply_Off_Date is null) AS D01
     on C07.PRJID = D01.PRJID
     and C07.FLOWID = D01.FLOWID
     left join 
     (select * from Bond_Qualitative_Assessment
     where Report_Date = @Report_Date 
     and Version = @Version 
     and Assessment_Result_Version = Result_Version_Confirm
     )AS D65
     on A41.Reference_Nbr = D65.Reference_Nbr
     left join 
     (select * from Bond_Qualitative_Assessment_Result
     where Report_Date = @Report_Date 
     and Version = @Version 
     and (Check_Item_Code like 'Pass_Count%' or Check_Item_Code like 'Z%')) AS D66
     ON D66.Reference_Nbr = D65.Reference_Nbr
	 AND D66.Assessment_Result_Version = D65.Assessment_Result_Version
)
Insert into IFRS9_EL
            (PRJID,
             FLOWID,
             Report_Date,
             Version,
             Processing_Date,
             Product_Code,
             Reference_Nbr,
             Bond_Number,
             Lots,
             Portfolio_Name,
             Segment_Name,
             Impairment_Stage,
             PD,
             Lifetime_EL,
             Y1_EL,
             EL,
             Ex_rate,
             Ori_Ex_rate,
             Lifetime_EL_Ex,
             Y1_EL_Ex,
             EL_Ex,
             Lifetime_EL_Ori_Ex,
             Y1_EL_Ori_Ex,
             EL_Ori_Ex,
             Exec_Date,
             Create_User,
             Create_Date,
             Create_Time)
Select
*,
Lifetime_EL * Ex_rate AS Lifetime_EL_Ex , --存續期間預期信用損失(報表日匯率台幣)
Y1_EL * Ex_rate AS Y1_EL_Ex , --一年期預期信用損失(報表日匯率台幣)
EL * Ex_rate AS EL_Ex, --累計減損(報表日匯率台幣)
Lifetime_EL * Ori_Ex_rate AS Lifetime_EL_Ori_Ex, --存續期間預期信用損失(成本匯率台幣)
Y1_EL * Ori_Ex_rate AS Y1_EL_Ori_Ex, --一年期預期信用損失(成本匯率台幣)
EL * Ori_Ex_rate AS EL_Ori_Ex, --累計減損(成本匯率台幣)
@Exec_Date  AS Exec_Date, --資料處理日期
@Create_User AS Create_User,
@Create_Date AS Create_Date,
@Create_Time AS Create_Time
FROM C07S ;
";

                string sql2 = $@"
WITH Temp AS
(
Select 
       A41.Reference_Nbr,
       CASE WHEN D66.Assessment_Stage = '2'
            THEN C07.Lifetime_EL
       	    WHEN D66.Assessment_Stage = '3' and C07.PD <> 0
       	    THEN C07.Y1_EL / C07.PD
       	    WHEN D66.Assessment_Stage = '3' and C07.PD = 0
       	    THEN 0
       ELSE C07.EL
       END AS EL, --累計減損(原幣)(最終預期信用損失)
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
       END AS Impairment_Stage, --減損階段
       C07.FLOWID,
       C07.PRJID
FROM (
SELECT * 
FROM Bond_Account_Info
WHERE Report_Date = @Report_Date
AND Version = @Version ) AS A41 
JOIN( 
SELECT * 
FROM EL_Data_Out 
WHERE Report_Date = @Report_Date 
AND Version = @Version 
AND Product_Code IN ('Bond_A','Bond_B','Bond_P') 
AND FLOWID in (select FLOWID from Flow_Info where Apply_Off_Date is null)
) AS C07
ON C07.Reference_Nbr = A41.Reference_Nbr
left join 
(select * from Bond_Qualitative_Assessment
where Report_Date = @Report_Date 
and Version = @Version 
and Assessment_Result_Version = Result_Version_Confirm
)AS D65
on A41.Reference_Nbr = D65.Reference_Nbr
left join 
(select * from Bond_Qualitative_Assessment_Result
where Report_Date = @Report_Date 
and Version = @Version  
and (Check_Item_Code like 'Pass_Count%' or Check_Item_Code like 'Z%')) AS D66
ON D66.Reference_Nbr = D65.Reference_Nbr
AND D66.Assessment_Result_Version = D65.Assessment_Result_Version
)
,D54Temp AS
(
SELECT C07.PRJID AS PRJID, --專案名稱
       C07.FLOWID AS FLOWID, --流程名稱
       C07.Report_Date AS Report_Date, --評估基準日/報導日
	   C07.Processing_Date AS Processing_Date, --資料處理日期
       C07.Product_Code AS Product_Code, --產品
       C07.Reference_Nbr AS Reference_Nbr, --案件編號/帳號
       C07.PD AS PD, --第一年違約率
       C07.Lifetime_EL AS Lifetime_EL, --存續期間預期信用損失(原幣)
       C07.Y1_EL AS Y1_EL, --一年期預期信用損失(原幣)
       t.EL AS EL, --累計減損(原幣)(最終預期信用損失)
       t.Impairment_Stage AS Impairment_Stage, --減損階段
       C07.Version AS Version, --資料版本
	   A41.Principal AS Principal, --金融資產餘額攤銷後之成本數(原幣) 
       A41.Ori_Amount , --帳列面額(原幣)
	   A41.Interest_Receivable , --應收利息(原幣)
	   A41.Amort_Amt_Ori_Tw  , --攤銷後成本(成本匯率台幣)
	   A41.Amort_Amt_Tw , --攤銷後成本(報表日匯率台幣)
	   A41.Ex_rate , --月底匯率(報表日匯率)
	   A41.Ori_Ex_rate , --成本匯率
	   ROUND( C07.Lifetime_EL * A41.Ex_rate ,0) AS Lifetime_EL_Ex,
	   --Lifetime_EL_Ex float, --存續期間預期信用損失(報表日匯率台幣)
	   ROUND( C07.Y1_EL * A41.Ex_rate ,0) AS Y1_EL_Ex,
	   --Y1_EL_Ex float, --一年期預期信用損失(報表日匯率台幣)
	   ROUND( t.EL * A41.Ex_rate ,0) AS EL_Ex,
	   --EL_Ex float, --累計減損(報表日匯率台幣)
	   CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 
                ROUND(t.EL * A41.Ex_rate , 0)
			ELSE 
			    CASE WHEN A41.Maturity_Date is not null
				     THEN
			             CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
				              THEN ROUND(t.EL * A41.Ex_rate , 0) - ROUND(ROUND(A41.Interest_Receivable * C07.PD *C09.Current_LGD,2) * A41.Ex_rate,0)
				         	  ELSE ROUND(t.EL * A41.Ex_rate , 0) - ROUND(ROUND(t.EL *  A41.Interest_Receivable/(A41.Principal + A41.Interest_Receivable),2) * A41.Ex_rate ,0)
				              END 
					 ELSE
					 ROUND(t.EL * A41.Ex_rate , 0) - ROUND(ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD,2) * A41.Ex_rate,0)
					 END
       END AS Principal_EL_Ex,
	   --Principal_EL_Ex float, --累計減損_本金(報表日匯率台幣) ps:債券的LGD 須從C09-減損計算輸入資料修正檔取得 EL_Data_In_Update
	   CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 0
			ELSE 
			    CASE WHEN A41.Maturity_Date is not null
				     THEN 
					      CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
					           THEN ROUND(ROUND(A41.Interest_Receivable * C07.PD *C09.Current_LGD,2) * A41.Ex_rate,0)
					      	   ELSE ROUND(ROUND(t.EL * A41.Interest_Receivable/(A41.Principal + A41.Interest_Receivable),2)*A41.Ex_rate ,0)
					      	   END
				     ELSE
					 ROUND(ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD,2) * A41.Ex_rate,0)
					 END
	   END AS Interest_Receivable_EL_Ex,
	   --Interest_Receivable_EL_Ex float, --累計減損_利息(報表日匯率台幣)
	   ROUND( C07.Lifetime_EL * A41.Ori_Ex_rate ,0) AS Lifetime_EL_Ori_Ex,
	   --Lifetime_EL_Ori_Ex float, --存續期間預期信用損失(成本匯率台幣)
	   ROUND( C07.Y1_EL * A41.Ori_Ex_rate ,0) AS Y1_EL_Ori_Ex,
	   --Y1_EL_Ori_Ex float, --一年期預期信用損失(成本匯率台幣)
	   ROUND( t.EL * A41.Ori_Ex_rate ,0) AS EL_Ori_Ex,
	   --EL_Ori_Ex float, --累計減損(成本匯率台幣)
	   CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 
                 ROUND(t.EL * A41.Ori_Ex_rate,0)
			ELSE 
			     CASE WHEN A41.Maturity_Date is not null
				      THEN
					       CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
						        THEN ROUND(t.EL * A41.Ori_Ex_rate,0) - ROUND(ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD,2) * A41.Ori_Ex_rate , 0)
								ELSE ROUND(t.EL * A41.Ori_Ex_rate,0) - ROUND(ROUND(t.EL * A41.Interest_Receivable /(A41.Principal + A41.Interest_Receivable),2) * A41.Ori_Ex_rate,0)
								END
					  ELSE
					  ROUND(t.EL * A41.Ori_Ex_rate,0) - ROUND(ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD,2) * A41.Ori_Ex_rate,0)
					  END
	   END AS Principal_EL_Ori_Ex,
	   --Principal_EL_Ori_Ex float, --累計減損_本金(成本匯率台幣)
	   CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 0
			ELSE 
			     CASE WHEN A41.Maturity_Date is not null
				      THEN 
					       CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
						        THEN ROUND(ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD,2) * A41.Ori_Ex_rate ,0)
								ELSE ROUND(ROUND(t.EL * A41.Interest_Receivable / (A41.Principal + A41.Interest_Receivable),2) * A41.Ori_Ex_rate , 0)
								END
					  ELSE
					  ROUND(A41.Interest_Receivable * C07.PD * C09.Current_LGD * A41.Ori_Ex_rate,0)
					  END
		END AS Interest_Receivable_EL_Ori_Ex,
	   --Interest_Receivable_EL_Ori_Ex float, --累計減損_利息(成本匯率台幣)
	   ROUND(t.EL * A41.Ex_rate , 0) - ROUND(t.EL * A41.Ori_Ex_rate , 0) AS EL_EX_Diff , --累計減損匯兌損益(台幣)
	   A41.ISSUER , --ISSUER
	   A41.Bond_Number , --債券編號 
	   A41.Lots , --Lots
       A41.Security_Name, --Security_Name
	   A41.Segment_Name , --債券(資產)名稱
	   A41.Origination_Date  , --債券購入(認列)日期
	   A41.Maturity_Date  , --債券到期日
	   A41.ISSUER_AREA , --Issuer所屬區域
	   A41.Industry_Sector  , --對手產業別
	   A41.IAS39_CATEGORY , --公報分類
	   A41.Bond_Aera  , --國內\國外
	   A41.Asset_Type , --金融商品分類
	   A41.Currency_Code, --幣別
	   A41.ASSET_SEG , --資產區隔
	   A41.IH_OS , --自操\委外
	   A41.Portfolio , --Portfolio
       A41.PRODUCT , --SMF
	   A41.Market_Value_Ori  , --市價(原幣)
	   A41.Market_Value_TW , --市價(報表日匯率台幣)
	   A41.Current_Int_Rate  , --合?利率/產品利率
	   A41.EIR , --有效利率
	   A41.Lien_position , --擔保順位
	   A41.Principal_Payment_Method_Code , --現金流類型
	   C09.Current_LGD , --違約損失率
	   C09.Exposure , --曝險額
	   B01.Original_External_Rating, --購買時評等(主標尺)
	   B01.Current_External_Rating , --最近一次評等(主標尺)
	   A41.Portfolio_Name, --Portfolio英文
	   D62.Value_Change_Ratio ,  --未實現累計損失率
       D62.Accumulation_Loss_last_Month  , --未實現累計損失月數_上個月狀況
	   D62.Accumulation_Loss_This_Month  , --未實現累計損失月數_本月狀況
	   D62.Watch_IND , --是否為觀察名單
	   D62.Map_Rule_Id_D70 ,  --對應D70-觀察名單參數檔規則編號
	   D62.Warning_IND , --是否為預警名單
	   D62.Map_Rule_Id_D71  , --對應D71-預警名單參數檔規則編號
	   A41.Ori_Amount * A41.Ori_Ex_rate AS Ori_Amount_Ori_Ex,
	   --Ori_Amount_Ori_Ex float , --帳列面額(成本匯率台幣)
	   	CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 
                 ROUND(t.EL , 2)
			ELSE 
			     CASE WHEN A41.Maturity_Date is not null
				      THEN 
					       CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
						        THEN ROUND(t.EL - ROUND(A41.Interest_Receivable * c07.PD * C09.Current_LGD,2),2)
								ELSE ROUND(t.EL - ROUND(t.EL * A41.Interest_Receivable / (A41.Principal + A41.Interest_Receivable),2),2)
								END
					  ELSE
					  t.EL - A41.Interest_Receivable * C07.PD * C09.Current_LGD
					  END
		END AS Principal_EL_Ori, 
	   --Principal_EL_Ori float , --累計減損-本金(原幣)
	   	   	CASE WHEN A41.PRODUCT like 'D21%'
	        THEN 0
			ELSE 
			     CASE WHEN A41.Maturity_Date is not null
				      THEN 
					       CASE WHEN datediff(MONTH,A41.Report_Date,A41.Maturity_Date) /12 > 1
						        THEN ROUND(A41.Interest_Receivable * c07.PD * C09.Current_LGD,2)
								ELSE ROUND(t.EL * A41.Interest_Receivable / (A41.Principal + A41.Interest_Receivable),2)
								END
					  ELSE
					  A41.Interest_Receivable * C07.PD * C09.Current_LGD
					  END
		END AS Interest_Receivable_EL_Ori, 
	   --Interest_Receivable_EL_Ori float , --累計減損-利息(原幣)
	   A41.Interest_Receivable_tw , --應收利息(報表日匯率台幣)
        ROUND(t.EL * A41.Ex_rate , 0) - ROUND(t.EL * A41.Ori_Ex_rate , 0)  AS Principal_Diff_Tw,
	   --預先註記  ((t.EL - (A41.Interest_Receivable * C07.PD * C09.Current_LGD)) * (A41.Ex_rate - A41.Ori_Ex_rate)) AS Principal_Diff_Tw,
	   --Principal_Diff_Tw float , --累計減損匯兌損益-本金(台幣)
	   A41.ISIN_Changed_Ind, --是否為換券
	   A41.Bond_Number_Old , --債券編號_換券前
	   A41.Lots_Old  , --Lots_舊
	   A41.Portfolio_Name_Old , --Portfolio英文_舊
	   A41.Origination_Date_Old   --舊券原始購入日
FROM (
SELECT * 
FROM Bond_Account_Info
WHERE Report_Date = @Report_Date 
AND Version = @Version  ) AS A41 
JOIN( 
SELECT * 
FROM EL_Data_Out 
WHERE Report_Date = @Report_Date 
AND Version = @Version 
AND Product_Code IN ({_Product_Codes}) ) AS C07
ON C07.Reference_Nbr = A41.Reference_Nbr
JOIN Temp t
ON t.Reference_Nbr = C07.Reference_Nbr
AND t.FLOWID = C07.FLOWID
AND t.PRJID = C07.PRJID
JOIN (
SELECT * 
FROM EL_Data_In_Update
WHERE Report_Date = @Report_Date 
AND Version = @Version ) AS C09
ON C07.Reference_Nbr = C09.Reference_Nbr
AND C07.PRJID = C09.PrjID
AND C07.FLOWID = C09.FlowID
AND C07.Product_Code = C09.Product_Code
JOIN (
SELECT * 
FROM IFRS9_Main
WHERE Report_Date = @Report_Date
AND Version = @Version  ) AS B01
ON A41.Reference_Nbr = B01.Reference_Nbr
left JOIN(
SELECT * 
FROM Bond_Basic_Assessment
WHERE Report_Date = @Report_Date
AND Version = @Version ) AS D62
ON A41.Reference_Nbr = D62.Reference_Nbr
)
INSERT INTO [IFRS9_Bond_Report]
           ([PRJID]
           ,[FLOWID]
           ,[Report_Date]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Reference_Nbr]
           ,[PD]
           ,[Lifetime_EL]
           ,[Y1_EL]
           ,[EL]
           ,[Impairment_Stage]
           ,[Version]
           ,[Principal]
           ,[Ori_Amount]
           ,[Interest_Receivable]
           ,[Amort_Amt_Ori_Tw]
           ,[Amort_Amt_Tw]
           ,[Ex_rate]
           ,[Ori_Ex_rate]
           ,[Lifetime_EL_Ex]
           ,[Y1_EL_Ex]
           ,[EL_Ex]
           ,[Principal_EL_Ex]
           ,[Interest_Receivable_EL_Ex]
           ,[Lifetime_EL_Ori_Ex]
           ,[Y1_EL_Ori_Ex]
           ,[EL_Ori_Ex]
           ,[Principal_EL_Ori_Ex]
           ,[Interest_Receivable_EL_Ori_Ex]
           ,[EL_EX_Diff]
           ,[ISSUER]
           ,[Bond_Number]
           ,[Lots]
           ,[Security_Name]
           ,[Segment_Name]
           ,[Origination_Date]
           ,[Maturity_Date]
           ,[ISSUER_AREA]
           ,[Industry_Sector]
           ,[IAS39_CATEGORY]
           ,[Bond_Aera]
           ,[Asset_Type]
           ,[Currency_Code]
           ,[ASSET_SEG]
           ,[IH_OS]
           ,[Portfolio]
           ,[PRODUCT]
           ,[Market_Value_Ori]
           ,[Market_Value_TW]
           ,[Current_Int_Rate]
           ,[EIR]
           ,[Lien_position]
           ,[Principal_Payment_Method_Code]
           ,[Current_LGD]
           ,[Exposure]
           ,[Original_External_Rating]
           ,[Current_External_Rating]
           ,[Portfolio_Name]
           ,[Value_Change_Ratio]
           ,[Accumulation_Loss_last_Month]
           ,[Accumulation_Loss_This_Month]
           ,[Watch_IND]
           ,[Map_Rule_Id_D70]
           ,[Warning_IND]
           ,[Map_Rule_Id_D71]
           ,[Ori_Amount_Ori_Ex]
           ,[Principal_EL_Ori]
           ,[Interest_Receivable_EL_Ori]
           ,[Interest_Receivable_Tw]
           ,[Principal_Diff_Tw]
           ,[ISIN_Changed_Ind]
           ,[Bond_Number_Old]
           ,[Lots_Old]
           ,[Portfolio_Name_Old]
           ,[Origination_Date_Old]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time])
select      [PRJID]
           ,[FLOWID]
           ,[Report_Date]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Reference_Nbr]
           ,[PD]
           ,[Lifetime_EL]
           ,[Y1_EL]
           ,[EL]
           ,[Impairment_Stage]
           ,[Version]
           ,[Principal]
           ,[Ori_Amount]
           ,[Interest_Receivable]
           ,[Amort_Amt_Ori_Tw]
           ,[Amort_Amt_Tw]
           ,[Ex_rate]
           ,[Ori_Ex_rate]
           ,[Lifetime_EL_Ex]
           ,[Y1_EL_Ex]
           ,[EL_Ex]
           ,[Principal_EL_Ex]
           ,[Interest_Receivable_EL_Ex]
           ,[Lifetime_EL_Ori_Ex]
           ,[Y1_EL_Ori_Ex]
           ,[EL_Ori_Ex]
           ,[Principal_EL_Ori_Ex]
           ,[Interest_Receivable_EL_Ori_Ex]
           ,[EL_EX_Diff]
           ,[ISSUER]
           ,[Bond_Number]
           ,[Lots]
           ,[Security_Name]
           ,[Segment_Name]
           ,[Origination_Date]
           ,[Maturity_Date]
           ,[ISSUER_AREA]
           ,[Industry_Sector]
           ,[IAS39_CATEGORY]
           ,[Bond_Aera]
           ,[Asset_Type]
           ,[Currency_Code]
           ,[ASSET_SEG]
           ,[IH_OS]
           ,[Portfolio]
           ,[PRODUCT]
           ,[Market_Value_Ori]
           ,[Market_Value_TW]
           ,[Current_Int_Rate]
           ,[EIR]
           ,[Lien_position]
           ,[Principal_Payment_Method_Code]
           ,[Current_LGD]
           ,[Exposure]
           ,[Original_External_Rating]
           ,[Current_External_Rating]
           ,[Portfolio_Name]
           ,[Value_Change_Ratio]
           ,[Accumulation_Loss_last_Month]
           ,[Accumulation_Loss_This_Month]
           ,[Watch_IND]
           ,[Map_Rule_Id_D70]
           ,[Warning_IND]
           ,[Map_Rule_Id_D71]
           ,[Ori_Amount_Ori_Ex]
           ,[Principal_EL_Ori]
           ,[Interest_Receivable_EL_Ori]
           ,[Interest_Receivable_Tw]
           ,(Principal_EL_Ex - Principal_EL_Ori_Ex)
           --,(ROUND(Principal_EL_Ori * Ex_rate , 0) - ROUND(Principal_EL_Ori * Ori_Ex_rate , 0))
           --,[Principal_Diff_Tw]
           ,[ISIN_Changed_Ind]
           ,[Bond_Number_Old]
           ,[Lots_Old]
           ,[Portfolio_Name_Old]
           ,[Origination_Date_Old]
           ,@Create_User
           ,@Create_Date
           ,@Create_Time
from D54Temp ;
";

                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    var sqs1 = new List<SqlParameter>();
                    sqs1.Add(new SqlParameter("Report_Date", reportDt.ToString("yyyy-MM-dd")));
                    sqs1.Add(new SqlParameter("Version", ver));
                    sqs1.Add(new SqlParameter("Exec_Date", startDr));
                    sqs1.Add(new SqlParameter("Create_User", _UserInfo._user));
                    sqs1.Add(new SqlParameter("Create_Date", _UserInfo._date));
                    sqs1.Add(new SqlParameter("Create_Time", _UserInfo._time));
                    try
                    {
                        int _D54 = db.Database.ExecuteSqlCommand(sql, sqs1.ToArray()); //saveD54
                        System.Reflection.FieldInfo info = typeof(SqlParameter).GetField("_parent", (System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic));
                        foreach (SqlParameter sp in sqs1)
                        {
                            info.SetValue(sp, null);
                        }
                        int _D52 = db.Database.ExecuteSqlCommand(sql2, sqs1.ToArray()); //saveD52
                        if (_D54 == _D52)
                        {
                            if (_D52 != 0)
                            {
                                dbContextTransaction.Commit();
                                result.RETURN_FLAG = true;
                                result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                            }
                            else
                            {
                                result.DESCRIPTION = "無新增任何資料";
                            } 
                        }
                        else
                        {
                            result.DESCRIPTION = "資料有缺漏D54(最終減損金額)新增筆數不等於D52(報表需求資訊)";
                        }
                    }
                    catch (Exception ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }
        #endregion

        #endregion Save Db

        #region Excel 部分
        /// <summary>
        /// Excel 資料轉成 D53ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<D53ViewModel>> GetD53Excel(string pathType, Stream stream)
        {
            List<D53ViewModel> dataModel = new List<D53ViewModel>();
            string message = string.Empty;
            DataSet resultData = new DataSet();
            try
            {
                IExcelDataReader reader = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        break;

                    case "xlsx":
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        break;
                }
                reader.IsFirstRowAsColumnNames = true;
                resultData = reader.AsDataSet();
                reader.Close();
                if (resultData.Tables[0].Rows.Count > 0) //判斷有無資料
                {
                    dataModel = common.getViewModel(resultData.Tables[0], Table_Type.D53)
                                      .Cast<D53ViewModel>().ToList();
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<D53ViewModel>>(message, dataModel);
        }
        #endregion Excel 部分

        #region Private Function

        #region Db 組成 D54ViewModel
        private D54ViewModel DbToD54ViewModel(IFRS9_EL data)
        {
            return new D54ViewModel()
            {
                Bond_Number = data.Bond_Number,
                EL = TypeTransfer.doubleNToString(data.EL),
                EL_Ex = TypeTransfer.doubleNToString(data.EL_Ex),
                EL_Ori_Ex = TypeTransfer.doubleNToString(data.EL_Ori_Ex),
                Exec_Date = TypeTransfer.dateTimeNToString(data.Exec_Date),
                Ex_rate = TypeTransfer.doubleNToString(data.Ex_rate),
                FLOWID = data.FLOWID,
                Impairment_Stage = data.Impairment_Stage,
                Lifetime_EL = TypeTransfer.doubleNToString(data.Lifetime_EL),
                Lifetime_EL_Ex = TypeTransfer.doubleNToString(data.Lifetime_EL_Ex),
                Lifetime_EL_Ori_Ex = TypeTransfer.doubleNToString(data.Lifetime_EL_Ori_Ex),
                Lots = data.Lots,
                Ori_Ex_rate = TypeTransfer.doubleNToString(data.Ori_Ex_rate),
                PD = TypeTransfer.doubleNToString(data.PD),
                Portfolio_Name = data.Portfolio_Name,
                PRJID = data.PRJID,
                Processing_Date = data.Processing_Date,
                Product_Code = data.Product_Code,
                Reference_Nbr = data.Reference_Nbr,
                Report_Date = data.Report_Date,
                Segment_Name = data.Segment_Name,
                Version = data.Version.ToString(),
                Y1_EL = TypeTransfer.doubleNToString(data.Y1_EL),
                Y1_EL_Ex = TypeTransfer.doubleNToString(data.Y1_EL_Ex),
                Y1_EL_Ori_Ex = TypeTransfer.doubleNToString(data.Y1_EL_Ori_Ex)
            };
        }
        #endregion Db 組成 D54ViewModel

        private string formateProduct_Code(string value)
        {
            if (value == "01" || value == "1")
                return "Bond_A";
            if (value == "02" || value == "2")
                return "Bond_B";
            if (value == "04" || value == "4")
                return "Bond_P";
            return value;
        }


        #endregion Private Function
    }
}