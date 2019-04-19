using Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Transactions;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A4Repository : IA4Repository
    {
        #region 其他

        public A4Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
        }


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
        #endregion 其他

        #region Get Data

        #region  get A41 data
        /// <summary>
        /// get A41 data
        /// </summary>
        /// <param name="ReportDate">報導日</param>
        /// <param name="OriginationDate">購入日</param>
        /// <param name="Version">版本</param>
        /// <param name="BondNumber">債券編號</param>
        /// <returns></returns>
        public List<A41ViewModel> GetA41(DateTime ReportDate, DateTime OriginationDate, int Version,string BondNumber)
        {
            string resultMessage = string.Empty;
            List<A41ViewModel> data = new List<A41ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                data = db.Bond_Account_Info.AsNoTracking()
                    .Where(x => x.Report_Date == ReportDate &&
                                x.Version != null &&
                                x.Version == Version)
                    .Where(x => x.Origination_Date != null &&
                                x.Origination_Date == OriginationDate,
                                  OriginationDate != DateTime.MinValue)
                    .Where(x => x.Bond_Number == BondNumber,
                                  !BondNumber.IsNullOrWhiteSpace())
                    .AsEnumerable()
                    .OrderBy(x => Convert.ToInt32(x.Reference_Nbr))
                    .Select(x => DbToA41Model(x)).ToList();
            }
            return data;
        }
        #endregion

        #region Get A45 Data
        public Tuple<bool, List<A45ViewModel>> getA45(string bloombergTicker, string processingDate)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Category_Info.Any())
                {
                    var query = from q in db.Bond_Category_Info.AsNoTracking()
                                select q;

                    if (!bloombergTicker.IsNullOrWhiteSpace())
                    {
                        query = query.Where(x => x.Bloomberg_Ticker == bloombergTicker);
                    }

                    if (!processingDate.IsNullOrWhiteSpace())
                    {
                        DateTime dateProcessingDate = DateTime.Parse(processingDate);
                        query = query.Where(x => x.Processing_Date == dateProcessingDate);
                    }

                    query.OrderBy(x => x.Processing_Date).ThenBy(x => x.Bloomberg_Ticker);

                    return new Tuple<bool, List<A45ViewModel>>(query.Any(), query.AsEnumerable()
                                                               .Select(x => { return DbToA45Model(x); }).ToList());
                }
            }
            return new Tuple<bool, List<A45ViewModel>>(false, new List<A45ViewModel>());
        }

        #endregion

        #region Get A46 Data
        /// <summary>
        /// Get A46 Data
        /// </summary>
        /// <param name="searchAll">查詢全部資料</param>
        /// <returns></returns>
        public List<A46ViewModel> GetA46(bool searchAll)
        {
            List<A46ViewModel> A46Data = new List<A46ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A46s = db.Fixed_Income_CEIC_Info.AsNoTracking();
                if (!A46s.Any())                 
                    return A46Data;
                var dbe = A46s.Where(x => x.Effective == "Y", !searchAll).AsEnumerable()
                    .OrderByDescending(x => x.Processing_Date).ThenBy(x => x.Country)
                    .ThenByDescending(x => x.Effective);
                A46Data = common.getViewModel(dbe, Table_Type.A46).OfType<A46ViewModel>().ToList();
                return A46Data;
            }
        }
        #endregion

        #region Get A47 Data
        /// <summary>
        /// Get A47 Data
        /// </summary>
        /// <param name="searchAll">查詢全部資料</param>
        /// <returns></returns>
        public List<A47ViewModel> GetA47(bool searchAll)
        {
            List<A47ViewModel> A47Data = new List<A47ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A47s = db.Fixed_Income_Moody_Info.AsNoTracking();
                if (!A47s.Any())
                    return A47Data;
                var dbe = A47s.Where(x => x.Effective == "Y", !searchAll).AsEnumerable()
                         .OrderByDescending(x => x.Processing_Date).ThenBy(x => x.Country)
                         .ThenByDescending(x => x.Effective);
                A47Data = common.getViewModel(dbe, Table_Type.A47).OfType<A47ViewModel>().ToList();
                return A47Data;
            }
        }
        #endregion

        #region Get A48 Data
        /// <summary>
        /// Get A48 Data
        /// </summary>
        /// <param name="searchAll">查詢全部資料</param>
        /// <returns></returns>
        public List<A48ViewModel> GetA48(bool searchAll)
        {
            List<A48ViewModel> A48Data = new List<A48ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A48s = db.Fixed_Income_AbuDhabi_Info.AsNoTracking();
                if (!A48s.Any())
                    return A48Data;
                var dbe = A48s.Where(x => x.Effective == "Y", !searchAll).AsEnumerable()
                              .OrderByDescending(x => x.Data_Year).ThenByDescending(x => x.Effective);
                A48Data = common.getViewModel(dbe, Table_Type.A48).OfType<A48ViewModel>().ToList();
                return A48Data;
            }
        }
        #endregion

        #region  get A95 data
        /// <summary>
        /// get A95 data
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="verison"></param>
        /// <param name="bondType">債券種類</param>
        /// <param name="ASK">產業別</param>
        /// <returns></returns>
        public Tuple<bool, List<A95ViewModel>> GetA95(DateTime reportDate, int verison,string bondType,string ASK)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A95s = db.Bond_Ticker_Info.AsNoTracking()
                    .Where(x => x.Report_Date == reportDate && x.Version == verison)
                    .Where(x => x.Bond_Type == bondType, bondType != "All")
                    .Where(x => x.Assessment_Sub_Kind == ASK,ASK != "All").AsEnumerable();
                if (A95s.Any())
                {
                    return new Tuple<bool, List<A95ViewModel>>(true,
                        (A95s.Select(x => DbToA95Model(x))).ToList());
                }                
            }
            return new Tuple<bool, List<A95ViewModel>>(false, new List<A95ViewModel>());
        }
        #endregion

        #region GetLogData

        /// <summary>
        /// get IFRS9_data
        /// </summary>
        /// <param name="tableTypes">table 代碼 (A41,B01,C01...)</param>
        /// <param name="debt">B : 債券 ,M : 房貸</param>
        /// <returns></returns>
        public List<string> GetLogData(List<string> tableTypes, string debt)
        {
            List<string> result = new List<string>();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.IFRS9_Log.Any() && tableTypes.Any())
                    {
                        foreach (string tableType in tableTypes)
                        {
                            var items = db.IFRS9_Log.AsNoTracking()
                                 .Where(x => tableType.Equals(x.Table_type) &&
                                 debt.Equals(x.Debt_Type)).ToList();
                            if (items.Any())
                            {
                                var lastDate = items.Max(y => y.Create_date);
                                result.AddRange(items.Where(x => lastDate.Equals(x.Create_date))
                                    .OrderByDescending(x => x.Create_time)
                                    .Select(x =>
                                    {
                                    return string.Format("{0} 基準日期:{1} {2} 轉檔日期:{3} 轉檔時間:{4} {5}",
                                        x.Table_type, 
                                        x.Report_Date != null ? x.Report_Date.Value.ToString("yyyy/MM/dd") : string.Empty,
                                        x.Version != null ? (x.Version.Value != 0 ?
                                        string.Format("版本:{0}", x.Version.Value.ToString()) : string.Empty) : string.Empty,
                                        x.Create_date,
                                        x.Create_time,
                                        "Y".Equals(x.TYPE) ? "成功" : "失敗"
                                        );
                                    }));
                            }
                        }
                    }
                }
            }
            catch
            {
            }
            return result;
        }

        #endregion GetLogData

        #endregion Get Data

        #region Get A42 Data

        /// <summary>
        /// get A42 data
        /// </summary>
        /// <param name="reportDate">評估基準日/報導日</param>
        /// <returns></returns>
        public Tuple<bool, List<A42ViewModel>> getA42(string reportDate)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Treasury_Securities_Info.Any())
                {
                    DateTime dateReportDate = DateTime.Parse(reportDate);

                    var data = from q in db.Treasury_Securities_Info.AsNoTracking()
                                            .Where(x => x.Report_Date == dateReportDate)
                               select q;

                    return new Tuple<bool, List<A42ViewModel>>(data.Any(), data.AsEnumerable()
                                                                           .Select(x => { return DbToA42Model(x); }).ToList());
                }
            }
            return new Tuple<bool, List<A42ViewModel>>(false, new List<A42ViewModel>());
        }

        #endregion Get A42 Data

        #region checkA42Duplicate

        /// <summary>
        /// checkA42Duplicate
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        public MSGReturnModel checkA42Duplicate(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                IEnumerable<Treasury_Securities_Info> query = null;
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    query = db.Treasury_Securities_Info.AsNoTracking().AsEnumerable()
                              .Where(x => x.Report_Date.ToString("yyyy/MM/dd") == reportDate);

                    if (query != null && query.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "評估基準日 / 報導日：" + reportDate + " 的資料已存在，不可重複上傳、轉檔";
                        return result;
                    }
                }

                result.RETURN_FLAG = true;
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }

        #endregion checkA42Duplicate

        #region getA42Excel
        public List<A42ViewModel> getA42Excel(string pathType, string path, string processingDate, string reportDate)
        {
            List<A42ViewModel> dataModel = new List<A42ViewModel>();
            try
            {
                IWorkbook wb = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new HSSFWorkbook(stream);
                        }
                        break;

                    case "xlsx":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new XSSFWorkbook(stream);
                        }
                        break;
                }

                ISheet sheet = wb.GetSheetAt(0);
                DataTable dt = sheet.ISheetToDataTable(true);
                if (dt.Rows.Count > 0) //判斷有無資料
                {
                    dataModel = dt.AsEnumerable()
                        .Select((x, y) =>
                        {
                            return getA42ViewModel(x, processingDate, reportDate);
                        }
                    ).ToList();
                }
            }
            catch (Exception ex)
            { }

            return dataModel;
        }

        #endregion

        #region Db 組成 A42ViewModel

        /// <summary>
        /// Db 組成 A42ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private A42ViewModel DbToA42Model(Treasury_Securities_Info data)
        {
            return new A42ViewModel()
            {
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Segment_Name = data.Segment_Name,
                Portfolio_Name = data.Portfolio_Name,
                Bond_Value = TypeTransfer.doubleNToString(data.Bond_Value),
                Ori_Amount = TypeTransfer.doubleNToString(data.Ori_Amount),
                Principal = TypeTransfer.doubleNToString(data.Principal),
                Amort_value = TypeTransfer.doubleNToString(data.Amort_value),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Security_Name = data.Security_Name
            };
        }

        #endregion Db 組成 A42ViewModel

        #region save A42

        /// <summary>
        /// A42 save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveA42(List<A42ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                if (!dataModel.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    return result;
                }

                StringBuilder sb = new StringBuilder();

                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    DateTime dtReportDate = DateTime.Parse(dataModel[0].Report_Date);
                    db.Treasury_Securities_Info.RemoveRange(db.Treasury_Securities_Info
                                                              .Where(x=>x.Report_Date == dtReportDate));
                    foreach (var item in dataModel)
                    {
                        db.Treasury_Securities_Info.Add(
                        new Treasury_Securities_Info()
                        {
                            Bond_Number = item.Bond_Number,
                            Lots = item.Lots,
                            Segment_Name = item.Segment_Name,
                            Portfolio_Name = item.Portfolio_Name,
                            Bond_Value = TypeTransfer.stringToDouble(item.Bond_Value),
                            Ori_Amount = TypeTransfer.stringToDouble(item.Ori_Amount),
                            Principal = TypeTransfer.stringToDouble(item.Principal),
                            Amort_value = TypeTransfer.stringToDouble(item.Amort_value),
                            Processing_Date = TypeTransfer.stringToDateTime(item.Processing_Date),
                            Report_Date = TypeTransfer.stringToDateTime(item.Report_Date),
                            Security_Name = item.Security_Name
                        });

                        sb.Append($@"UPDATE Bond_Account_Info 
                                     SET Principal = {item.Principal}
                                     WHERE Report_Date = '{item.Report_Date}'
                                           AND Security_Name = '{item.Security_Name}'
                                           AND Lots = '{item.Lots}' 
                                           AND Portfolio_Name = '{item.Portfolio_Name}';");
                    }

                    db.SaveChanges(); //Save

                    db.Database.ExecuteSqlCommand(sb.ToString());

                    result.RETURN_FLAG = true;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }

        #endregion Save A42

        #region datarow 組成 A42ViewModel

        /// <summary>
        /// datarow 組成 A42ViewModel
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>A42ViewModel</returns>
        private A42ViewModel getA42ViewModel(DataRow item, string processingDate, string reportDate)
        {
            return new A42ViewModel()
            {
                //Bond_Number = TypeTransfer.objToString(item[0]),
                Bond_Number = TypeTransfer.objToString(item[8]),
                Lots = TypeTransfer.objToString(item[1]),
                Segment_Name = TypeTransfer.objToString(item[2]),
                Portfolio_Name = TypeTransfer.objToString(item[3]),
                Bond_Value = TypeTransfer.objToString(item[4]),
                Ori_Amount = TypeTransfer.objToString(item[5]),
                Principal = TypeTransfer.objToString(item[6]),
                Amort_value = TypeTransfer.objToString(item[7]),
                Processing_Date = processingDate,
                Report_Date = reportDate,
                Security_Name = TypeTransfer.objToString(item[8])
            };
        }

        #endregion datarow 組成 A42ViewModel

        #region DownLoadA42Excel
        public MSGReturnModel DownLoadA42Excel(string path, List<A42ViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail.GetDescription();

            if (dbDatas.Any())
            {
                DataTable datas = getA42ModelFromDb(dbDatas).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.A42);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }

        #endregion

        #region getA42ModelFromDb
        private Tuple<DataTable> getA42ModelFromDb(List<A42ViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                dt.Columns.Add("債券編號", typeof(object));
                dt.Columns.Add("Lots", typeof(object));
                dt.Columns.Add("債券(資產)名稱", typeof(object));
                dt.Columns.Add("Portfolio_Name", typeof(object));
                dt.Columns.Add("面額", typeof(object));
                dt.Columns.Add("成交金額", typeof(object));
                dt.Columns.Add("攤銷後成本", typeof(object));
                dt.Columns.Add("本期攤銷數", typeof(object));
                dt.Columns.Add("Security_Name", typeof(object));

                foreach (A42ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    nrow["債券編號"] = item.Bond_Number;
                    nrow["Lots"] = item.Lots;
                    nrow["債券(資產)名稱"] = item.Segment_Name;
                    nrow["Portfolio_Name"] = item.Portfolio_Name;
                    nrow["面額"] = item.Bond_Value;
                    nrow["成交金額"] = item.Ori_Amount;
                    nrow["攤銷後成本"] = item.Principal;
                    nrow["本期攤銷數"] = item.Amort_value;
                    nrow["Security_Name"] = item.Security_Name;

                    dt.Rows.Add(nrow);
                }
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }
        #endregion

        #region Save DB 部分

        #region Save A41

        /// <summary>
        ///  A41 save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public MSGReturnModel saveA41(List<A41ViewModel> dataModel, string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime start = DateTime.Now;
            _UserInfo._date = start.Date;
            _UserInfo._time = start.TimeOfDay;
            string type = Table_Type.A41.ToString();
            List<Rating_Info_SampleInfo> A53SampleInfos = new List<Rating_Info_SampleInfo>();
            List<Assessment_Sub_Kind_Ticker> A95_1Datas = new List<Assessment_Sub_Kind_Ticker>();
            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return result;
            }
            DateTime dt = TypeTransfer.stringToDateTime(reportDate);
            int verInt = 0;
            if (!Int32.TryParse(version, out verInt))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (!common.checkTransferCheck(type, type, dt, verInt) ||
                     db.Bond_Account_Info
                    .Any(x => x.Report_Date != null &&
                    x.Report_Date == dt &&
                    x.Version == verInt))
                //資料裡面已經有相同的 Report_Date , version
                {
                    common.saveTransferCheck(
                        type,
                        false,
                        dt,
                        verInt,
                        start,
                        DateTime.Now,
                        Message_Type.already_Save.GetDescription()
                    );
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.already_Save.GetDescription();
                    return result;
                }
                A53SampleInfos = db.Rating_Info_SampleInfo.AsNoTracking().Where(x=>x.Report_Date == dt).ToList();
                A95_1Datas = db.Assessment_Sub_Kind_Ticker.AsNoTracking().ToList();
                StringBuilder sb = new StringBuilder();
                foreach (var item in dataModel)
                {
                    #region insert A41
                    //A41 Bond_Account_Info
                    sb.Append(
 $@" INSERT INTO [Bond_Account_Info]
           ([Reference_Nbr]
           ,[Bond_Number]
           ,[Lots]
           ,[Segment_Name]
           ,[CURR_SP_Issuer]
           ,[CURR_Moodys_Issuer]
           ,[CURR_Fitch_Issuer]
           ,[CURR_TW_Issuer]
           ,[CURR_SP_Issue]
           ,[CURR_Moodys_Issue]
           ,[CURR_Fitch_Issue]
           ,[CURR_TW_Issue]
           ,[Ori_Amount]
           ,[Current_Int_Rate]
           ,[Origination_Date]
           ,[Maturity_Date]
           ,[Principal_Payment_Method_Code]
           ,[Payment_Frequency]
           ,[Baloon_Freq]
           ,[ISSUER_AREA]
           ,[Industry_Sector]
           ,[PRODUCT]
           ,[IAS39_CATEGORY]
           ,[Principal]
           ,[Amort_Amt_Tw]
           ,[Interest_Receivable]
           ,[Interest_Receivable_tw]
           ,[Interest_Rate_Type]
           ,[IMPAIR_YN]
           ,[EIR]
           ,[Currency_Code]
           ,[Report_Date]
           ,[ISSUER]
           ,[Country_Risk]
           ,[Ex_rate]
           ,[Lien_position]
           ,[Portfolio]
           ,[ASSET_SEG]
           ,[Ori_Ex_rate]
           ,[Bond_Type]
           ,[Assessment_Sub_Kind]
           ,[Processing_Date]
           ,[Version]
           ,[Bond_Aera]
           ,[Asset_Type]
           ,[IH_OS]
           ,[Amount_TW_Ori_Ex_rate]
           ,[Amort_Amt_Ori_Tw]
           ,[Market_Value_Ori]
           ,[Market_Value_TW]
           ,[Value_date]
           ,[Portfolio_Name]
           ,[Security_Name]
           ,[ISIN_Changed_Ind]
           ,[Bond_Number_Old]
           ,[Lots_Old]
           ,[Portfolio_Name_Old]
           ,[Origination_Date_Old]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
)
     VALUES
           (
           {item.Reference_Nbr.stringToStrSql()},
           {item.Bond_Number.stringToStrSql()},  
           {item.Lots.stringToStrSql()},  
           {item.Segment_Name.stringToStrSql()}, 
           {item.Curr_Sp_Issuer.stringToStrSql()}, 
           {item.Curr_Moodys_Issuer.stringToStrSql()}, 
           {item.Curr_Fitch_Issuer.stringToStrSql()}, 
           {item.Curr_Tw_Issuer.stringToStrSql()}, 
           {item.Curr_Sp_Issue.stringToStrSql()}, 
           {item.Curr_Moodys_Issue.stringToStrSql()}, 
           {item.Curr_Fitch_Issue.stringToStrSql()}, 
           {item.Curr_Tw_Issue.stringToStrSql()}, 
           {item.Ori_Amount.stringToDblSql()}, 
           {item.Current_Int_Rate.stringToDblSql()}, 
           {item.Origination_Date.stringToDtSql()}, 
           {item.Maturity_Date.stringToDtSql()}, 
           {item.Principal_Payment_Method_Code.stringToStrSql()}, 
           {item.Payment_Frequency.stringToStrSql()}, 
           {item.Baloon_Freq.stringToStrSql()}, 
           {item.Issuer_Area.stringToStrSql()}, 
           {item.Industry_Sector.stringToStrSql()},   
           {item.Product.stringToStrSql()},   
           {item.Ias39_Category.stringToStrSql()},
           {item.Principal.stringToDblSql()},
           {item.Amort_Amt_Tw.stringToDblSql()},
           {item.Interest_Receivable.stringToDblSql()},
           {item.Int_Receivable_Tw.stringToDblSql()},
           {item.Interest_Rate_Type.stringToStrSql()},
           {item.Impair_Yn.stringToStrSql()},
           {item.Eir.stringToDblSql()},
           {item.Currency_Code.stringToStrSql()},
           '{dt.ToString("yyyy/MM/dd")}',
           {item.Issuer.stringToStrSql()},
           {item.Country_Risk.stringToStrSql()},
           {item.Ex_rate.stringToDblSql()},
           {item.Lien_position.stringToStrSql()},
           {item.Portfolio.stringToStrSql()},
           {item.Asset_Seg.stringToStrSql()},
           {item.Ori_Ex_rate.stringToDblSql()},
           {item.Bond_Type.stringToStrSql()},
           {item.Assessment_Sub_Kind.stringToStrSql()},  
           '{start.ToString("yyyy/MM/dd")}',
           {verInt.ToString()},
           {item.Bond_Aera.stringToStrSql()},
           {item.Asset_Type.stringToStrSql()},
           {item.IH_OS.stringToStrSql()},
           {item.Amt_TW_Ori_Ex_rate.stringToDblSql()},
           {item.Amort_Amt_Ori_Tw.stringToDblSql()},
           {item.Market_Value_Ori.stringToDblSql()},
           {item.Market_Value_TW.stringToDblSql()},
           {item.Value_date.stringToDtSql()},
           {item.Portfolio_Name.stringToStrSql()},
           {item.Security_Name.stringToStrSql()},
           {item.ISIN_Changed_Ind.stringToStrSql()},
           {item.Bond_Number_Old.stringToStrSql()},
           {item.Lots_Old.stringToStrSql()},
           {item.Portfolio_Name_Old.stringToStrSql()},
           {item.Origination_Date_Old.stringToDtSql()},
           {_UserInfo._user.stringToStrSql()},
           {_UserInfo._date.dateTimeToStrSql()},
           {_UserInfo._time.timeSpanToStrSql()}
); ");
                    #endregion
                    #region insert A95
                    string Security_Ticker = item.Product == null ? null :
                        item.Product.StartsWith("A11") || item.Product.StartsWith("932") ?
                        string.Format("{0} Mtge", item.Bond_Number) : string.Format("{0} Corp", item.Bond_Number);

                    string Security_Des = null;
                    string Bloomberg_Ticker = null;

                    var A53SampleInfo = A53SampleInfos.FirstOrDefault(x =>
                            x.Bond_Number == item.Bond_Number);
                    if (A53SampleInfo != null)
                    {
                        Security_Des = A53SampleInfo.Security_Des;
                        Bloomberg_Ticker = A53SampleInfo.Bloomberg_Ticker;
                    }

                    var A95_1 = A95_1Datas.FirstOrDefault(x =>
                            x.Bond_Number == item.Bond_Number);
                    if (Security_Des.IsNullOrWhiteSpace() && A95_1 != null)
                    {
                        Security_Des = A95_1.Security_Des;                     
                    }
                    if (Bloomberg_Ticker.IsNullOrWhiteSpace() && A95_1 != null)
                    {
                        Bloomberg_Ticker = A95_1.Bloomberg_Ticker;
                    }

                    //A95 Bond_Ticker_Info 
                    sb.Append(
$@" INSERT INTO [Bond_Ticker_Info]
           ([Report_Date]
           ,[Version]
           ,[Bond_Number]
           ,[Lots]
           ,[Portfolio_Name]
           ,[PRODUCT]
           ,[Lien_position]
           ,[Security_Ticker]
           ,[Security_Des]
           ,[Bloomberg_Ticker]
           ,[Bond_Type]
           ,[Assessment_Sub_Kind]
           ,[Processing_Date]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
)
     VALUES
           (
           '{dt.ToString("yyyy/MM/dd")}',
           {verInt.ToString()},
           {item.Bond_Number.stringToStrSql()},
           {item.Lots.stringToStrSql()},
           {item.Portfolio_Name.stringToStrSql()},
           {item.Product.stringToStrSql()},
           {item.Lien_position.stringToStrSql()},
           {Security_Ticker.stringToStrSql()},
           {Security_Des.stringToStrSql()},
           {Bloomberg_Ticker.stringToStrSql()},
           {item.Bond_Type.stringToStrSql()},
           {item.Assessment_Sub_Kind.stringToStrSql()},
           '{start.ToString("yyyy/MM/dd")}',
           {_UserInfo._user.stringToStrSql()},
           {_UserInfo._date.dateTimeToStrSql()},
           {_UserInfo._time.timeSpanToStrSql()}
) ; ");
                    #endregion
                }
                #region 抓最後一版A95
                string lastA95 = $@"
--A95TEMP
with TEMP2 AS
(
select 
TOP 1
Report_Date ,Version 
from Bond_Account_Info 
where 
(
Report_Date != '{dt.ToString("yyyy/MM/dd")}' 
or 
Version != {verInt.ToString()}
)
group by Report_Date ,Version
order by Report_Date DESC,Version DESC
),
TEMP AS 
(
select A95.Bond_Number,
       A95.Lots,
	   A95.Portfolio_Name,
	   A95.Security_Des,
	   A95.Bloomberg_Ticker
from   (select * from Bond_Ticker_Info where Bloomberg_Ticker is not null ) AS A95,TEMP2
where  A95.Report_Date = TEMP2.Report_Date
and    A95.Version = TEMP2.Version
),
";
                #endregion
                #region 假如這一次的A95 Security_Des,Bloomberg_Ticker,Bond_Type,Assessment_Sub_Kind is null,由A95最後版本帶入
                sb.Append(
                    lastA95 +
                    $@"
A95TEMP AS
(
select 
A95.Report_Date,
A95.Version,
A95.Lots,
A95.Bond_Number,
A95.Portfolio_Name,
T1.Bloomberg_Ticker,
T1.Security_Des,
CASE WHEN A45.Bond_Type = 'Quasi Sovereign'
     THEN '主權及國營事業債'
	 WHEN A45.Bond_Type = 'Non Quasi Sovereign'
	 THEN '其他債券'
	 ELSE A45.Bond_Type
	 END  AS Bond_Type,
CASE WHEN (A45.Assessment_Sub_Kind = '金融債' AND A95.Lien_position = '次順位' )
     THEN '金融債次順位債券'
	 WHEN (A45.Assessment_Sub_Kind = '金融債' AND( A95.Lien_position = '' OR A95.Lien_position is null) )
	 THEN '金融債主順位債券'
	 ELSE A45.Assessment_Sub_Kind
	 END AS Assessment_Sub_Kind
from  
(select * from Bond_Ticker_Info 
where Report_Date = '{dt.ToString("yyyy/MM/dd")}'
and   version = {verInt.ToString()}
and   Security_Des is null
and   Bloomberg_Ticker is null
and   Bond_Type is null
and   Assessment_Sub_Kind is null
) AS A95
JOIN TEMP T1
ON A95.Lots = T1.Lots
AND A95.Bond_Number = T1.Bond_Number
AND A95.Portfolio_Name = T1.Portfolio_Name
LEFT JOIN Bond_Category_Info A45
ON T1.Bloomberg_Ticker = A45.Bloomberg_Ticker
)
update Bond_Ticker_Info 
set Security_Des = A95TEMP.Security_Des,
    Bloomberg_Ticker = A95TEMP.Bloomberg_Ticker,
    Assessment_Sub_Kind = A95TEMP.Assessment_Sub_Kind,
    Bond_Type = A95TEMP.Bond_Type
from Bond_Ticker_Info A95, A95TEMP 
where A95.Report_Date = A95TEMP.Report_Date
and A95.Version = A95TEMP.Version
and A95.Lots = A95TEMP.Lots
and A95.Bond_Number = A95TEMP.Bond_Number
and A95.Portfolio_Name = A95TEMP.Portfolio_Name ;
");
                #endregion
                #region 假如這一次的A41 Bond_Type,Assessment_Sub_Kind is null,由A95最後版本帶入
                sb.Append(
                    lastA95 +
                    $@"
A41TEMP AS
(
select A41.Reference_Nbr,
CASE WHEN A45.Bond_Type = 'Quasi Sovereign'
     THEN '主權及國營事業債'
	 WHEN A45.Bond_Type = 'Non Quasi Sovereign'
	 THEN '其他債券'
	 ELSE A45.Bond_Type
	 END  AS Bond_Type,
CASE WHEN (A45.Assessment_Sub_Kind = '金融債' AND A41.Lien_position = '次順位' )
     THEN '金融債次順位債券'
	 WHEN (A45.Assessment_Sub_Kind = '金融債' AND( A41.Lien_position = '' OR A41.Lien_position is null) )
	 THEN '金融債主順位債券'
	 ELSE A45.Assessment_Sub_Kind
	 END AS Assessment_Sub_Kind
from (select * from Bond_Account_Info 
where Report_Date =  '{dt.ToString("yyyy/MM/dd")}'
and   version = {verInt.ToString()}
and   Bond_Type is null
and   Assessment_Sub_Kind is null ) AS A41
JOIN TEMP T1
ON A41.Lots = T1.Lots
AND A41.Bond_Number = T1.Bond_Number
AND A41.Portfolio_Name = T1.Portfolio_Name
JOIN Bond_Category_Info A45
ON T1.Bloomberg_Ticker = A45.Bloomberg_Ticker
)
update Bond_Account_Info
set Assessment_Sub_Kind = A41TEMP.Assessment_Sub_Kind,
    Bond_Type = A41TEMP.Bond_Type
FROM Bond_Account_Info A41,  A41TEMP
where A41.Reference_Nbr = A41TEMP.Reference_Nbr ;
"
                    );
                #endregion
                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    try
                    {             
                        db.Database.ExecuteSqlCommand(sb.ToString());
                        //優化 -由A41上傳中SMF為D21的券，直接帶入A42需修改的名單中(請排入後續優化) (投會)
                        var A42s = db.Treasury_Securities_Info.AsNoTracking()
                            .Where(x => x.Report_Date == dt).ToList();
                        StringBuilder sb2 = new StringBuilder();
                        db.Bond_Account_Info.AsNoTracking()
                            .Where(x => x.Report_Date == dt &&
                            x.Version == verInt &&
                            x.PRODUCT.StartsWith("D21")).ToList()
                            .ForEach(x =>
                            {
                                if (!A42s.Any(y =>
                                 y.Security_Name == x.Security_Name &&
                                 y.Lots == x.Lots &&
                                 y.Portfolio_Name == x.Portfolio_Name))
                                {
                                    double _Ori_Amount = x.Ori_Amount == null ? 0 : x.Ori_Amount.Value;
                                    double _Principal = x.Principal == null ? 0 : x.Principal.Value;
                                    sb2.Append($@"
INSERT INTO [Treasury_Securities_Info]
           ([Bond_Number]
           ,[Lots]
           ,[Segment_Name]
           ,[Portfolio_Name]
           ,[Bond_Value]
           ,[Ori_Amount]
           ,[Principal]
           ,[Amort_value]
           ,[Processing_Date]
           ,[Report_Date]
           ,[Security_Name])
     VALUES
           ({x.Bond_Number.stringToStrSql()}
           ,{x.Lots.stringToStrSql()}
           ,{x.Segment_Name.stringToStrSql()}
           ,{x.Portfolio_Name.stringToStrSql()}
           ,0
           ,{_Ori_Amount}
           ,{_Principal}
           ,0
           ,'{start.ToString("yyyy/MM/dd")}'
           ,'{dt.ToString("yyyy/MM/dd")}'
           ,{x.Security_Name.stringToStrSql()}) ;
");
                                }
                            });
                        if(sb2.Length > 0)
                        db.Database.ExecuteSqlCommand(sb2.ToString());
                        dbContextTransaction.Commit();                      
                        common.saveTransferCheck(
                               type,
                               true,
                               dt,
                               verInt,
                               start,
                               DateTime.Now
                           );
                        result.RETURN_FLAG = true;
                    
                    }
                    catch (Exception ex)
                    {
                        //dbContextTransaction.Rollback();
                        result.RETURN_FLAG = false;
                        var _message = Message_Type
                                .save_Fail.GetDescription(type, ex.exceptionMessage());
                        result.DESCRIPTION = _message;
                        common.saveTransferCheck(
                                type,
                                false,
                                dt,
                                verInt,
                                start,
                                DateTime.Now,
                                _message
                            );
                    }
                }
            }
            return result;
        }

        #endregion Save A41

        #region Save A45

        /// <summary>
        /// A45 save db
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveA45(List<A45ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                if (!dataModel.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    return result;
                }

                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    Bond_Ticker_Info bti = db.Bond_Ticker_Info.AsNoTracking()
                     .OrderByDescending(x => x.Report_Date)
                     .ThenByDescending(x => x.Version)
                     .FirstOrDefault();
                    List<Bond_Ticker_Info> listBTI = db.Bond_Ticker_Info.AsNoTracking()
                        .Where(x => x.Report_Date == bti.Report_Date
                            && x.Version == bti.Version
                            //&& x.PRODUCT != "411 Gov.CENTRAL"
                            && x.PRODUCT != "931 CDO"
                            && x.PRODUCT != "A11 AGENCY MBS"
                            && x.PRODUCT != "932 CLO"
                            //&& x.PRODUCT != "421 Gov.LOCAL"
                            && (x.PRODUCT == null || !x.PRODUCT.StartsWith("4"))
                         //開頭為4 = 國際債券 update 2018/04/23 by mark
                         ).ToList();
                    bool commitFlag = true;
                    using (var dbContextTransaction = db.Database.BeginTransaction())
                    {
                        
                        StringBuilder sb = new StringBuilder();
                        List<SqlParameter> sps = new List<SqlParameter>();
                        try
                        {
                            db.Database.ExecuteSqlCommand(@" delete Bond_Category_Info ;");
                            int i = 0;
                            bool A95Flag = false;
                            bool alreadyParm = false;
                            foreach (var A45 in dataModel)
                            {
                                i += 1;
                                sb.AppendLine($@"
INSERT INTO [Bond_Category_Info]
           ([Memo]
           ,[Bloomberg_Ticker]
           ,[Bond_Type]
           ,[Corp_NonCorp]
           ,[Sector_External]
           ,[Sector_Internal]
           ,[Assessment_Sub_Kind]
           ,[Sector_Research]
           ,[Sector_Extra]
           ,[Processing_Date])
     VALUES
           (@Memo{i}
           ,@Bloomberg_Ticker{i}
           ,@Bond_Type{i}
           ,@Corp_NonCorp{i}
           ,@Sector_External{i}
           ,@Sector_Internal{i}
           ,@Assessment_Sub_Kind{i}
           ,@Sector_Research{i}
           ,@Sector_Extra{i}
           ,@Processing_Date{i} ) ;
");
                                sps.Add(new SqlParameter($"Memo{i}", A45.Memo));
                                sps.Add(new SqlParameter($"Bloomberg_Ticker{i}", A45.Bloomberg_Ticker));
                                sps.Add(new SqlParameter($"Bond_Type{i}", A45.Bond_Type));
                                sps.Add(new SqlParameter($"Corp_NonCorp{i}", A45.Corp_NonCorp));
                                sps.Add(new SqlParameter($"Sector_External{i}", A45.Sector_External));
                                sps.Add(new SqlParameter($"Sector_Internal{i}", A45.Sector_Internal));
                                sps.Add(new SqlParameter($"Assessment_Sub_Kind{i}", A45.Assessment_Sub_Kind));
                                sps.Add(new SqlParameter($"Sector_Research{i}", A45.Sector_Research));
                                sps.Add(new SqlParameter($"Sector_Extra{i}", A45.Sector_Extra));
                                sps.Add(new SqlParameter($"Processing_Date{i}", DateTime.Parse(A45.Processing_Date)));
                                var _Bloomberg_Ticker = A45.Bloomberg_Ticker;
                                var _Bond_Ticker_Infos = listBTI.Where(x => x.Bloomberg_Ticker == _Bloomberg_Ticker).ToList();
                                var _Bond_Ticker_Info = _Bond_Ticker_Infos.FirstOrDefault();

                                if (_Bond_Ticker_Info != null)
                                {
                                    A95Flag = true;
                                    var _Bond_Type = formateBondType(A45.Bond_Type);
                                    var _Assessment_Sub_Kind = formateAssessmentSubKind(A45.Assessment_Sub_Kind, _Bond_Ticker_Info.Lien_position);
                                    sb.AppendLine(
                                        $@"
                                     update Bond_Ticker_Info
                                     set Bond_Type = @T_Bond_Type{i},
                                         Assessment_Sub_Kind = @T_Assessment_Sub_Kind{i},
                                         LastUpdate_User = 'System',
                                         LastUpdate_Date = @LastUpdate_Date,
                                         LastUpdate_Time = @LastUpdate_Time
                                     where Report_Date = @Report_Date
                                     and Version = @Version
                                     and Bloomberg_Ticker = @Bloomberg_Ticker{i} ;
                                     "
                                        );
                                    sps.Add(new SqlParameter($@"T_Bond_Type{i}", _Bond_Type));
                                    sps.Add(new SqlParameter($@"T_Assessment_Sub_Kind{i}", _Assessment_Sub_Kind));
                                    int j = 0;
                                    foreach (var A95 in _Bond_Ticker_Infos)
                                    {
                                        j += 1;
                                        sb.AppendLine(
                                 $@"     update Bond_Account_Info
                                    set Bond_Type = @T_Bond_Type{i},
                                         Assessment_Sub_Kind = @T_Assessment_Sub_Kind{i},
                                         LastUpdate_User = 'System',
                                         LastUpdate_Date = @LastUpdate_Date,
                                         LastUpdate_Time = @LastUpdate_Time
                                     where Report_Date = @Report_Date
                                     and Version = @Version
                                     and Bond_Number = @Bond_Number{i}_{j}
                                     and Lots = @Lots{i}_{j}
                                     and Portfolio_Name = @Portfolio_Name{i}_{j} ;
                                ");
                                        sps.Add(new SqlParameter($@"Bond_Number{i}_{j}", A95.Bond_Number));
                                        sps.Add(new SqlParameter($@"Lots{i}_{j}", A95.Lots));
                                        sps.Add(new SqlParameter($@"Portfolio_Name{i}_{j}", A95.Portfolio_Name));
                                    }
                                }
                                if (A95Flag && !alreadyParm)
                                {
                                    alreadyParm = true;
                                    sps.Add(new SqlParameter($@"LastUpdate_Date", _UserInfo._date));
                                    sps.Add(new SqlParameter($@"LastUpdate_Time", _UserInfo._time));
                                    sps.Add(new SqlParameter($@"Report_Date", bti.Report_Date));
                                    sps.Add(new SqlParameter($@"Version", bti.Version));
                                }
                                if (i % 30 == 0)
                                {
                                    if(sb.Length > 0)
                                    db.Database.ExecuteSqlCommand(sb.ToString(), sps.ToArray());
                                    sb = new StringBuilder();
                                    sps = new List<SqlParameter>();
                                    A95Flag = false;
                                    alreadyParm = false;
                                }
                            }
                            if(sb.Length > 0)
                            db.Database.ExecuteSqlCommand(sb.ToString(), sps.ToArray());
                        }
                        catch (Exception ex)
                        {
                            commitFlag = false;
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION =  ex.exceptionMessage();
                        }
                        if (commitFlag)
                        {
                            dbContextTransaction.Commit();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }

            return result;
        }

        #endregion

        #region Save A46
        /// <summary>
        /// Save A46
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel saveA46(List<A46ViewModel> datas)
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
                var dt = DateTime.Now.Date;
                db.Fixed_Income_CEIC_Info.RemoveRange(db.Fixed_Income_CEIC_Info.Where(x => x.Processing_Date == dt));
                foreach (var item in db.Fixed_Income_CEIC_Info.Where(x => x.Effective == "Y"))
                {
                    item.Effective = "N";
                }
                db.Fixed_Income_CEIC_Info.AddRange(datas.Select(x => 
                new Fixed_Income_CEIC_Info
                {
                    CEIC_Value = TypeTransfer.stringToDoubleN(x.CEIC_Value),
                    Country = x.Country,
                    Processing_Date = dt,
                    Effective = "Y"
                }));
                try
                {
                    var validateMessage = db.GetValidationErrors().getValidateString();
                    if (validateMessage.Any())
                    {
                        result.DESCRIPTION = validateMessage;
                        return result;
                    }
                    db.SaveChanges();
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
        #endregion

        #region Save A47
        /// <summary>
        /// Save A47
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel saveA47(List<A47ViewModel> datas)
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
                var dt = DateTime.Now.Date;
                db.Fixed_Income_Moody_Info.RemoveRange(db.Fixed_Income_Moody_Info.Where(x => x.Processing_Date == dt));
                foreach (var item in db.Fixed_Income_Moody_Info.Where(x => x.Effective == "Y"))
                {
                    item.Effective = "N";
                }
                db.Fixed_Income_Moody_Info.AddRange(datas.Select(x =>
                new Fixed_Income_Moody_Info
                {
                    Country = x.Country,
                    External_Debt_Rate = TypeTransfer.stringToDoubleN(x.External_Debt_Rate),
                    FC_Indexed_Debt_Rate = TypeTransfer.stringToDoubleN(x.FC_Indexed_Debt_Rate),
                    Processing_Date = dt,
                    Effective = "Y"
                }));
                try
                {
                    var validateMessage = db.GetValidationErrors().getValidateString();
                    if (validateMessage.Any())
                    {
                        result.DESCRIPTION = validateMessage;
                        return result;
                    }
                    db.SaveChanges();
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
        #endregion

        #region Save A48
        public MSGReturnModel saveA48(List<A48ViewModel> datas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if(!datas.Any())
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var dt = DateTime.Now.Date;
                var Country = "UNITED ARAB EMI";
                var dtyear = DateTime.Now.Year;
                db.Fixed_Income_AbuDhabi_Info.RemoveRange(db.Fixed_Income_AbuDhabi_Info.Where(x => x.Processing_Date == dt));
                foreach (var item in db.Fixed_Income_AbuDhabi_Info.Where(x => x.Effective == "Y" ))
                {
                    item.Effective = "N";
                }                 
                datas.ForEach(x =>
                {
                    int year = TypeTransfer.stringToInt(x.Data_Year);
                    db.Fixed_Income_AbuDhabi_Info.Add(
                        new Fixed_Income_AbuDhabi_Info()
                        {
                            Country = Country,
                            Data_Year = year,
                            CEIC_Value = TypeTransfer.stringToDoubleN(x.CEIC_Value),
                            External_Debt_Rate = TypeTransfer.stringToDoubleN(x.External_Debt_Rate),
                            FC_Indexed_Debt_Rate = TypeTransfer.stringToDoubleN(x.FC_Indexed_Debt_Rate),
                            Foreign_Exchange = TypeTransfer.stringToDoubleN(x.Foreign_Exchange),
                            GDP_Yearly = TypeTransfer.stringToDoubleN(x.GDP_Yearly),
                            IGS_Index = TypeTransfer.stringToDoubleN(x.IGS_Index),
                            Processing_Date = dt,
                            Short_term_Debt = TypeTransfer.stringToDoubleN(x.Short_term_Debt),
                            Effective =  "Y" 
                        });
                });
                try
                {
                    var validateMessage = db.GetValidationErrors().getValidateString();
                    if (validateMessage.Any())
                    {
                        result.DESCRIPTION = validateMessage;
                        return result;
                    }
                    db.SaveChanges();
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
        #endregion

        #region Save A95

        /// <summary>
        /// Save A95 & A41 to DB
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        public MSGReturnModel saveA95(List<A95ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime startTime = DateTime.Now;
            List<StringBuilder> sbs = new List<StringBuilder>();
            _UserInfo._date = startTime.Date;
            _UserInfo._time = startTime.TimeOfDay;
            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var reportDate = TypeTransfer.stringToDateTime(dataModel.First().Report_Date);
                int ver = TypeTransfer.stringToInt(dataModel.First().Version);
                var A95s = db.Bond_Ticker_Info.AsNoTracking()
                    .Where(x=>x.Report_Date == reportDate && x.Version == ver).ToList();
                var A45s = db.Bond_Category_Info.AsNoTracking().ToList();
                dataModel.Where(x=>!x.Security_Des.IsNullOrWhiteSpace()) //只抓Security_Des有變更的
                    .ToList().ForEach(x =>
                {
                    var A95 = A95s.FirstOrDefault(y =>
                                y.Bond_Number == x.Bond_Number &&
                                y.Lots == x.Lots &&
                                y.Portfolio_Name == x.Portfolio_Name);
                    if (A95 != null) //A95
                    {
                        if (A95.Security_Des != x.Security_Des) //不一樣才需要改
                        {
                            A95.Security_Des = x.Security_Des;
                            A95.Bloomberg_Ticker = x.Security_Des.Split(' ')[0];
                            A95.Processing_Date = startTime;
                            var A45 =  A45s.FirstOrDefault(i => i.Bloomberg_Ticker == A95.Bloomberg_Ticker);
                            if (A45 != null) //比對到的A45
                            {
                                var Assessment_Sub_Kind = formateAssessmentSubKind(A45.Assessment_Sub_Kind, x.Lien_position);
                                var Bond_Type = formateBondType(A45.Bond_Type);
                                sbs.Add(new StringBuilder($@"
UPDATE  Bond_Ticker_Info
SET Security_Des = {A95.Security_Des.stringToStrSql()} ,
    Bloomberg_Ticker = {A95.Bloomberg_Ticker.stringToStrSql()} ,
    Processing_Date = {startTime.dateTimeToStrSql()} ,
    Assessment_Sub_Kind = {Assessment_Sub_Kind.stringToStrSql()},
    Bond_Type = {Bond_Type.stringToStrSql()},
    LastUpdate_User = {_UserInfo._user.stringToStrSql()},
    LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
    LastUpdate_Time = {_UserInfo._time.timeSpanToStrSql()}
WHERE Report_Date = {reportDate.dateTimeToStrSql()}
AND  Version = {ver.ToString()}
AND  Bond_Number = {A95.Bond_Number.stringToStrSql()}
AND  Lots = {x.Lots.stringToStrSql()} 
AND  Portfolio_Name = {x.Portfolio_Name.stringToStrSql()} ;

UPDATE Bond_Account_Info
SET Assessment_Sub_Kind = {Assessment_Sub_Kind.stringToStrSql()},
    Bond_Type = {Bond_Type.stringToStrSql()} ,
    Processing_Date = {startTime.dateTimeToStrSql()} ,
    LastUpdate_User = {_UserInfo._user.stringToStrSql()},
    LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
    LastUpdate_Time = {_UserInfo._time.timeSpanToStrSql()}
WHERE Report_Date = {reportDate.dateTimeToStrSql()}
AND  Version = {ver.ToString()}
AND  Bond_Number = {A95.Bond_Number.stringToStrSql()} 
AND  Lots = {x.Lots.stringToStrSql()} 
AND  Portfolio_Name = {x.Portfolio_Name.stringToStrSql()} ;
"));
                            }
                            else
                            {
                                sbs.Add(new StringBuilder($@"
UPDATE  Bond_Ticker_Info
SET Security_Des = {A95.Security_Des.stringToStrSql()} ,
    Bloomberg_Ticker = {A95.Bloomberg_Ticker.stringToStrSql()} ,
    Processing_Date = {startTime.dateTimeToStrSql()},
    LastUpdate_User = {_UserInfo._user.stringToStrSql()},
    LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
    LastUpdate_Time = {_UserInfo._time.timeSpanToStrSql()}
WHERE Report_Date = {reportDate.dateTimeToStrSql()}
AND  Version = {ver.ToString()}
AND  Bond_Number = {A95.Bond_Number.stringToStrSql()} 
AND  Lots = {x.Lots.stringToStrSql()} 
AND  Portfolio_Name = {x.Portfolio_Name.stringToStrSql()} ;
"));
                            }
                        }
                    }
                });
                if (!sbs.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = "沒有比對到修改資料";
                    return result;
                }
                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        int size = 5000;
                        for (int q = 0; (sbs.Count() / size) >= q; q += 1)
                        {
                            StringBuilder sql = new StringBuilder();
                            sbs.Skip((q) * size).Take(size).ToList()
                                .ForEach(x =>
                                {
                                    sql.Append(x.ToString());
                                });
                            if(sql.Length > 0)
                            db.Database.ExecuteSqlCommand(sql.ToString());
                        }
                        dbContextTransaction.Commit();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                    }
                    catch (Exception ex)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }

        #endregion

        #region Save B01

        /// <summary>
        /// save B01
        /// </summary>
        /// <param name="version"></param>
        /// <param name="date">Report_Date</param>
        /// <param name="type">M = 房貸 ,B = 債券</param>
        /// <returns></returns>
        public MSGReturnModel saveB01(int version, DateTime date, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime startDt = DateTime.Now;
            try
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                    .not_Find_Any.GetDescription("B01");
                List<string> Product_Codes = new List<string>()
                {
                    Product_Code.B_A.GetDescription(),
                    Product_Code.B_B.GetDescription(),
                    Product_Code.B_P.GetDescription()
                };
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (Debt_Type.B.ToString().Equals(type)) //債券
                    {
                        string fileName = Table_Type.B01.ToString();
                        #region 檢核A41 有無資料
                        if (!db.Bond_Account_Info.AsNoTracking()
                        .Any(x => x.Report_Date != null &&
                         date == x.Report_Date.Value &&
                         version == x.Version))
                        {
                            var _message = Message_Type
                                .query_Not_Find.GetDescription(fileName);
                            result.DESCRIPTION = _message;
                            common.saveTransferCheck(
                                fileName,
                                false,
                                date,
                                version,
                                startDt,
                                DateTime.Now,
                                _message
                                );
                            return result;
                        }
                        #endregion
                        #region 檢核轉檔紀錄
                        if (!common.checkTransferCheck(fileName, Table_Type.A58.ToString(), date, version))
                        {
                            var _message = Message_Type.transferError.GetDescription(Table_Type.B01.ToString());
                            common.saveTransferCheck(
                                fileName,
                                false,
                                date,
                                version,
                                startDt,
                                DateTime.Now,
                                _message
                                );
                            result.DESCRIPTION = _message;
                            return result;
                        }
                        #endregion
                        #region 檢核有無信評缺漏狀況
                        var ratingType = Rating_Type.B.GetDescription();
                        var A58Count = db.Bond_Rating_Summary
                                         .AsNoTracking()
                                         .Where(x =>
                                         x.Report_Date == date &&
                                         x.Version == version &&
                                         x.Rating_Type == ratingType &&
                                         x.Rating_Selection != null &&
                                         x.Grade_Adjust != null)
                                         .GroupBy(x => x.Reference_Nbr).Count();
                        var A41Count = db.Bond_Account_Info.AsNoTracking()
                            .Where(x => x.Report_Date == date &&
                            x.Version == version).Count();
                        //A58 有評等資料筆數,不等於A41筆數
                        if (A58Count != A41Count)
                        {
                            result.DESCRIPTION = "報導日信評有缺漏狀況,請先進行信評補登!";
                            return result;
                        }
                        #endregion
                        #region B01(債券) 轉檔前檢核 A41
                        var _check = new BondsCheckRepository<Bond_Account_Info>(
                        db.Bond_Account_Info.AsNoTracking()
                        .Where(x => x.Report_Date == date &&
                        x.Version == version), Check_Table_Type.Bonds_B01_Before_Check);
                        if (_check.ErrorFlag)
                        {
                            var _message = Message_Type
                            .table_Check_Fail.GetDescription(Table_Type.B01.ToString());
                            result.DESCRIPTION = _message;
                            common.saveTransferCheck(
                                fileName,
                                false,
                                date,
                                version,
                                startDt,
                                DateTime.Now,
                                _message
                                );
                            return result;
                        }
                        #endregion
                        #region B01(債券) 轉檔
                        string sql = string.Empty;
                        sql += $@"
Update BA_Info
Set   BA_Info.Principal = TS_Info.Principal,
      BA_Info.Amort_Amt_Tw = TS_Info.Principal,
      BA_Info.Processing_Date = '{startDt.ToString("yyyy/MM/dd")}'
FROM  (Select * from Bond_Account_Info
where Report_Date = '{date.ToString("yyyy/MM/dd")}'
and   Version = {version.ToString()} ) As BA_Info
JOIN  (Select * from Treasury_Securities_Info
where Report_Date = '{date.ToString("yyyy/MM/dd")}' ) AS TS_Info
ON    BA_Info.Bond_Number = TS_Info.Bond_Number
AND   BA_Info.Lots = TS_Info.Lots
AND   BA_Info.Portfolio_Name = TS_Info.Portfolio_Name
AND   BA_Info.Report_Date = TS_Info.Report_Date ; 
";

                        sql += $@" 
with RTA AS
(
    select Reference_Nbr,Grade_Adjust,Rating_Priority,Rating_Selection
    from Bond_Rating_Summary
	where Report_Date = '{date.ToString("yyyy/MM/dd")}'
	and Version	 = {version.ToString()}
	AND Rating_Type = '{Rating_Type.B.GetDescription()}'
	AND Rating_Priority is not null
    AND Grade_Adjust is not null
	group by Reference_Nbr,Grade_Adjust,Rating_Priority,Rating_Selection
),
RTB AS
(
    select Reference_Nbr,Grade_Adjust,Rating_Priority,Rating_Selection
    from Bond_Rating_Summary
	where Report_Date = '{date.ToString("yyyy/MM/dd")}'
	and Version	 = {version.ToString()}
	AND Rating_Type = '{Rating_Type.A.GetDescription()}'
	AND Rating_Priority is not null
    AND Grade_Adjust is not null
	group by Reference_Nbr,Grade_Adjust,Rating_Priority,Rating_Selection
),
A62 AS
(
    select 
	LGD ,
	Lien_Position
	from Moody_LGD_Info
	where Status = '1'
	and Lien_Position = 'Senior Unsecured Bonds'
	union all
	select 
	LGD ,
	Lien_Position
	from Moody_LGD_Info
	where Status = '1'
	and Lien_Position = 'Subordinated Bonds'
)
           INSERT INTO [IFRS9_Main]
           ([Current_LGD]
           ,[Reference_Nbr]
           ,[Principal]
           ,[Interest_Receivable]
           ,[Principal_Payment_Method_Code]
           ,[Current_Int_Rate]
           ,[Eir]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Current_Rating_Code]
           ,[Report_Date]
           ,[Maturity_Date]
           ,[Current_External_Rating]
           ,[Original_External_Rating]
           ,[Version]
           ,[Lien_position]
           ,[Ori_Amount]
           ,[Payment_Frequency])
SELECT
CASE WHEN BA_Info.Lien_position is null
     THEN (select LGD from A62 where Lien_position = 'Senior Unsecured Bonds')
	 WHEN BA_Info.Lien_position = '次順位'
	 THEN (select LGD from A62 where Lien_position = 'Subordinated Bonds')
	 ELSE null 
end AS Current_LGD,
BA_Info.Reference_Nbr ,
BA_Info.Principal,
BA_Info.Interest_Receivable,
BA_Info.Principal_Payment_Method_Code,
CASE WHEN (ISNULL(BA_Info.Current_Int_Rate,0) = 0  AND ISNULL(BA_Info.EIR,0) = 0)
     THEN 0.00001   -- 修改於2018/04/09 由 0.000001 => 0.00001 與Eir一致 by mark
     WHEN(ISNULL(BA_Info.Current_Int_Rate,0) = 0  AND ISNULL(BA_Info.EIR,0) <> 0)
	 THEN BA_Info.EIR/ CAST(100 AS float)
ELSE BA_Info.Current_Int_Rate / CAST(100 AS float)
END AS Current_Int_Rate,
CASE WHEN (ISNULL(BA_Info.EIR,0) > 0)
     THEN BA_Info.EIR /  CAST(100 AS float)
ELSE 0.00001 END AS Eir,
'{startDt.ToString("yyyy/MM/dd")}' AS Processing_Date,
CASE WHEN (BA_Info.Principal_Payment_Method_Code = '01')
     THEN '{Product_Code.B_A.GetDescription()}'
     WHEN (BA_Info.Principal_Payment_Method_Code = '02')
     THEN '{Product_Code.B_B.GetDescription()}'
     WHEN (BA_Info.Principal_Payment_Method_Code = '04')
     THEN '{Product_Code.B_P.GetDescription()}'
ELSE BA_Info.Principal_Payment_Method_Code END AS Product_Code,
CAST(
(select 
CASE WHEN Rating_Selection = '1'
     THEN MIN (BR.Grade_Adjust)
	 WHEN Rating_Selection = '2'
	 THEN MAX (BR.Grade_Adjust)
END
from RTA BR
WHERE BR.Reference_Nbr = BA_Info.Reference_Nbr
AND  BR.Rating_Priority = (Select MIN(Rating_Priority) FROM RTA T WHERE T.Reference_Nbr = BA_Info.Reference_Nbr)
group by Rating_Selection) 
AS varchar(50)) 
AS Current_Rating_Code,
BA_Info.Report_Date,
BA_Info.Maturity_Date,
CAST(
(select 
CASE WHEN Rating_Selection = '1'
     THEN MIN (BR.Grade_Adjust)
	 WHEN Rating_Selection = '2'
	 THEN MAX (BR.Grade_Adjust)
END
from RTA BR
WHERE BR.Reference_Nbr = BA_Info.Reference_Nbr
AND  BR.Rating_Priority = (Select MIN(Rating_Priority) FROM RTA T WHERE T.Reference_Nbr = BA_Info.Reference_Nbr)
group by Rating_Selection) 
AS varchar(50))
AS Current_External_Rating,
CAST(
(select 
CASE WHEN Rating_Selection = '1'
     THEN MIN (BR.Grade_Adjust)
	 WHEN Rating_Selection = '2'
	 THEN MAX (BR.Grade_Adjust)
END
from RTB BR
WHERE BR.Reference_Nbr = BA_Info.Reference_Nbr
AND  BR.Rating_Priority = (Select MIN(Rating_Priority) FROM RTB T WHERE T.Reference_Nbr = BA_Info.Reference_Nbr)
group by Rating_Selection) 
AS varchar(50))
AS Original_External_Rating,
BA_Info.Version ,
BA_Info.Lien_position,
BA_Info.Ori_Amount, 
CASE WHEN(BA_Info.Payment_Frequency = 'A')
     THEN '1'
     WHEN(BA_Info.Payment_Frequency = 'S')
     THEN '2'
ELSE BA_Info.Payment_Frequency
END AS Payment_Frequency
FROM
Bond_Account_Info BA_Info
where Report_Date = '{date.ToString("yyyy/MM/dd")}'
and Version = {version.ToString()}; ";
                        db.Database.ExecuteSqlCommand(sql);
                        #endregion
                        #region 新增轉檔紀錄
                        common.saveTransferCheck(
                                   fileName,
                                   true,
                                   date,
                                   version,
                                   startDt,
                                   DateTime.Now
                                   );
                        #endregion
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type
                            .save_Success.GetDescription(fileName);
                        #region B01(債券) 轉檔後驗證
                        new BondsCheckRepository<IFRS9_Main>(
                            db.IFRS9_Main.AsNoTracking()
                            .Where(x => x.Report_Date == date &&
                            x.Version == version &&
                            Product_Codes.Contains(x.Product_Code)), Check_Table_Type.Bonds_B01_Transfer_Check);
                        #endregion
                    }
                    if (Debt_Type.M.ToString().Equals(type)) //房貸
                    {
                        //查詢A01 有無資料
                        if (db.Loan_IAS39_Info.Any(x=>x.Report_Date == date))
                        {
                            #region B01(房貸) 轉檔前檢核 A01
                            var _check = new MortgageCheckRepository<Loan_IAS39_Info>(db.Loan_IAS39_Info.AsNoTracking()
                                                            .Where(x => x.Report_Date == date),
                                                            Check_Table_Type.Mortgage_B01_Before_Check);
                            if (_check.ErrorFlag)
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = Message_Type
                                                     .table_Check_Fail.GetDescription("B01");
                                return result;
                            }
                            #endregion
                            #region B01(房貸) 轉檔
                            var _Loan_1 = Product_Code.M.GetDescription();
                            string sql = string.Empty;
                            sql = $@"
Delete IFRS9_Main 
where Report_Date = '{date.ToString("yyyy/MM/dd")}'
and  Product_Code = '{_Loan_1}' ;

INSERT INTO [IFRS9_Main]
           ([Reference_Nbr]
           ,[Principal]
           ,[Interest_Receivable]
           ,[Current_Int_Rate]
           ,[Current_Lgd]
           ,[Eir]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Current_Rating_Code]
           ,[Report_Date]
           ,[Maturity_Date]
           ,[Collateral_Legal_Action_Ind]
           ,[Delinquent_Days]
           ,[Ias39_Impaire_Ind]
           ,[Ias39_Impaire_Desc]
           ,[Restructure_Ind]
           ,[Payment_Frequency]
           ,[Version]
           ,[Ori_Amount])
select 
A01.Reference_Nbr AS Reference_Nbr,
A01.Principal AS Principal,
A01.Interest_Receivable AS Interest_Receivable,
CASE WHEN (ISNULL(A02.Current_Int_Rate,0) <= 0 AND A01.EIR = 0)
     THEN 0.000001  
     WHEN (ISNULL(A02.Current_Int_Rate,0) <= 0 AND A01.EIR <> 0)
	 THEN A01.EIR / 100
ELSE A02.Current_Int_Rate / 100
END AS Current_Int_Rate,
A02.Current_Lgd,
CASE WHEN (A01.EIR  <= 0)
     THEN 0.000001
ELSE  A01.EIR/100 
END AS Eir,
'{startDt.ToString("yyyy/MM/dd")}' AS Processing_Date,
'{_Loan_1}' AS Product_Code,
A02.Current_Rating_Code,
'{date.ToString("yyyy/MM/dd")}' AS Report_Date,
CASE WHEN (LEN(A02.Lexp_Date) = 7 )
     THEN CONVERT(date, CONVERT(varchar(4),(CONVERT(int,SUBSTRING(A02.Lexp_Date,1,3))+1911)) + '/' + SUBSTRING(A02.Lexp_Date,4,2) + '/'  +  SUBSTRING(A02.Lexp_Date,6,2))
	 WHEN (LEN(A02.Lexp_Date) = 6 )
	 THEN CONVERT(date, CONVERT(varchar(4),(CONVERT(int,SUBSTRING(A02.Lexp_Date,1,2))+1911)) + '/' + SUBSTRING(A02.Lexp_Date,3,2) + '/'  +  SUBSTRING(A02.Lexp_Date,5,2))
ELSE null
END AS Maturity_Date,
A02.Collateral_Legal_Action_Ind ,
A02.Delinquent_Days,
A01.IAS39_Impaire_Ind,
A01.IAS39_Impaire_Desc,
A02.Restructure_Ind,
CASE WHEN (A02.Pay_Frq = 'M')
     THEN 12
ELSE 0
END AS Payment_Frequency,
1,
A02.ORG_AMT
from Loan_IAS39_Info A01
LEFT JOIN  Loan_Account_Info A02
ON A01.Report_Date = A02.Report_Date
AND A01.Reference_Nbr = A02.Reference_Nbr
AND A01.Report_Date = '{date.ToString("yyyy/MM/dd")}' 
WHERE A01.Report_Date = '{date.ToString("yyyy/MM/dd")}' ; ";
                            db.Database.ExecuteSqlCommand(sql);
                            #endregion
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type
                                                 .save_Success.GetDescription(Table_Type.B01.ToString());
                            #region  B01(房貸) 轉檔後驗證
                            new MortgageCheckRepository<IFRS9_Main>(
                                db.IFRS9_Main.AsNoTracking()
                                .Where(x => x.Report_Date == date 
                                         && x.Product_Code == _Loan_1), Check_Table_Type.Mortgage_B01_Transfer_Check);
                            #endregion                                      
                        }
                        else
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("A01(Loan_IAS39_Info)");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                var _message = Message_Type
                        .save_Fail.GetDescription(Table_Type.B01.ToString(), ex.exceptionMessage());
                if (Debt_Type.B.ToString().Equals(type)) //債券
                {
                    //新增轉檔紀錄
                    common.saveTransferCheck(
                           Table_Type.B01.ToString(),
                           false,
                           date,
                           version,
                           startDt,
                           DateTime.Now,
                           _message);
                }

                result.RETURN_FLAG = false;
                result.DESCRIPTION = _message;
            }
            return result;
        }

        #endregion Save B01

        #region Save C01

        /// <summary>
        /// Save C01
        /// </summary>
        /// <param name="version"></param>
        /// <param name="date">Report_Date</param>
        /// <param name="type">M = 房貸 ,B = 債券</param>
        /// <returns></returns>
        public MSGReturnModel saveC01(int version, DateTime date, string type)
        {
            string fileName = Table_Type.C01.ToString();
            DateTime startDt = DateTime.Now;
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type
                    .not_Find_Any.GetDescription(fileName);
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.IFRS9_Main.Any())
                    {
                        #region 檢核轉檔紀錄 & 設定 Product_Code
                        List<string> reportCodes = new List<string>();
                        if (Debt_Type.M.ToString().Equals(type)) //房貸
                            reportCodes.Add(Product_Code.M.GetDescription());
                        if (Debt_Type.B.ToString().Equals(type)) //債券
                        {
                            reportCodes = new List<string> {
                               Product_Code.B_A.GetDescription(),
                               Product_Code.B_B.GetDescription(),
                               Product_Code.B_P.GetDescription()
                            };
                            #region 檢核轉檔紀錄
                            if (!common.checkTransferCheck(
                                Table_Type.C01.ToString(),
                                Table_Type.B01.ToString(),
                                date,
                                version))
                            {
                                common.saveTransferCheck(
                                    fileName,
                                    false,
                                    date,
                                    version,
                                    startDt,
                                    DateTime.Now,
                                    Message_Type.transferError.GetDescription(fileName)
                                    );
                                result.DESCRIPTION = Message_Type.transferError.GetDescription(fileName);
                                return result;
                            }
                            #endregion
                        }
                        #endregion
                        #region 檢核B01 有無資料
                        IEnumerable<IFRS9_Main> addData = null; //這次要新增的資料
                        if (Debt_Type.B.ToString().Equals(type)) //債券
                        {
                            addData = db.IFRS9_Main.AsNoTracking()
                                        .Where(x => x.Report_Date != null &&
                                        date == x.Report_Date && //抓取相同的Report_date
                                        reportCodes.Contains(x.Product_Code) &&
                                        x.Version == version);  //抓取符合的 Product_Code
                        }
                        if (Debt_Type.M.ToString().Equals(type)) //房貸
                        {
                            addData = db.IFRS9_Main.AsNoTracking()
                                        .Where(x => x.Report_Date != null &&
                                        date == x.Report_Date && //抓取相同的Report_date
                                        reportCodes.Contains(x.Product_Code));
                        }
                        if (!addData.Any())
                        {
                            result.DESCRIPTION = Message_Type
                                .query_Not_Find.GetDescription(Table_Type.B01.tableNameGetDescription());
                            return result;
                        }
                        #endregion
                        string sql = string.Empty;
                        #region C01 轉檔
                        if (Debt_Type.M.ToString().Equals(type)) //房貸
                        {
                            #region C01(房貸) 轉檔前檢核 B01
                            var _check = new MortgageCheckRepository<IFRS9_Main>(
                                             db.IFRS9_Main.AsNoTracking()
                                             .Where(x => x.Report_Date == date 
                                             && reportCodes.Contains(x.Product_Code)), Check_Table_Type.Mortgage_C01_Before_Check);
                            if (_check.ErrorFlag)
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = Message_Type
                                                    .table_Check_Fail.GetDescription("C01");
                                return result;
                            }
                            #endregion
                            #region C01(房貸) 轉檔
                            sql = $@"
                                    delete EL_Data_In
                                    where Report_Date = '{date.ToString("yyyy/MM/dd")}'
                                    and  Product_Code = '{Product_Code.M.GetDescription()}' ;

                                    INSERT INTO  [EL_Data_In]
                                               ([Report_Date]
                                               ,[Processing_Date]
                                               ,[Product_Code]
                                               ,[Reference_Nbr]
                                               ,[Current_Rating_Code]
                                               ,[Exposure]
                                               ,[Actual_Year_To_Maturity]
                                               ,[Duration_Year]
                                               ,[Remaining_Month]
                                               ,[Current_LGD]
                                               ,[Current_Int_Rate]
                                               ,[EIR]
                                               ,[Impairment_Stage]
                                               ,[Version]
                                               ,[Lien_position]
                                               ,[Ori_Amount]
                                               ,[Principal]
                                               ,[Interest_Receivable]
                                               ,[Payment_Frequency])
                                    Select
                                    B01.Report_Date, 
                                    '{startDt.ToString("yyyy/MM/dd")}',
                                    B01.Product_Code,
                                    B01.Reference_Nbr,
                                    B01.Current_Rating_Code,
                                    ISNULL(B01.Principal,0) + ISNULL(B01.Interest_Receivable,0) AS Exposure,
                                    CASE WHEN (B01.Maturity_Date is not null)
                                         THEN DATEDIFF(MONTH, B01.Report_Date ,B01.Maturity_Date )/ CAST(12 AS float)
                                    ELSE 0
                                    END AS Actual_Year_To_Maturity,
                                    CASE WHEN(B01.Remaining_Month is null)
                                         THEN(
                                         CASE WHEN (B01.Maturity_Date is not null)
                                              THEN
                                                   (CASE WHEN  (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )/CAST(12 AS float)) =0
                                                        THEN  1/CAST(12 AS float)
                                                        ELSE DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )/CAST(12 AS float)
                                                        END)
                                         ELSE 1/CAST(12 AS float) END)
                                         WHEN(B01.Remaining_Month > 0 )
                                         THEN B01.Remaining_Month/CAST(12 AS float)
                                    ELSE 1/CAST(12 AS float)
                                    END AS Duration_Year,
                                    CASE WHEN (B01.Remaining_Month is null)
                                         THEN(
                                         CASE WHEN (B01.Maturity_Date is not null)
                                              THEN
                                                   (CASE WHEN (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )) = 0
                                                   THEN 1
                                                   ELSE (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date ))
                                                   END)
                                         ELSE NULL END)
                                         WHEN(B01.Remaining_Month > 0 )
                                         THEN B01.Remaining_Month
                                    ELSE 1
                                    END AS Remaining_Month,
                                    CASE WHEN (ISNULL(B01.Current_Lgd,0) > 1 )
                                         THEN 1
	                                     WHEN (ISNULL(B01.Current_Lgd,0) < 0 )
	                                     THEN 0
                                    ELSE ISNULL(B01.Current_Lgd,0)
                                    END  AS Current_LGD,
                                    B01.Current_Int_Rate,
                                    B01.Eir AS EIR,
                                    CASE WHEN (B01.IAS39_Impaire_Ind= 'Y' AND  B01.IAS39_Impaire_Desc <> '逾期 29天' ) THEN '3'
                                         WHEN B01.Delinquent_Days>=100 THEN '3'
                                         WHEN B01.Delinquent_Days IS NULL   THEN '1'    /*帳卡上已結清*/
                                         WHEN B01.Restructure_Ind = 'Y' THEN '2'   /* 紓困名單*/
                                         WHEN B01.Collateral_Legal_Action_Ind ='Y' THEN '2'       /*假扣押名單*/
                                         WHEN B01.Delinquent_Days>29  AND B01.Delinquent_Days < 90 THEN '2'   /*7/7 金控風控建議修改*/ 
                                         WHEN B01.Delinquent_Days<30 THEN '1'                                 /*7/7 金控風控建議修改*/
                                    ELSE null
                                    END AS Impairment_Stage,
                                    1,
                                    null,
                                    B01.Ori_Amount,
                                    null,
                                    null,
                                    B01.Payment_Frequency
                                    from IFRS9_Main B01
                                    where B01.Report_Date = '{date.ToString("yyyy/MM/dd")}'
                                    AND   B01.Version = 1
                                    AND   B01.Product_Code = '{Product_Code.M.GetDescription()}' ; ";
                            #region  C01(房貸) 轉檔後驗證
                            new MortgageCheckRepository<EL_Data_In>(
                            db.EL_Data_In.AsNoTracking()
                            .Where(x => x.Report_Date == date
                             && reportCodes.Contains(x.Product_Code)), Check_Table_Type.Mortgage_C01_Transfer_Check);
                            #endregion
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.save_Success.GetDescription("C01");
                            #endregion
                        }
                        if (Debt_Type.B.ToString().Equals(type)) //債券
                        {
                            #region C01(債券) 轉檔前檢核 B01
                            var _check = new BondsCheckRepository<IFRS9_Main>(
                                    db.IFRS9_Main.AsNoTracking()
                                    .Where(x => x.Report_Date == date &&
                                    x.Version == version &&
                                    reportCodes.Contains(x.Product_Code)), Check_Table_Type.Bonds_C01_Before_Check);
                            if (_check.ErrorFlag)
                            {
                                var _message = Message_Type
                                .table_Check_Fail.GetDescription(Table_Type.C01.ToString());
                                result.DESCRIPTION = _message;
                                common.saveTransferCheck(
                                    fileName,
                                    false,
                                    date,
                                    version,
                                    startDt,
                                    DateTime.Now,
                                    _message
                                    );
                                return result;
                            }
                            #endregion
                            #region C01(債券) 轉檔
                            sql = $@"
INSERT INTO [EL_Data_In]
           ([Report_Date]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Reference_Nbr]
           ,[Current_Rating_Code]
           ,[Exposure]
           ,[Actual_Year_To_Maturity]
           ,[Duration_Year]
           ,[Remaining_Month]
           ,[Current_LGD]
           ,[Current_Int_Rate]
           ,[EIR]
           ,[Impairment_Stage]
           ,[Version]
           ,[Lien_position]
           ,[Ori_Amount]
           ,[Principal]
           ,[Interest_Receivable]
           ,[Payment_Frequency])
Select
B01.Report_Date,
'{startDt.ToString("yyyy/MM/dd")}',
B01.Product_Code,
B01.Reference_Nbr,
B01.Current_Rating_Code,
ISNULL(B01.Principal,0) + ISNULL(B01.Interest_Receivable,0) AS Exposure,
CASE WHEN(B01.Maturity_Date is not null)
     THEN DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )/CAST(12 AS float)
ELSE 0
END AS Actual_Year_To_Maturity,
CASE WHEN(B01.Remaining_Month is null)
     THEN(
     CASE WHEN (B01.Maturity_Date is not null)
          THEN
               (CASE WHEN  (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )/CAST(12 AS float)) =0
                    THEN  1/CAST(12 AS float)
                    ELSE DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )/CAST(12 AS float)
                    END)
     ELSE 1/CAST(12 AS float) END)
     WHEN(B01.Remaining_Month > 0 )
     THEN B01.Remaining_Month/CAST(12 AS float)
ELSE 1/CAST(12 AS float)
END AS Duration_Year,
CASE WHEN (B01.Remaining_Month is null)
     THEN(
     CASE WHEN (B01.Maturity_Date is not null)
          THEN
               (CASE WHEN (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date )) = 0
               THEN 1
               ELSE (DATEDIFF(MONTH, B01.Report_Date , B01.Maturity_Date ))
               END)
     ELSE NULL END)
     WHEN(B01.Remaining_Month > 0 )
     THEN B01.Remaining_Month
ELSE 1
END AS Remaining_Month,
CASE WHEN (ISNULL(B01.Current_Lgd,0) > 1 )
     THEN 1
     WHEN(ISNULL(B01.Current_Lgd,0) < 0 )
	 THEN 0
ELSE ISNULL(B01.Current_Lgd,0)
END AS Current_LGD,
B01.Current_Int_Rate,
B01.Eir AS EIR,
CASE WHEN B01.Product_Code = 'Bond_A' OR B01.Product_Code = 'Bond_B' OR B01.Product_Code = 'Bond_P'
     THEN '1'
ELSE null
END AS Impairment_Stage,
B01.Version,
B01.Lien_position,
B01.Ori_Amount,
B01.Principal,
B01.Interest_Receivable,
B01.Payment_Frequency
from IFRS9_Main B01
where B01.Report_Date = '{date.ToString("yyyy/MM/dd")}'
AND B01.Version = {version.ToString()}
AND B01.Product_Code IN ({reportCodes.stringListToInSql()}); ";
                            #endregion
                        }
                        db.Database.ExecuteSqlCommand(sql);
                        if (Debt_Type.B.ToString().Equals(type)) //債券
                        {
                            #region 新增轉檔紀錄
                            common.saveTransferCheck(
                                   fileName,
                                   true,
                                   date,
                                   version,
                                   startDt,
                                   DateTime.Now);
                            #endregion
                            #region C01(債券) 轉檔後驗證
                            new BondsCheckRepository<EL_Data_In>(
                                db.EL_Data_In.AsNoTracking()
                                .Where(x => x.Report_Date == date &&
                                x.Version == version &&
                                reportCodes.Contains(x.Product_Code)), Check_Table_Type.Bonds_C01_Transfer_Check);
                            #endregion
                        }
                        #endregion
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type
                            .save_Success.GetDescription(Table_Type.C01.ToString());
                    }
                }              
            }
            catch (Exception ex)
            {
                var _message = Message_Type
                        .save_Fail.GetDescription(Table_Type.C01.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
                if (Debt_Type.B.ToString().Equals(type)) //債券
                {
                    //新增轉檔紀錄
                    common.saveTransferCheck(
                           fileName,
                           false,
                           date,
                           version,
                           startDt,
                           DateTime.Now,
                           _message
                           );
                }
                result.RETURN_FLAG = false;
                result.DESCRIPTION = _message;
            }
            return result;
        }

        #endregion Save C01

        #region Save C02

        /// <summary>
        /// Save C02
        /// </summary>
        /// <param name="version"></param>
        /// <param name="date">Report_Date</param>
        /// <param name="type">M = 房貸</param>
        /// <returns></returns>
        public MSGReturnModel saveC02(int version, DateTime date, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type
                    .not_Find_Any.GetDescription(Table_Type.C02.ToString());
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Loan_Account_Info.Any())
                    {
                        List<string> reportCodes = new List<string>();

                        if (Debt_Type.M.ToString().Equals(type)) //房貸
                        {
                            string productCode = GroupProductCode.M.GetDescription();
                            productCode = db.Group_Product_Code_Mapping.Where(x => x.Group_Product_Code.StartsWith(productCode)).FirstOrDefault().Product_Code;
                            reportCodes.Add(productCode);
                        }

                        DateTime date13Month = date.AddMonths(-13);

                        var A02 = (from x in db.Loan_Account_Info
                                   where (x.Report_Date >= date13Month
                                         && x.Report_Date <= date)
                                   select x).ToList();

                        if (!A02.Any())
                        {
                            result.DESCRIPTION = "Loan_Account_Info 沒有 " + date13Month.ToString("yyyy/MM/dd") + " ~ " + date.ToString("yyyy/MM/dd") + " 的資料";
                            return result;
                        }

                        var _check = new MortgageCheckRepository<Loan_Account_Info>(A02.AsEnumerable(), Check_Table_Type.Mortgage_C02_Before_Check);
                        if (_check.ErrorFlag)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type
                                                .table_Check_Fail.GetDescription("C02");
                            return result;
                        }

                        string sql = $@"WITH A02 AS
                                        (
	                                        SELECT *
                                            FROM Loan_Account_Info
	                                        WHERE Report_Date >= '{date13Month.ToString("yyyy/MM/dd")}'
	                                              AND Report_Date <= '{date.ToString("yyyy/MM/dd")}'
                                         ),
                                        C02 AS
                                        (
	                                        SELECT A.* 
	                                        FROM A02 A
	                                        JOIN
	                                        (
		                                        SELECT Reference_Nbr,
			                                           Rating_Date,
			                                           MAX(Report_Date) Report_Date
		                                        FROM A02
		                                        GROUP BY Reference_Nbr,Rating_Date
	                                        ) B ON A.Reference_Nbr = B.Reference_Nbr AND A.Report_Date = B.Report_Date 
                                        )";
                        string pc = reportCodes[0];
                        db.Database.ExecuteSqlCommand(sql + $@" DELETE A FROM Rating_History A 
                                                               JOIN C02 B ON A.Reference_Nbr = B.Reference_Nbr 
                                                                          AND A.Rating_Date = replace(convert(varchar, CAST(B.Rating_Date as date), 111),'/','-')  
                                                                          AND A.Product_Code = '{pc}' "); //要問

                       
                        string nowDate = DateTime.Now.ToString("yyyy-MM-dd");

                        db.Database.ExecuteSqlCommand(sql + $@" INSERT INTO Rating_History (Data_ID,Processing_Date,Product_Code,Reference_Nbr,Current_Rating_Code,Rating_Date)
                                                                SELECT '','{nowDate}','{pc}',A.Reference_Nbr,A.Current_Rating_Code,
                                                                         replace(convert(varchar, CAST(A.Rating_Date as date), 111),'/','-') AS Rating_Date
                                                                FROM C02 A ");

                        new MortgageCheckRepository<Loan_Account_Info>(A02.AsEnumerable(), Check_Table_Type.Mortgage_C02_Transfer_Check);

                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type
                            .save_Success.GetDescription("C02");
                    }
                }             
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                        .save_Fail.GetDescription(Table_Type.C02.ToString(),
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
            }

            return result;
        }

        protected class temp
        {
            public string Reference_Nbr { get; set; }
            public string Rating_Date { get; set; }
            public DateTime Report_Date { get; set; }
            public string Current_Rating_Code { get; set; }
        }

        #endregion Save C02

        #endregion Save DB 部分

        #region Excel 部分

        #region Excel 資料轉成 A41ViewModel

        /// <summary>
        /// Excel 資料轉成 A41ViewModel
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="stream"></param>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public Tuple<string, List<A41ViewModel>> getExcel(
            string pathType,
            Stream stream,
            DateTime reportDate,
            int version)
        {
            string message = string.Empty;
            DataSet resultData = new DataSet();
            List<A41ViewModel> dataModel = new List<A41ViewModel>();
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
                int idNum = 0;
                List<SMF_Info> D53s = new List<SMF_Info>();
                List<Rating_Info_SampleInfo> A53SampleInfos = new List<Rating_Info_SampleInfo>();
                List<Assessment_Sub_Kind_Ticker> A95_1s = new List<Assessment_Sub_Kind_Ticker>();
                List<Bond_Category_Info> A45s = new List<Bond_Category_Info>();
                List<Bond_ISIN_Changed_Info> A44s = new List<Bond_ISIN_Changed_Info>();
                List<Treasury_Securities_Info> A42s = new List<Treasury_Securities_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    int i = 0;
                    D53s = db.SMF_Info.AsNoTracking().ToList();
                    A53SampleInfos = db.Rating_Info_SampleInfo.AsNoTracking()
                        .Where(x=>x.Report_Date == reportDate).ToList();
                    A95_1s = db.Assessment_Sub_Kind_Ticker.AsNoTracking().ToList();
                    A45s = db.Bond_Category_Info.AsNoTracking().ToList();
                    A44s = db.Bond_ISIN_Changed_Info.AsNoTracking().ToList();
                    A42s = db.Treasury_Securities_Info.AsNoTracking()
                        .Where(x => x.Report_Date == reportDate).ToList();
                    if (db.Bond_Account_Info.Any())
                    {
                        try
                        {
                            var Ref = db.Bond_Account_Info.AsNoTracking()
                                .Max(x => x.Reference_Nbr);
                            Int32.TryParse(Ref, out idNum);                       
                        }
                        finally
                        { 
                            if(idNum == 0)
                            idNum = db.Bond_Account_Info.AsNoTracking()
                                      .Select(x => x.Reference_Nbr).Distinct()
                                      .AsEnumerable().Where(x => Int32.TryParse(x, out i))
                                      .Max(x => Convert.ToInt32(x));
                        }
                    }

                }
                if (resultData.Tables[0].Rows.Count > 2) //判斷有無資料
                {
                    //Bond_Number & Security_Name 有一個就給過
                    var data = resultData.Tables[0].AsEnumerable()
                        .Where(x => (x.Field<object>("Bond_Number") != null &&
                                    !x.Field<object>("Bond_Number").ToString().IsNullOrWhiteSpace()) || 
                                    (x.Field<object>("Security_Name") != null &&
                                    !x.Field<object>("Security_Name").ToString().IsNullOrWhiteSpace()));

                    List<string> titles = resultData.Tables[0].Columns
                                                     .Cast<DataColumn>()
                                                     .Select(x => x.ColumnName.Trim().Replace("\t", string.Empty))
                                                     .ToList();

                    dataModel = data //第二行開始
                        .Select((x, y) =>
                        {
                            return getA41Model(x,
                                //(y + 1 + idNum).ToString().PadLeft(10, '0'),
                                titles,
                                D53s,
                                A53SampleInfos,
                                A95_1s,
                                A45s,
                                A44s,
                                A42s,
                                reportDate,
                                version);
                        }
                        ).ToList();
                    //條件一：SMF = 931 CDO
                    //條件二：攤銷後之成本數(原幣) = 0
                    //條件一 及 條件二 都成立的話，即可刪除。
                    dataModel = dataModel.Where(x => 
                    !(x.Product == "931 CDO" &&
                    TypeTransfer.stringToDouble(x.Principal) == 0)
                    ).ToList()
                        .Select((x, y) =>
                        {
                            x.Reference_Nbr = (y + 1 + idNum).ToString().PadLeft(10, '0');
                            return x;
                        }).ToList();
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            return new Tuple<string, List<A41ViewModel>>(message,dataModel);
        }

        #endregion Excel 資料轉成 A41ViewModel

        #region Excel 資料轉成 A45ViewModel

        /// <summary>
        /// Excel 資料轉成 A45ViewModel
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public List<A45ViewModel> getA45Excel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<A45ViewModel> dataModel = new List<A45ViewModel>();
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

                if (resultData.Tables[0].Rows.Count > 2) //判斷有無資料
                {
                    dataModel = (from q in resultData.Tables[0].AsEnumerable().Skip(1)
                                 select getA45ViewModel(q, DateTime.Now.ToString("yyyy/MM/dd"))).ToList();
                }
            }
            catch (Exception ex)
            {
            }

            return dataModel;
        }

        #endregion

        #region Excel 資料轉成 A46ViewModel
        /// <summary>
        /// Excel 資料轉成 A46ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<A46ViewModel>> getA46Excel(string pathType, Stream stream)
        {
            List<A46ViewModel> dataModel = new List<A46ViewModel>();
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
                    var _table = resultData.Tables[0];
                    List<string> titles = _table.Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName.Trim()).ToList();
                    if (titles.Count < 2)
                    {
                        message = @"
上傳的Excel檔案欄位數量錯誤:
格式應符合下面內容
第一欄 : 國家,
第二欄 : (經常帳+FDI淨流入)/GDP";
                    }
                    else
                    {
                        dataModel.AddRange(_table.AsEnumerable()
                        .Select(x => new A46ViewModel()
                        {
                            Country = TypeTransfer.objToString(x[0]),
                            CEIC_Value = TypeTransfer.objToString(x[1])
                        }));
                        var dt = DateTime.Now.ToString("yyyy/MM/dd");
                        foreach (var item in dataModel)
                        {
                            item.Processing_Date = dt;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any() && message.IsNullOrWhiteSpace())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<A46ViewModel>>(message, dataModel);
        }
        #endregion

        #region Excel 資料轉成 A47ViewModel
        /// <summary>
        /// Excel 資料轉成 A47ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<A47ViewModel>> getA47Excel(string pathType, Stream stream)
        {
            List<A47ViewModel> dataModel = new List<A47ViewModel>();
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
                    var _table = resultData.Tables[0];
                    List<string> titles = _table.Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName.Trim()).ToList();
                    if (titles.Count < 3)
                    {
                        message = 
@"
上傳的Excel檔案欄位數量錯誤:
格式應符合下面內容
第一欄 : 國家,
第二欄 : 外幣計價政府債券/總政府債務
第三欄 : 外人持有政府/總政府債務";
                    }
                    else
                    {
                        dataModel.AddRange(_table.AsEnumerable()
                            .Select(x => new A47ViewModel()
                            {
                                Country = TypeTransfer.objToString(x[0]),
                                FC_Indexed_Debt_Rate = TypeTransfer.objToString(x[1]),
                                External_Debt_Rate = TypeTransfer.objToString(x[2])
                            }));
                        var dt = DateTime.Now.ToString("yyyy/MM/dd");
                        foreach (var item in dataModel)
                        {
                            item.Processing_Date = dt;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any() && message.IsNullOrWhiteSpace())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<A47ViewModel>>(message, dataModel);
        }
        #endregion

        #region Excel 資料轉成 A48ViewModel
        /// <summary>
        /// Excel 資料轉成 A48ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<A48ViewModel>> getA48Excel(string pathType, Stream stream)
        {
            List<A48ViewModel> dataModel = new List<A48ViewModel>();
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
                    var _table = resultData.Tables[0];
                    List<string> titles = _table.Columns.Cast<DataColumn>()
                        .Select(x => x.ColumnName.Trim()).ToList();
                    if (titles.Count < 8)
                    {
                        message = 
@"
上傳的Excel檔案欄位數量錯誤:
格式應符合下面內容
第一欄 : 年度,
第二欄 : 政府債務/GDP
第三欄 : 外人持有政府/總政府債務
第四欄 : 外幣計價政府債券/總政府債務
第五欄 : (經常帳+FDI淨流入)/GDP
第六欄 : 短期外債
第七欄 : 外匯儲備
第八欄 : 年度GDP Y/Y";
                    }
                    else
                    {
                        dataModel.AddRange(resultData.Tables[0].AsEnumerable()
                        .Select(x => new A48ViewModel()
                        {
                            Data_Year = TypeTransfer.objToString(x[0]),
                            IGS_Index = TypeTransfer.objToString(x[1]),
                            External_Debt_Rate = TypeTransfer.objToString(x[2]),
                            FC_Indexed_Debt_Rate = TypeTransfer.objToString(x[3]),
                            CEIC_Value = TypeTransfer.objToString(x[4]),
                            Short_term_Debt = TypeTransfer.objToString(x[5]),
                            Foreign_Exchange = TypeTransfer.objToString(x[6]),
                            GDP_Yearly = TypeTransfer.objToString(x[7])
                        }));
                        //dataModel = common.getViewModel(resultData.Tables[0], Table_Type.A48)
                        //    .Cast<A48ViewModel>().ToList();
                        var dt = DateTime.Now.ToString("yyyy/MM/dd");
                        foreach (var item in dataModel)
                        {
                            item.Processing_Date = dt;
                            item.Country = "Abu Dhabi";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any() && message.IsNullOrWhiteSpace())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<A48ViewModel>>(message, dataModel);
        }
        #endregion

        #region Excel 資料轉成 A95ViewModel

        /// <summary>
        /// Excel 資料轉成 A95ViewModel
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="path"></param>
        /// <returns></returns>
        public List<A95ViewModel> getA95Excel(string pathType, string path)
        {
            List<A95ViewModel> dataModel = new List<A95ViewModel>();
            try
            {
                IWorkbook wb = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new HSSFWorkbook(stream);
                        }                            
                        break;

                    case "xlsx":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new XSSFWorkbook(stream);
                        }
                        break;
                }
                ISheet sheet = wb.GetSheetAt(0);
                DataTable dt = sheet.ISheetToDataTable(true);
                if (dt.Rows.Count > 0) //判斷有無資料
                {
                    dataModel = dt.AsEnumerable()
                        .Select((x, y) =>
                        {
                            return getA95ViewModelInExcel(x);
                        }
                        ).ToList();
                }
            }
            catch (Exception ex)
            { }

            return dataModel;
        }

        #endregion Excel 資料轉成 A95ViewModel

        #region 下載 Excel

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(A95)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            if (Excel_DownloadName.A95.ToString().Equals(type))
            {
                List<A95ViewModel> A95Data = data.Cast<A95ViewModel>().ToList();
                DataTable dt = A95Data.ToDataTable();
                result.DESCRIPTION = FileRelated.DataTableToExcel(dt, path, Excel_DownloadName.A95,new A95ViewModel().GetFormateTitles());
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }

        #endregion 下載 Excel

        #endregion Excel 部分

        #region Private Function

        #region datarow 組成 A41ViewModel

        /// <summary>
        /// datarow 組成 A41ViewModel
        /// </summary>
        /// <param name="item">每一行excel</param>
        /// <param name="num"></param>
        /// <param name="titles"></param>
        /// <param name="D53s"></param>
        /// <param name="A53SampleInfos"></param>
        /// <param name="A45s"></param>
        /// <param name="A44s"></param>
        /// <param name="A42s"></param>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        private A41ViewModel getA41Model(
            DataRow item, 
            //string num, 
            List<string> titles,
            List<SMF_Info> D53s,
            List<Rating_Info_SampleInfo> A53SampleInfos,
            List<Assessment_Sub_Kind_Ticker> A95_1s,
            List<Bond_Category_Info> A45s,
            List<Bond_ISIN_Changed_Info> A44s,
            List<Treasury_Securities_Info> A42s,
            DateTime reportDate,
            int version)
        {
            if (!titles.Any())
                return new A41ViewModel();
            var A41 = new A41ViewModel();
            var A41pro = A41.GetType().GetProperties();
            for (int i = 0; i < titles.Count; i++) //每一行所有資料
            {
                if (item[i] != null)
                {
                    string data = null;
                    if (item[i].GetType().Name.Equals("DateTime"))
                        data = TypeTransfer.objDateToString(item[i]);
                    else
                        data = TypeTransfer.objToString(item[i]);
                    if (!data.IsNullOrWhiteSpace()) //資料有值
                    {
                        var A41PInfo = A41pro.Where(x => x.Name.Trim().ToLower() == titles[i].Trim().ToLower())
                                             .FirstOrDefault();
                        if (A41PInfo != null)
                            A41PInfo.SetValue(A41, data);
                    }
                }
            }
            #region 國庫券Bond Number(ISIN code)為空值的解決方法
            if (!A41.Product.IsNullOrWhiteSpace() && A41.Product.StartsWith("D21"))
            {
                A41.Bond_Number = A41.Security_Name;
            }
            #endregion
            #region reportDate & version
            A41.Report_Date = reportDate.ToString("yyyy/MM/dd");
            A41.Version = version.ToString();
            #endregion
            #region Principal_Payment_Method_Code
            var Maturity_Date = A41.Maturity_Date;
            DateTime MDate = DateTime.MinValue;
            var Baloon_Freq = A41.Baloon_Freq;
            string Principal_Payment_Method_Code = "01";
            //Year( Maturity_Date) > = 2100  Then  '04'
            if (DateTime.TryParse(Maturity_Date, out MDate) && MDate.Year >= 2100)
            {
                Principal_Payment_Method_Code = "04";
            }
            // Baloon_Freq IN ('0') OR  Baloon_Freq IS NULL THEN '02' 
            else if (Baloon_Freq.IsNullOrWhiteSpace() || "0".Equals(Baloon_Freq))
            {
                    Principal_Payment_Method_Code = "02";
            }             
            A41.Principal_Payment_Method_Code = Principal_Payment_Method_Code;
            #endregion
            #region IH OS
            string IHOS = string.Empty;
            if (A41.IH_OS == "IH")
                IHOS = "自操";
            if (A41.IH_OS == "OS")
                IHOS = "委外";
            A41.IH_OS = IHOS;
            #endregion
            #region Asset_Type
            var Product = A41.Product;
            var SMF_Code = Product == null ? null :
                Product.IndexOf(' ') > -1 ?
                Product.Substring(0, Product.IndexOf(' ')) : Product;
            A41.Asset_Type = D53s.FirstOrDefault(x => x.SMF_Code == SMF_Code)?.Asset_Type;
            #endregion
            #region Bond_Type & Assessment_Sub_Kind & Bond_Aera 
            //  特殊 PRODUCT
            string Bond_Type = null;
            string Assessment_Sub_Kind = null;
            string Bond_Aera = A41.Currency_Code == null ? null :
                "TWD".Equals(A41.Currency_Code) ? "國內" : "國外";

            //國際債券 Product : 4開頭為 主權及國營事業債 & 主權政府債 2018/04/23 修改
            if (!A41.Product.IsNullOrWhiteSpace() && A41.Product.StartsWith("4")) 
            {
                Bond_Type = "主權及國營事業債";
                Assessment_Sub_Kind = "主權政府債";
            }
            else if (A41.Product == "931 CDO" ||
                A41.Product == "A11 AGENCY MBS")
            {
                Bond_Type = "其他債券";
                Assessment_Sub_Kind = "其他";
            }
            else if (A41.Product == "932 CLO")
            {
                Bond_Type = "其他債券";
                Assessment_Sub_Kind = "CLO";
            }
            else
            {
                var A53SampleInfo = A53SampleInfos.FirstOrDefault(x =>
                    x.Report_Date == reportDate &&
                    x.Bond_Number == A41.Bond_Number);
                if (A53SampleInfo != null && !A53SampleInfo.Bloomberg_Ticker.IsNullOrWhiteSpace())
                {
                    var A45 = A45s.FirstOrDefault(x => x.Bloomberg_Ticker == A53SampleInfo.Bloomberg_Ticker);
                    if (A45 != null)
                    {
                        Bond_Type = formateBondType(A45.Bond_Type);
                        Assessment_Sub_Kind = formateAssessmentSubKind(A45.Assessment_Sub_Kind, A41.Lien_position);
                    }
                }
                else {
                    var A95_1 = A95_1s.FirstOrDefault(x => x.Bond_Number == A41.Bond_Number);
                    if (A95_1 != null)
                    {
                        var A45 = A45s.FirstOrDefault(x => x.Bloomberg_Ticker == A95_1.Bloomberg_Ticker);
                        if (A45 != null)
                        {
                            Bond_Type = formateBondType(A45.Bond_Type);
                            Assessment_Sub_Kind = formateAssessmentSubKind(A45.Assessment_Sub_Kind, A41.Lien_position);
                        }
                    }
                }
            }
            A41.Bond_Aera = Bond_Aera;
            A41.Bond_Type = Bond_Type;
            A41.Assessment_Sub_Kind = Assessment_Sub_Kind;
            #endregion
            #region ISIN_Changed_Ind,Bond_Number_Old,Lots_Old,Portfolio_Name_Old
            var A44 = A44s.FirstOrDefault(x => x.Bond_Number_New == A41.Bond_Number &&
                                               x.Lots_New == A41.Lots &&
                                               x.Portfolio_Name_New == A41.Portfolio_Name);
            if (A44 != null)
            {
                A41.ISIN_Changed_Ind = "Y";
                A41.Bond_Number_Old = A44.Bond_Number_Old;
                A41.Lots_Old = A44.Lots_Old;
                A41.Portfolio_Name_Old = A44.Portfolio_Name_Old;
                A41.Origination_Date_Old = TypeTransfer.dateTimeNToString(A44.Origination_Date_Old);
            }
            if (A41.Assessment_Sub_Kind.IsNullOrWhiteSpace())
                A41.Assessment_Sub_Kind = "其他";
            if (A41.Bond_Type.IsNullOrWhiteSpace())
                A41.Bond_Type = "其他債券";
            #endregion
            #region Principal & Ori_Amount
            Treasury_Securities_Info A42 = null;
            //國卷
            if (!A41.Product.IsNullOrWhiteSpace() && A41.Product.StartsWith("D21"))
            {
                A42 = A42s.FirstOrDefault(x =>
                          x.Security_Name == A41.Security_Name &&
                          x.Lots == A41.Lots &&
                          x.Portfolio_Name == A41.Portfolio_Name);
            }
            //其他
            else
            {
                A42 = A42s.FirstOrDefault(x =>
                          x.Bond_Number == A41.Bond_Number &&
                          x.Lots == A41.Lots &&
                          x.Portfolio_Name == A41.Portfolio_Name);
            }
            if (A42 != null)
            {
                A41.Principal = TypeTransfer.doubleNToString(A42.Principal);
                A41.Ori_Amount = TypeTransfer.doubleNToString(A42.Ori_Amount);
            }
            #endregion
            return A41;
        }

        #endregion datarow 組成 A41ViewModel

        #region Db 組成 A41ViewModel

        /// <summary>
        /// Db 組成 A41ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private A41ViewModel DbToA41Model(Bond_Account_Info data)
        {
            return new A41ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr.PadLeft(10, '0'),
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Segment_Name = data.Segment_Name,
                Curr_Sp_Issuer = data.CURR_SP_Issuer,
                Curr_Moodys_Issuer = data.CURR_Moodys_Issuer,
                Curr_Fitch_Issuer = data.CURR_Fitch_Issuer,
                Curr_Tw_Issuer = data.CURR_TW_Issuer,
                Curr_Sp_Issue = data.CURR_SP_Issue,
                Curr_Moodys_Issue = data.CURR_Moodys_Issue,
                Curr_Fitch_Issue = data.CURR_Fitch_Issue,
                Curr_Tw_Issue = data.CURR_TW_Issue,
                Ori_Amount = TypeTransfer.doubleNToString(data.Ori_Amount),
                Current_Int_Rate = TypeTransfer.doubleNToString(data.Current_Int_Rate),
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Maturity_Date = TypeTransfer.dateTimeNToString(data.Maturity_Date),
                Principal_Payment_Method_Code = data.Principal_Payment_Method_Code,
                Payment_Frequency = data.Payment_Frequency,
                Baloon_Freq = data.Baloon_Freq,
                Issuer_Area = data.ISSUER_AREA,
                Industry_Sector = data.Industry_Sector,
                Product = data.PRODUCT,
                Ias39_Category = data.IAS39_CATEGORY,
                Principal = TypeTransfer.doubleNToString(data.Principal),
                Amort_Amt_Tw = TypeTransfer.doubleNToString(data.Amort_Amt_Tw),
                Interest_Receivable = TypeTransfer.doubleNToString(data.Interest_Receivable),
                Int_Receivable_Tw = TypeTransfer.doubleNToString(data.Interest_Receivable_Tw),              
                Interest_Rate_Type = data.Interest_Rate_Type,
                Impair_Yn = data.IMPAIR_YN,
                Eir = TypeTransfer.doubleNToString(data.EIR),
                Currency_Code = data.Currency_Code,
                Report_Date = TypeTransfer.dateTimeNToString(data.Report_Date),
                Issuer = data.ISSUER,
                Country_Risk = data.Country_Risk,
                Ex_rate = TypeTransfer.doubleNToString(data.Ex_rate),
                Lien_position = data.Lien_position,
                Portfolio = data.Portfolio,
                Asset_Seg = data.ASSET_SEG,
                Ori_Ex_rate = TypeTransfer.doubleNToString(data.Ori_Ex_rate),
                Bond_Type = data.Bond_Type,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Processing_Date = TypeTransfer.dateTimeNToString(data.Processing_Date),
                Version = TypeTransfer.intNToString(data.Version),
                Bond_Aera = data.Bond_Aera,
                Asset_Type = data.Asset_Type,
                IH_OS = data.IH_OS,
                Amt_TW_Ori_Ex_rate = TypeTransfer.doubleNToString(data.Amount_TW_Ori_Ex_rate),
                Amort_Amt_Ori_Tw = TypeTransfer.doubleNToString(data.Amort_Amt_Ori_Tw),
                Market_Value_Ori = TypeTransfer.doubleNToString(data.Market_Value_Ori),
                Market_Value_TW = TypeTransfer.doubleNToString(data.Market_Value_TW),
                Value_date = TypeTransfer.dateTimeNToString(data.Value_date),
                Portfolio_Name = data.Portfolio_Name,
                ISIN_Changed_Ind = data.ISIN_Changed_Ind,
                Bond_Number_Old = data.Bond_Number_Old,
                Lots_Old = data.Lots_Old,
                Portfolio_Name_Old = data.Portfolio_Name_Old,
                Security_Name = data.Security_Name,
                Origination_Date_Old = TypeTransfer.dateTimeNToString(data.Origination_Date_Old)
            };
        }

        #endregion Db 組成 A41ViewModel

        #region Db 組成 A45ViewModel

        /// <summary>
        /// Db 組成 A45ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private A45ViewModel DbToA45Model(Bond_Category_Info data)
        {
            return new A45ViewModel()
            {
                Memo = data.Memo,
                Bloomberg_Ticker = data.Bloomberg_Ticker,
                Bond_Type = data.Bond_Type,
                Corp_NonCorp = data.Corp_NonCorp,
                Sector_External = data.Sector_External,
                Sector_Internal = data.Sector_Internal,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind,
                Sector_Research = data.Sector_Research,
                Sector_Extra = data.Sector_Extra,
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
            };
        }

        #endregion

        #region Db 組成 A95ViewModel

        /// <summary>
        /// Db 組成 A41ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private A95ViewModel DbToA95Model(Bond_Ticker_Info data)
        {
            return new A95ViewModel()
            {
                Report_Date = TypeTransfer.dateTimeNToString(data.Report_Date),
                Version = TypeTransfer.intNToString(data.Version),
                Bond_Number = data.Bond_Number,
                Lots= data.Lots,
                Portfolio_Name = data.Portfolio_Name,
                PRODUCT = data.PRODUCT,
                Lien_position = data.Lien_position,
                Security_Ticker = data.Security_Ticker,
                Security_Des = data.Security_Des,
                Bloomberg_Ticker = data.Bloomberg_Ticker,
                Bond_Type = data.Bond_Type,
                Assessment_Sub_Kind = data.Assessment_Sub_Kind
            };
        }

        #endregion Db 組成 A95ViewModel

        #region datarow 組成 A45ViewModel

        /// <summary>
        /// datarow 組成 A45ViewModel
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>A45ViewModel</returns>
        private A45ViewModel getA45ViewModel(DataRow item, string processingDate)
        {
            return new A45ViewModel()
            {
                Memo = TypeTransfer.objToString(item[0]),
                Bloomberg_Ticker = TypeTransfer.objToString(item[1]),
                Bond_Type = TypeTransfer.objToString(item[2]),
                Corp_NonCorp = TypeTransfer.objToString(item[3]),
                Sector_External = TypeTransfer.objToString(item[4]),
                Sector_Internal = TypeTransfer.objToString(item[5]),
                Assessment_Sub_Kind = TypeTransfer.objToString(item[6]),
                Sector_Research = TypeTransfer.objToString(item[7]),
                Sector_Extra = TypeTransfer.objToString(item[8]),
                Processing_Date = processingDate
            };
        }

        #endregion

        #region Excel 組成 A95ViewModel

        private A95ViewModel getA95ViewModelInExcel(DataRow item)
        {
            return new A95ViewModel()
            {
                Report_Date = TypeTransfer.objToString(item[0]),
                Version = TypeTransfer.objToString(item[1]),
                Bond_Number = TypeTransfer.objToString(item[2]),
                Lots = TypeTransfer.objToString(item[3]),
                Portfolio_Name = TypeTransfer.objToString(item[4]),
                PRODUCT = TypeTransfer.objToString(item[5]),
                Lien_position = TypeTransfer.objToString(item[6]),
                Security_Ticker = TypeTransfer.objToString(item[7]),
                Security_Des = TypeTransfer.objToString(item[8]),
                Bloomberg_Ticker = TypeTransfer.objToString(item[9]),
                Bond_Type = TypeTransfer.objToString(item[10]),
                Assessment_Sub_Kind = TypeTransfer.objToString(item[11]),
            };
        }

        #endregion Excel 組成 A95ViewModel

        #region A95 formate
        /// <summary>
        /// formate A95 BondType
        /// </summary>
        /// <param name="bondType"></param>
        /// <returns></returns>
        private string formateBondType(string bondType)
        {
            if (bondType.IsNullOrWhiteSpace())
                return bondType;
            if (bondType.Trim() == "Quasi Sovereign")
                return "主權及國營事業債";
            if (bondType.Trim() == "Non Quasi Sovereign")
                return "其他債券";
            return bondType;
        }

        /// <summary>
        /// formate A95 AssessmentSubKind
        /// </summary>
        /// <param name="assessmentSubKind"></param>
        /// <param name="lienPosition"></param>
        /// <returns></returns>
        private string formateAssessmentSubKind(string assessmentSubKind, string lienPosition)
        {
            if (assessmentSubKind.IsNullOrWhiteSpace())
                return assessmentSubKind;
            if (assessmentSubKind.Trim() == "金融債")
            {
                if (lienPosition.IsNullOrEmpty())
                    return "金融債主順位債券";
                if (lienPosition.Trim() == "次順位")
                    return "金融債次順位債券";
            }
            return assessmentSubKind;
        }
        #endregion

        #endregion Private Function

        #region getA44All
        public Tuple<bool, List<A44ViewModel>> getA44All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_ISIN_Changed_Info.Any())
                {
                    return new Tuple<bool, List<A44ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Bond_ISIN_Changed_Info.AsNoTracking()
                                                    .OrderByDescending(x=>x.Processing_Date).AsEnumerable()
                                        select DbToA44Model(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<A44ViewModel>>(true, new List<A44ViewModel>());
        }
        #endregion

        #region DbToA44Model
        private A44ViewModel DbToA44Model(Bond_ISIN_Changed_Info data)
        {
            return new A44ViewModel()
            {
                Bond_Number_New = data.Bond_Number_New,
                Lots_New = data.Lots_New,
                Portfolio_Name_New = data.Portfolio_Name_New,
                Bond_Number_Old = data.Bond_Number_Old,
                Lots_Old = data.Lots_Old,
                Portfolio_Name_Old = data.Portfolio_Name_Old,
                Issuer_Ticker_Old = data.Issuer_Ticker_Old,
                Guarantor_Name_Old = data.Guarantor_Name_Old,
                Guarantor_EQY_Ticker_Old = data.Guarantor_EQY_Ticker_Old,
                Change_Date = data.Change_Date.ToString("yyyy/MM/dd"),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Origination_Date_Old = TypeTransfer.dateTimeNToString(data.Origination_Date_Old)
            };
        }
        #endregion

        #region  getA44
        public Tuple<bool, List<A44ViewModel>> getA44(A44ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_ISIN_Changed_Info.Any())
                {
                    var query = from q in db.Bond_ISIN_Changed_Info.AsNoTracking()
                                select q;

                    if (dataModel.Bond_Number_New.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Bond_Number_New.Contains(dataModel.Bond_Number_New));
                    }

                    if (dataModel.Lots_New.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Lots_New.Contains(dataModel.Lots_New));
                    }

                    if (dataModel.Portfolio_Name_New.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Portfolio_Name_New.Contains(dataModel.Portfolio_Name_New));
                    }

                    if (dataModel.Bond_Number_Old.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Bond_Number_Old.Contains(dataModel.Bond_Number_Old));
                    }

                    if (dataModel.Lots_Old.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Lots_Old.Contains(dataModel.Lots_Old));
                    }

                    if (dataModel.Portfolio_Name_Old.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Portfolio_Name_Old.Contains(dataModel.Portfolio_Name_Old));
                    }

                    if (dataModel.Change_Date.IsNullOrWhiteSpace() == false)
                    {
                        DateTime changeDate = DateTime.Parse(dataModel.Change_Date);
                        query = query.Where(x => x.Change_Date == changeDate);
                    }

                    return new Tuple<bool, List<A44ViewModel>>(query.Any(), query.OrderByDescending(x => x.Processing_Date)
                                                                                 .AsEnumerable()
                                                                                 .Select(x => { return DbToA44Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<A44ViewModel>>(false, new List<A44ViewModel>());
        }
        #endregion

        #region saveA44
        public MSGReturnModel saveA44(string actionType, A44ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Rating_Info_SampleInfo ratingInfoSampleInfo = db.Rating_Info_SampleInfo.AsNoTracking()
                                                                    .Where(x => x.Bond_Number == dataModel.Bond_Number_Old)
                                                                    .OrderByDescending(x => x.Report_Date)
                                                                    .FirstOrDefault();
                    if (ratingInfoSampleInfo == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"Rating_Info_SampleInfo 沒有 {dataModel.Bond_Number_Old} 的資料";
                        return result;
                    }

                    Bond_Account_Info A41 = db.Bond_Account_Info
                                              .Where(x => x.Bond_Number == dataModel.Bond_Number_Old
                                                       && x.Lots == dataModel.Lots_Old
                                                       && x.Portfolio_Name == dataModel.Portfolio_Name_Old)
                                              .OrderByDescending(x => x.Report_Date)
                                              .ThenByDescending(x => x.Version)
                                              .FirstOrDefault();
                    if (A41 == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "您輸入的 債券編號_舊 、 Lots_舊 、 Portfolio英文_舊 在 A41：Bond_Account_Info 找不到";
                        return result;
                    }

                    Bond_ISIN_Changed_Info editData = new Bond_ISIN_Changed_Info();

                    if (actionType == "Add")
                    {
                        List<Bond_ISIN_Changed_Info> A44List = db.Bond_ISIN_Changed_Info.AsNoTracking().ToList();

                        if (A44List
                            .Any(x => x.Bond_Number_New == dataModel.Bond_Number_New
                                    && x.Lots_New == dataModel.Lots_New
                                    && x.Portfolio_Name_New == dataModel.Portfolio_Name_New))
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：您輸入的 債券編號_新、Lots_新、Portfolio英文_新 已存在";
                            return result;
                        }

                        if (A44List
                            .Any(x => x.Bond_Number_Old == dataModel.Bond_Number_Old
                                    && x.Lots_Old == dataModel.Lots_Old
                                    && x.Portfolio_Name_Old == dataModel.Portfolio_Name_Old))
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：您輸入的 債券編號_舊、Lots_舊、Portfolio英文_舊 已存在";
                            return result;
                        }


                        if (A44List
                            .Any(x => x.Bond_Number_Old == dataModel.Bond_Number_New
                                     && x.Lots_Old == dataModel.Lots_New
                                     && x.Portfolio_Name_Old == dataModel.Portfolio_Name_New))
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "您輸入的 債券編號_新、Lots_新、Portfolio英文_新，不可與目前存在的 債券編號_舊、Lots_舊、Portfolio英文_舊 重複";
                            return result;
                        }

                        editData.Bond_Number_New = dataModel.Bond_Number_New;
                        editData.Lots_New = dataModel.Lots_New;
                        editData.Portfolio_Name_New = dataModel.Portfolio_Name_New;
                    }
                    else if (actionType == "Modify")
                    {
                        editData = db.Bond_ISIN_Changed_Info
                                     .Where(x => x.Bond_Number_New == dataModel.Bond_Number_New
                                              && x.Lots_New == dataModel.Lots_New
                                              && x.Portfolio_Name_New == dataModel.Portfolio_Name_New)
                                     .FirstOrDefault();
                    }

                    editData.Bond_Number_Old = dataModel.Bond_Number_Old;
                    editData.Lots_Old = dataModel.Lots_Old;
                    editData.Portfolio_Name_Old = dataModel.Portfolio_Name_Old;
                    editData.Issuer_Ticker_Old = ratingInfoSampleInfo.ISSUER_TICKER;
                    editData.Guarantor_Name_Old = ratingInfoSampleInfo.GUARANTOR_NAME;
                    editData.Guarantor_EQY_Ticker_Old = ratingInfoSampleInfo.GUARANTOR_EQY_TICKER;
                    editData.Change_Date = DateTime.Parse(dataModel.Change_Date);
                    editData.Processing_Date = DateTime.Now.Date;
                    editData.Origination_Date_Old = A41.Origination_Date;

                    var _message = string.Empty;
                    if (actionType == "Add")
                    {
                        db.Bond_ISIN_Changed_Info.Add(editData);
                        #region update A41,A57,A58
                        string _sql = @"
with temp as
(
select top 1 Report_Date,Version from Bond_Rating_Summary
where Rating_Type = '原始投資信評'
group by Report_Date,Version
order by Report_Date desc,Version desc
),
temp2 as 
(
select * from temp
union all
select top 1 A58.Report_Date,A58.Version from temp tp,Bond_Rating_Summary A58
where A58.Report_Date < tp.Report_Date
and A58.Rating_Type = '原始投資信評'
group by A58.Report_Date,A58.Version
order by Report_Date desc,Version desc
)
select * from temp2
order by temp2.Report_Date desc
";
                        var _data = db.Database.DynamicSqlQuery(_sql);
                        if (_data.Any() && _data.Count == 2)
                        {
                            List<Bond_Rating_Info> addA57s = new List<Bond_Rating_Info>();
                            List<Bond_Rating_Summary> addA58s = new List<Bond_Rating_Summary>();
                            var _first = _data.First();
                            var _last = _data.Last();
                            DateTime _newReportDate = TypeTransfer.objDateToDateTime(_first.Report_Date);
                            int _newVersion = TypeTransfer.objToInt(_first.Version);
                            DateTime _oldReportDate = TypeTransfer.objDateToDateTime(_last.Report_Date);
                            int _oldVersion = TypeTransfer.objToInt(_last.Version);
                            var newA41 = db.Bond_Account_Info.FirstOrDefault(x =>
                            x.Report_Date == _newReportDate &&
                            x.Version == _newVersion &&
                            x.Bond_Number == editData.Bond_Number_New &&
                            x.Lots == editData.Lots_New &&
                            x.Portfolio_Name == editData.Portfolio_Name_New
                            );
                            if (newA41 != null)
                            {
                                newA41.Bond_Number_Old = editData.Bond_Number_Old;
                                newA41.Lots_Old = editData.Lots_Old;
                                newA41.Portfolio_Name_Old = editData.Portfolio_Name_Old;
                                newA41.ISIN_Changed_Ind = "Y";
                                newA41.Origination_Date_Old = editData.Origination_Date_Old;
                                newA41.LastUpdate_User = "System";
                                newA41.LastUpdate_Date = _UserInfo._date;
                                newA41.LastUpdate_Time = _UserInfo._time;
                            }
                            var _ratingType = Rating_Type.A.GetDescription();
                            var newA57s = db.Bond_Rating_Info.Where(x =>
                            x.Report_Date == _newReportDate &&
                            x.Version == _newVersion &&
                            x.Bond_Number == editData.Bond_Number_New &&
                            x.Lots == editData.Lots_New &&
                            x.Portfolio_Name == editData.Portfolio_Name_New &&
                            x.Rating_Type == _ratingType).ToList();
                            var oldA57s = db.Bond_Rating_Info.AsNoTracking().Where(x =>
                            x.Report_Date == _oldReportDate &&
                            x.Version == _oldVersion &&
                            x.Bond_Number == editData.Bond_Number_Old &&
                            x.Lots == editData.Lots_Old &&
                            x.Portfolio_Name == editData.Portfolio_Name_Old &&
                            x.Rating_Type == _ratingType).ToList();
                            var newA58s = db.Bond_Rating_Summary.Where(x =>
                            x.Report_Date == _newReportDate &&
                            x.Version == _newVersion &&
                            x.Bond_Number == editData.Bond_Number_New &&
                            x.Lots == editData.Lots_New &&
                            x.Portfolio_Name == editData.Portfolio_Name_New &&
                            x.Rating_Type == _ratingType).ToList();
                            var oldA58s = db.Bond_Rating_Summary.AsNoTracking().Where(x =>
                            x.Report_Date == _oldReportDate &&
                            x.Version == _oldVersion &&
                            x.Bond_Number == editData.Bond_Number_Old &&
                            x.Lots == editData.Lots_Old &&
                            x.Portfolio_Name == editData.Portfolio_Name_Old &&
                            x.Rating_Type == _ratingType).ToList();
                            if (newA57s.Any() && oldA57s.Any() && newA58s.Any() && oldA58s.Any())
                            {
                                var _newA57First = newA57s.First();
                                oldA57s.ForEach(x =>
                                {
                                    var copyA57 = _newA57First.ModelConvert<Bond_Rating_Info, Bond_Rating_Info>();
                                    copyA57.Rating_Object = x.Rating_Object; //評等對象
                                    copyA57.Rating_Org = x.Rating_Org; //評等機構
                                    copyA57.Rating = x.Rating; //評等內容
                                    copyA57.Rating_Org_Area = x.Rating_Org_Area; //國內/國外
                                    copyA57.Parm_ID = x.Parm_ID; //信評優先選擇參數編號
                                    copyA57.PD_Grade = x.PD_Grade; //評等主標尺_原始
                                    copyA57.Grade_Adjust = x.Grade_Adjust; //評等主標尺_轉換
                                    copyA57.RTG_Bloomberg_Field = x.RTG_Bloomberg_Field;
                                    copyA57.Bond_Number_Old = editData.Bond_Number_Old;
                                    copyA57.Lots_Old = editData.Lots_Old;
                                    copyA57.Portfolio_Name_Old = editData.Portfolio_Name_Old;
                                    copyA57.Origination_Date_Old = editData.Origination_Date_Old;
                                    copyA57.ISIN_Changed_Ind = "Y";
                                    copyA57.Create_User = "System";
                                    copyA57.Create_Date = _UserInfo._date;
                                    copyA57.Create_Time = _UserInfo._time;
                                    copyA57.LastUpdate_User = null;
                                    copyA57.LastUpdate_Date = null;
                                    copyA57.LastUpdate_Time = null;
                                    addA57s.Add(copyA57);
                                });
                                var _newA58First = newA58s.First();
                                oldA58s.ForEach(x =>
                                {
                                    var copyA58 = _newA58First.ModelConvert<Bond_Rating_Summary, Bond_Rating_Summary>();
                                    copyA58.Parm_ID = x.Parm_ID;
                                    copyA58.Rating_Object = x.Rating_Object; //評等對象
                                    copyA58.Rating_Org_Area = x.Rating_Org_Area; //國內/國外
                                    copyA58.Rating_Selection = x.Rating_Selection; //1:孰高 2:孰低
                                    copyA58.Grade_Adjust = x.Grade_Adjust; //評等主標尺_轉換
                                    copyA58.Rating_Priority = x.Rating_Priority; //優先順序
                                    copyA58.Processing_Date = _UserInfo._date;
                                    copyA58.Bond_Number_Old = x.Bond_Number_Old;
                                    copyA58.Lots_Old = x.Lots_Old;
                                    copyA58.Portfolio_Name_Old = x.Portfolio_Name_Old;
                                    copyA58.Origination_Date_Old = x.Origination_Date_Old;
                                    copyA58.ISIN_Changed_Ind = "Y";
                                    copyA58.Create_User = "System";
                                    copyA58.Create_Date = _UserInfo._date;
                                    copyA58.Create_Time = _UserInfo._time;
                                    copyA58.LastUpdate_User = null;
                                    copyA58.LastUpdate_Date = null;
                                    copyA58.LastUpdate_Time = null;
                                    addA58s.Add(copyA58);
                                });
                                db.Bond_Rating_Info.RemoveRange(newA57s);
                                db.Bond_Rating_Info.AddRange(addA57s);
                                db.Bond_Rating_Summary.RemoveRange(newA58s);
                                db.Bond_Rating_Summary.AddRange(addA58s);
                                _message = $" A57債券信評檔 更新{addA57s.Count}筆,A58債券信評檔_整理檔 更新{addA58s.Count}筆";
                            }
                        }
                        #endregion
                    }
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription("A44換券資訊檔") + _message;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }

            return result;
        }
        #endregion

        #region deleteA44
        public MSGReturnModel deleteA44(string bondNumberNew, string lotsNew, string portfolioNameNew)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Bond_ISIN_Changed_Info
                                  .Where(x=> x.Bond_Number_New == bondNumberNew
                                          && x.Lots_New == lotsNew
                                          && x.Portfolio_Name_New == portfolioNameNew);

                    db.Bond_ISIN_Changed_Info.RemoveRange(query);

                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region  getA49
        public Tuple<bool, List<A49ViewModel>> getA49(A49ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Accounting_EL.Any())
                {
                    var query = from q in db.Bond_Accounting_EL.AsNoTracking()
                                select q;

                    if (dataModel.Report_Date.IsNullOrWhiteSpace() == false)
                    {
                        DateTime reportDate = DateTime.Parse(dataModel.Report_Date);
                        query = query.Where(x => x.Report_Date == reportDate);
                    }

                    return new Tuple<bool, List<A49ViewModel>>(query.Any(), 
                                                               query.AsEnumerable()
                                                                    .Select(x => { return DbToA49Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<A49ViewModel>>(false, new List<A49ViewModel>());
        }
        #endregion

        #region DbToA49Model
        private A49ViewModel DbToA49Model(Bond_Accounting_EL data)
        {
            return new A49ViewModel()
            {
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Portfolio = data.Portfolio,
                Accounting_EL = data.Accounting_EL.ToString(),
                Product_Code = data.Product_Code,
                Reference_Nbr = data.Reference_Nbr,
                Impairment_Stage = data.Impairment_Stage,
                Version = data.Version.ToString(),
                IAS39_CATEGORY = data.IAS39_CATEGORY,
                Grade_Adjust = data.Grade_Adjust.ToString(),
                Portfolio_Name = data.Portfolio_Name,
                Low_Grade = data.Low_Grade.ToString(),
                High_Grade = data.High_Grade.ToString(),
                Risk_Level = data.Risk_Level.ToString()
            };
        }
        #endregion

        #region getA49Excel
        public List<A49ViewModel> getA49Excel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<A49ViewModel> dataModel = new List<A49ViewModel>();
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
                    var data = resultData.Tables[0].AsEnumerable()
                    .Where(x => (x.Field<object>("債券編號") != null &&
                                !x.Field<object>("債券編號").ToString().IsNullOrWhiteSpace()));
                    dataModel = (from q in data
                                 select getA49ViewModel(q)).ToList();
                }
            }
            catch (Exception ex)
            {
            }

            return dataModel;
        }

        #endregion

        #region getA49ViewModel
        private A49ViewModel getA49ViewModel(DataRow item)
        {
            return new A49ViewModel()
            {
                Bond_Number = TypeTransfer.objToString(item[0]),
                Lots = TypeTransfer.objToString(item[1]),
                Portfolio = TypeTransfer.objToString(item[2]),
                Accounting_EL = TypeTransfer.objToString(item[3])
            };
        }
        #endregion

        #region Save A49
        public MSGReturnModel saveA49(List<A49ViewModel> dataModel, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                if (!dataModel.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    return result;
                }

                DateTime dateReport = DateTime.Parse(reportDate);

                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    List<Bond_Account_Info> A41s = db.Bond_Account_Info.AsNoTracking()
                                                     .Where(x=>x.Report_Date == dateReport)
                                                     .ToList();
                    if (A41s.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A41 (Bond_Account_Info) 沒有 基準日={reportDate} 的資料";
                        return result;
                    }

                    int version = A41s.DefaultIfEmpty().Max(x => x.Version == null ? 0 : x.Version.Value);
                    A41s = A41s.Where(x => x.Version == version).ToList();

                    List<IFRS9_EL> D54s = (
                                              from a in db.IFRS9_EL.AsNoTracking().ToList()
                                              join b in A41s
                                              on new { a.Reference_Nbr }
                                              equals new { b.Reference_Nbr }
                                              select a
                                          ).ToList();
                    if (D54s.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"D54 (IFRS9_EL) 沒有與A41相關聯的資料";
                        return result;
                    }
                    var _Rating_Type = Rating_Type.B.GetDescription();
                    List<Bond_Rating_Summary> A58s = (
                                                          from a in db.Bond_Rating_Summary.AsNoTracking()
                                                          .Where(x => 
                                                          x.Report_Date == dateReport &&
                                                          x.Rating_Type == _Rating_Type)
                                                          .ToList()
                                                          join b in A41s
                                                          on new { a.Reference_Nbr }
                                                          equals new { b.Reference_Nbr }
                                                          select a
                                                     ).ToList();
                    if (A58s.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A58 (Bond_Rating_Summary) 沒有與A41相關聯的資料";
                        return result;
                    }

                    string thisYear = DateTime.Now.Year.ToString();

                    List<Grade_Moody_Info> A51s = db.Grade_Moody_Info.AsNoTracking()
                                                    .Where(x=>x.Status == "1")
                                                    .ToList();
                    if (A51s.Any() == false)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A51 (Grade_Moody_Info) 沒有資料";
                        return result;
                    }

                    // 7 = A-
                    //Grade_Moody_Info A51_7 = A51s
                    //                         .Where(x => x.PD_Grade == 7)
                    //                         .FirstOrDefault();
                    //if (A51_7 == null)
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = $"A51 (Grade_Moody_Info) 沒有 PD_Grade=7 的資料";
                    //    return result;
                    //}

                    // 10 = BBB-
                    Grade_Moody_Info A51_10 = A51s
                         .Where(x => x.PD_Grade == 10)
                         .FirstOrDefault();
                    if (A51_10 == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A51 (Grade_Moody_Info) 沒有 PD_Grade=10 的資料";
                        return result;
                    }

                    // 15 = B
                    //Grade_Moody_Info A51_15 = A51s
                    //                         .Where(x => x.PD_Grade == 15)
                    //                         .FirstOrDefault();
                    //if (A51_15 == null)
                    //{
                    //    result.RETURN_FLAG = false;
                    //    result.DESCRIPTION = $"A51 (Grade_Moody_Info) 沒有 PD_Grade=15 的資料";
                    //    return result;
                    //}

                    // 16 = B-
                    Grade_Moody_Info A51_16 = A51s
                         .Where(x => x.PD_Grade == 16)
                         .FirstOrDefault();
                    if (A51_16 == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = $"A51 (Grade_Moody_Info) 沒有 PD_Grade=16 的資料";
                        return result;
                    }

                    db.Database.ExecuteSqlCommand(string.Format(@"DELETE FROM Bond_Accounting_EL 
                                                                  WHERE Report_Date = '{0}'", reportDate));

                    DateTime nowDate = DateTime.Now.Date;

                    StringBuilder sb = new StringBuilder();

                    for (int i = 0; i < dataModel.Count; i++)
                    {
                        var Report_Date = reportDate;
                        var Processing_Date = nowDate.ToString("yyyy/MM/dd");
                        var Bond_Number = dataModel[i].Bond_Number;
                        var Lots = dataModel[i].Lots;
                        var Portfolio = dataModel[i].Portfolio;
                        var Accounting_EL = dataModel[i].Accounting_EL;

                        var A41 = A41s.Where(x => x.Bond_Number == Bond_Number
                                               && x.Lots == Lots
                                               && x.Portfolio == Portfolio).FirstOrDefault();
                        if (A41 != null)
                        {
                            var Reference_Nbr = A41.Reference_Nbr;

                            var D54 = D54s.Where(x => x.Reference_Nbr == Reference_Nbr).FirstOrDefault();
                            if (D54 == null)
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = $"D54 (IFRS9_EL) 沒有 Reference_Nbr={Reference_Nbr} 的資料";
                                return result;
                            }

                            var Product_Code = D54.Product_Code;
                            var Impairment_Stage = D54.Impairment_Stage;
                            var Version = A41.Version;
                            var IAS39_CATEGORY = A41.IAS39_CATEGORY;

                            var A58 = A58s.Where(x => x.Reference_Nbr == Reference_Nbr).OrderBy(z=>z.Rating_Priority).FirstOrDefault();
                            if (A58 == null)
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = $"A58 (Bond_Rating_Summary) 沒有 Reference_Nbr={Reference_Nbr} 的資料";
                                return result;
                            }

                            var Grade_Adjust = A58.Grade_Adjust;
                            var Portfolio_Name = A41.Portfolio_Name;
                            var Low_Grade = A51_10.Grade_Adjust;
                            var High_Grade = A51_16.Grade_Adjust;

                            var Risk_Level = "";

                            if (Grade_Adjust <= Low_Grade)
                            {
                                Risk_Level = "低風險";
                            }
                            else if (Grade_Adjust >= High_Grade)
                            {
                                Risk_Level = "高風險";
                            }
                            else
                            {
                                Risk_Level = "中風險";
                            }

                            sb.Append(
                                         $@" INSERT INTO Bond_Accounting_EL(
                                                                           Report_Date,
                                                                           Processing_Date,
                                                                           Bond_Number,
                                                                           Lots,
                                                                           Portfolio,
                                                                           Accounting_EL,
                                                                           Product_Code,
                                                                           Reference_Nbr,
                                                                           Impairment_Stage,
                                                                           Version,
                                                                           IAS39_CATEGORY,
                                                                           Grade_Adjust,
                                                                           Portfolio_Name,
                                                                           Low_Grade,
                                                                           High_Grade,
                                                                           Risk_Level
                                                                       )
                                                                VALUES (
                                                                           '{Report_Date}',
                                                                           '{Processing_Date}',
                                                                           '{Bond_Number}',
                                                                           '{Lots}',
                                                                           '{Portfolio}',
                                                                           '{Accounting_EL}',
                                                                           '{Product_Code}',
                                                                           '{Reference_Nbr}',
                                                                           '{Impairment_Stage}',
                                                                           '{Version}',
                                                                           '{IAS39_CATEGORY}',
                                                                           '{Grade_Adjust}',
                                                                           '{Portfolio_Name}',
                                                                           '{Low_Grade}',
                                                                           '{High_Grade}',
                                                                           '{Risk_Level}'
                                                                       );
                                       "

                                );
                        }
                    }
                    if (sb.Length > 0)
                    {
                        db.Database.ExecuteSqlCommand(sb.ToString());
                        result.RETURN_FLAG = true;
                    }
                    else
                        result.DESCRIPTION = "無更新任何資料";
                }               
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return result;
        }
        #endregion
    }
}