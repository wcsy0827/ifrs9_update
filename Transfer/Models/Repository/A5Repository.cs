using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static AutoTransfer.Enum.Ref;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A5Repository : IA5Repository
    {
        #region 其他

        //補正方式代碼 01:置換特殊字元成空值
        private List<string> rule_1s = new List<string>();
        //補正方式代碼 02:以新值取代舊值 item1 = 特殊字元 ex:(mx),item2 = 新取代字元 ex:(twn)
        private List<Tuple<string, string>> rule_2s = new List<Tuple<string, string>>();

        private List<IFRS9_User> users = new List<IFRS9_User>();

        public A5Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
            users = common.getAllUsers();
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

        #region 190222單獨將執行信評轉檔後驗證功能拉出來
        public void GetA58TransferCheck(DateTime dt, int version)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                new BondsCheckRepository<Bond_Rating_Summary>(
                    db.Bond_Rating_Summary.AsNoTracking()
                    .Where(x => x.Report_Date == dt &&
                    x.Version == version).AsEnumerable(), Check_Table_Type.Bonds_A58_Transfer_Check, dt, version);
            }
        }
        #endregion

        #region 190222單獨把A59Filled datatable儲存成EXCEL檔案功能拉出來
        public MSGReturnModel SaveA59Excel(string type, string path, List<A59ViewModel> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DataTable dt = data.ToDataTable();
            result.DESCRIPTION = FileRelated.DataTableToExcel(dt, path, Excel_DownloadName.A59Filled);  //成功儲存傳回空字串，失敗傳回catch的ex
            result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION); //如果沒有內容表示成功，判斷空字串為true
            return result;
        }
        #endregion

        #region 190222 單獨把A59Save結果寫入Log功能拉出來
        public void SaveA59TransLog(MSGReturnModel result_autotrasfer, string datepicker, DateTime starttime, int ver)
        {
            DateTime reportdate = DateTime.MinValue;
            DateTime.TryParse(datepicker, out reportdate);
            if (result_autotrasfer.RETURN_FLAG)
            {
                bool A59TransFlag = common.checkTransferCheck(Table_Type.A59Trans.ToString(), "A59Trans", reportdate, ver);
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    string sql = string.Empty;
                    sql += $@"UPDATE Transfer_CheckTable Set TransferType = 'R' 
                    WHERE ReportDate = '{datepicker.Replace("/", "-")}'
                    AND Version ={ver}
                    AND TransferType = 'Y'
                    AND File_Name ='{Table_Type.A59Trans.ToString()}' ;";
                    db.Database.ExecuteSqlCommand(sql);
                }
            }
            common.saveTransferCheck(Table_Type.A59Trans.ToString(), result_autotrasfer.RETURN_FLAG, reportdate, ver, starttime, DateTime.Now, result_autotrasfer.DESCRIPTION);
        }
        #endregion

        #endregion 其他

        #region Get Data

        public Tuple<bool, List<A51ViewModel>> getA51(
            Audit_Type type,
            string dataYear, 
            string rating = null, 
            string pdGrade = null, 
            string ratingAdjust = null,
            string gradeAdjust = null, 
            string moodysPD = null)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _pdGrade = TypeTransfer.stringToIntN(pdGrade);
                var _gradeAdjust = TypeTransfer.stringToIntN(gradeAdjust);
                var _moodysPD = TypeTransfer.stringToDoubleN(moodysPD);
                var data = db.Grade_Moody_Info.AsNoTracking()
                    .Where(x => x.Status == null,type == Audit_Type.None)
                    .Where(x => x.Status == "1",type == Audit_Type.Enable)
                    .Where(x => x.Status == "2",type == Audit_Type.TempDisabled)
                    .Where(x => x.Status == "3",type == Audit_Type.Disabled)
                    .Where(x => x.Data_Year == dataYear, !dataYear.IsNullOrWhiteSpace())
                    .Where(x => x.Rating == rating, !rating.IsNullOrWhiteSpace())
                    .Where(x => x.PD_Grade == _pdGrade, _pdGrade.HasValue)
                    .Where(x => x.Rating_Adjust == ratingAdjust, !ratingAdjust.IsNullOrWhiteSpace())
                    .Where(x => x.Grade_Adjust == _gradeAdjust, _gradeAdjust.HasValue)
                    .Where(x => x.Moodys_PD == _moodysPD, _moodysPD.HasValue);
                return new Tuple<bool, List<A51ViewModel>>(true, data.AsEnumerable().Select(x => { return getA51ViewModel(x);}).ToList());
            }
        }

        public bool getA53(
            DateTime date
            )
        {
            bool flag = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                flag = db.Rating_Info.AsNoTracking()
                    .Any(x => x.Report_Date == date);
            }
            return flag;
        }

        #region getA52All

        public Tuple<bool, List<A52ViewModel>> getA52All()
        {
            return getA52();
        }

        #endregion getA52All

        #region getA52

        /// <summary>
        /// get A52 
        /// </summary>
        /// <param name="ratingOrg">RatingOrg</param>
        /// <param name="pdGrade">PdGrade</param>
        /// <param name="rating">Rating</param>
        /// <param name="IsActive">是否有效(All:全部,Y:生效,N:失效)</param>
        /// <param name="Status">複核結果(All:全部)</param>
        /// <returns></returns>
        public Tuple<bool, List<A52ViewModel>> getA52(string ratingOrg = "All", string pdGrade = "All", string rating = "All", string IsActive = "All", string Status = "All")
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var result = db.Grade_Mapping_Info.AsNoTracking()
                                           .Where(x => x.Rating_Org == ratingOrg, ratingOrg != "All")
                                           .Where(x => x.PD_Grade.ToString() == pdGrade, pdGrade != "All")
                                           .Where(x => x.Rating == rating, rating != "All")
                                           .Where(x => x.IsActive == IsActive, IsActive != "All")
                                           .Where(x => x.Status == Status, Status != "All")
                                           .OrderBy(x => x.Rating_Org)
                                           .ThenBy(x => x.PD_Grade)
                                           .ThenBy(x => x.Rating)
                                           .AsEnumerable()
                                           .Select(x => { return getA52ViewModel(x); }).ToList();
                return new Tuple<bool, List<A52ViewModel>>(result.Any(), result);
            }
        }

        #endregion getA52

        #region getA52byAuditdate
        public List<A52ViewModel> GetA52byAuditdate(string Audit_date)
        {
            var result = new List<A52ViewModel>();
            DateTime? dt = TypeTransfer.stringToDateTimeN(Audit_date);
            var users = common.getAllUsers();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result.AddRange(db.Grade_Mapping_Info.AsNoTracking()
                    .Where(x => x.Audit_date == dt, dt != null)
                    .AsEnumerable()
                    .Select(x => {return getA52ViewModel(x);}));
            }
            return result;
        }
        #endregion

        #region getA52 複核時間
        public List<SelectOption> GetA52Auditdate(string Audit_date)
        {
            var data = new List<SelectOption>(){ new SelectOption(){ Text = "全部(All)", Value = "All" } };
            DateTime? _Audit_date1 = TypeTransfer.stringToDateTimeN(Audit_date);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                data.AddRange(db.Grade_Mapping_Info.AsNoTracking()
                    .Where(x => x.Audit_date != null)
                    .Where(x => x.Audit_date == _Audit_date1, _Audit_date1.HasValue)
                    .Select(x => x.Audit_date)
                    .Distinct()
                    .OrderByDescending(x => x)
                    .AsEnumerable()
                    .Select(x => new SelectOption()
                    {
                        Text = x.Value.ToString("yyyy/MM/dd"),
                        Value =x.Value.ToString("yyyy/MM/dd")
                    }));
            }
            return data;
        }
        #endregion

        #region getA52 畫面查詢
        public Tuple<List<SelectOption>, List<SelectOption>, List<SelectOption>> GetA52SearchData(string ratingOrg, string IsActive,string pdGrade ,string rating)
        {
            var All = new SelectOption() { Text = "全部(All)", Value = "All" };
            List<SelectOption> _Rating_Orgs = new List<SelectOption>() { All };
            List<SelectOption> _PD_Grades = new List<SelectOption>() { All };
            List<SelectOption> _Ratings = new List<SelectOption>() { All };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A52s = db.Grade_Mapping_Info.AsNoTracking()
                    .Where(x => x.Rating_Org == ratingOrg, ratingOrg != "All")
                    .Where(x => x.IsActive == IsActive, IsActive != "All")
                    .Where(x => x.PD_Grade.ToString() == pdGrade, pdGrade != "All")
                    .Where(x => x.Rating == rating, rating != "All")
                    .ToList();
                var _A52s = db.Grade_Mapping_Info.AsNoTracking().ToList();
                if (ratingOrg != "All" && pdGrade == "All" && rating == "All" && IsActive == "All")
                    _Rating_Orgs.AddRange(_A52s.GroupBy(x => x.Rating_Org)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x, Value = x }));
                else
                    _Rating_Orgs.AddRange(A52s.GroupBy(x => x.Rating_Org)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x, Value = x }));
                if (ratingOrg == "All" && pdGrade != "All" && rating == "All" && IsActive == "All")
                    _PD_Grades.AddRange(_A52s.GroupBy(x => x.PD_Grade)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x.ToString(), Value = x.ToString() }));
                else
                    _PD_Grades.AddRange(A52s.GroupBy(x => x.PD_Grade)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x.ToString(), Value = x.ToString() }));
                if (ratingOrg == "All" && pdGrade == "All" && rating != "All" && IsActive == "All")
                    _Ratings.AddRange(_A52s.GroupBy(x => x.Rating)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x, Value = x }));
                else
                    _Ratings.AddRange(A52s.GroupBy(x => x.Rating)
                    .Select(x => x.Key).Distinct().OrderBy(x => x)
                    .Select(x => new SelectOption() { Text = x, Value = x }));
            }
            return new Tuple<List<SelectOption>, List<SelectOption>, List<SelectOption>>(_Rating_Orgs, _PD_Grades, _Ratings);
        }
        #endregion

        /// <summary>
        /// Get A56
        /// </summary>
        /// <param name="IsActive">是否生效</param>
        /// <param name="Replace_Object">特殊字元</param>
        /// <returns></returns>
        public List<A56ViewModel> GetA56(string IsActive, string Replace_Object)
        {
            List<A56ViewModel> result = new List<A56ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Rating_Update_Info.AsNoTracking()
                            .Where(x => x.IsActive == IsActive, !IsActive.IsNullOrWhiteSpace())
                            .Where(x => x.Replace_Object == Replace_Object, !Replace_Object.IsNullOrWhiteSpace())
                            .OrderBy(x=>x.ID)
                            .AsEnumerable()                           
                            .Select(x => new A56ViewModel()
                            {
                                ID = x.ID.ToString(),
                                IsActive = x.IsActive,
                                Replace_Object = x.Replace_Object,
                                Update_Method = x.Update_Method,
                                Method_MEMO = x.Method_MEMO,
                                Replace_Word = x.Replace_Word,
                                LastUpdate_User = x.LastUpdate_User,
                                LastUpdate_DateTime = TypeTransfer.dateTimeNTimeSpanNToString(x.LastUpdate_Date,x.Create_Time) 
                            }).ToList();
            }
            return result;
        }

        /// <summary>
        /// Get A56 Replace_Object
        /// </summary>
        /// <returns></returns>
        public List<string> GetA56_Replace_Object()
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Rating_Update_Info.AsNoTracking()
                            .Where(x => x.IsActive == "Y").Select(x => x.Replace_Object)
                            .OrderBy(x=>x).ToList();
            }
            return result;
        }

        /// <summary>
        /// get A57 Data
        /// </summary>
        /// <param name="datepicker">Report_Date</param>
        /// <param name="version">Version</param>
        /// <param name="df">Origination_Date 開始時間</param>
        /// <param name="dt">Origination_Date 結束時間</param>
        /// <param name="SMF">SMF</param>
        /// <param name="stype">Rating_Type查詢 All(為全查)</param>
        /// <param name="bondNumber">Bond_Number</param>
        /// <param name="issuer">ISSUER</param>
        /// <param name="status">查詢信評缺漏狀態</param>
        /// <returns></returns>
        public List<A57ViewModel> GetA57(
            DateTime datepicker,
            int version,
            DateTime? df,
            DateTime? dt,
            string SMF,
            string stype,
            string bondNumber,
            string issuer,
            Rating_Status status)
        {
            List<A57ViewModel> result = new List<A57ViewModel>();

            Rating_Status type = Rating_Status.All;
            if (status == Rating_Status.AllNull)
                type = Rating_Status.BondsNull | Rating_Status.GUARANTORNull | Rating_Status.ISSUERNull;
            else
                type = status;

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var data = db.Bond_Rating_Info.AsNoTracking()
                    .Where(x => x.Report_Date == datepicker && x.Version == version)
                    .Where(x => x.Origination_Date != null &&
                                x.Origination_Date >= df, df.HasValue)
                    .Where(x => x.Origination_Date != null &&
                                x.Origination_Date <= dt, dt.HasValue)
                    .Where(x => x.SMF == SMF, !SMF.IsNullOrWhiteSpace())
                    .Where(x => x.Rating_Type == stype, stype != "All")
                    .Where(x => x.Bond_Number == bondNumber, !bondNumber.IsNullOrWhiteSpace())
                    .Where(x => x.ISSUER == issuer, !issuer.IsNullOrWhiteSpace());
                  
                List<string> refs = new List<string>();
                if (data.Any())
                {
                    var datas = data.AsEnumerable();
                    if ((type & Rating_Status.All) != Rating_Status.All) //不是全部搜尋要判斷條件
                    {
                        datas.GroupBy(x => x.Reference_Nbr).ToList() //信評判斷
                            .ForEach(x =>
                            {
                                //債項無信評
                                if (((type & Rating_Status.BondsNull) == Rating_Status.BondsNull) &&
                                     (!x.Any(y => y.Rating_Object == RatingObject.Bonds.GetDescription() && y.Grade_Adjust != null)))
                                    refs.Add(x.Key);
                                //發行人無信評
                                if (((type & Rating_Status.ISSUERNull) == Rating_Status.ISSUERNull) &&
                                    (!x.Any(y => y.Rating_Object == RatingObject.ISSUER.GetDescription() && y.Grade_Adjust != null)))
                                    refs.Add(x.Key);
                                //保證人無信評
                                if (((type & Rating_Status.GUARANTORNull) == Rating_Status.GUARANTORNull) &&
                                    (!x.Any(y => y.Rating_Object == RatingObject.GUARANTOR.GetDescription() && y.Grade_Adjust != null)))
                                    refs.Add(x.Key);
                            });
                        if (status == Rating_Status.AllNull) //完全無信評要以上三個都符合
                            refs = refs.GroupBy(z => z).Where(z => z.Count() == 3).Select(z => z.Key).ToList();
                        datas = datas.Where(x => refs.Contains(x.Reference_Nbr));
                    }
                    if (datas.Any())
                    {
                        result = datas.OrderBy(x => x.Reference_Nbr).Select(x => DbToA57ViewModel(x)).ToList();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// get A58 Data
        /// </summary>
        /// <param name="datepicker">ReportDate</param>
        /// <param name="sType">Rating_Type</param>
        /// <param name="from">Origination_Date start</param>
        /// <param name="to">Origination_Date to</param>
        /// <param name="bondNumber">bondNumber</param>
        /// <param name="version">version</param>
        /// <param name="search">全部or缺漏</param>
        /// <param name="portfolio">portfolio</param>
        /// <returns></returns>
        public Tuple<bool, List<A58ViewModel>> GetA58(
            string datepicker,
            string sType,
            string from,
            string to,
            string bondNumber,
            string version,
            string search,
            string portfolio)
        {
            DateTime? dp = TypeTransfer.stringToDateTimeN(datepicker);
            DateTime? df = TypeTransfer.stringToDateTimeN(from);
            DateTime? dt = TypeTransfer.stringToDateTimeN(to);
            int ver = 0;
            Int32.TryParse(version.Trim(), out ver);
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var data = db.Bond_Rating_Summary.AsNoTracking()
                    .Where(x => x.Report_Date == dp.Value &&
                                x.Version == ver)
                    .Where(x => x.Origination_Date != null &&
                                x.Origination_Date >= df, df.HasValue)
                    .Where(x => x.Origination_Date != null &&
                                x.Origination_Date <= dt, dt.HasValue)
                    .Where(x => x.Rating_Type == sType, !sType.IsNullOrWhiteSpace())
                    .Where(x => x.Bond_Number == bondNumber, !bondNumber.IsNullOrWhiteSpace())
                    .Where(x => x.Portfolio.IndexOf(portfolio) > -1, portfolio != "All");

                var result = data.AsEnumerable();
                if ("Miss".Equals(search)) //找缺漏
                {
                    var group = from item in result
                                group item by new { item.Reference_Nbr, item.Rating_Type } into g
                                where g.Any(x => x.Grade_Adjust != null)
                                select g; //無缺漏(有資料)
                    if (group.Any())
                    {
                        var refs = group.Select(y => y.Key.Reference_Nbr).ToList();
                        result = result.Where(x => !refs.Contains(x.Reference_Nbr)); //把有資料的排除掉
                    }
                }
                if (result.Any())
                {
                    return new Tuple<bool, List<A58ViewModel>>(true,
                        result.OrderBy(x => x.Reference_Nbr).Select(x => { return getA58ViewModel(x); }).ToList());
                }
                return new Tuple<bool, List<A58ViewModel>>(false, new List<A58ViewModel>());
            }
        }

        /// <summary>
        ///  get 轉檔紀錄Table 資料
        /// </summary>
        /// <param name="fileNames">A41,A53...</param>
        /// <returns></returns>
        public List<CheckTableViewModel> getCheckTable(List<string> fileNames)
        {
            List<CheckTableViewModel> data = new List<CheckTableViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (fileNames.Any() && db.Transfer_CheckTable.AsNoTracking().Any(x => fileNames.Contains(x.File_Name)))
                    data = db.Transfer_CheckTable.AsNoTracking().Where(x =>
                    fileNames.Contains(x.File_Name))
                    .OrderByDescending(x => x.Create_date)
                    .ThenByDescending(x => x.Create_time)
                    .AsEnumerable()
                    .Select(x => new CheckTableViewModel()
                    {
                        ReportDate = x.ReportDate.ToString("yyyy/MM/dd"),
                        Version = x.Version.ToString(),
                        TransferType = x.TransferType,
                        Create_Date = x.Create_date,
                        Create_Time = x.Create_time,
                        End_Date = x.End_date,
                        End_Time = x.End_time,
                        File_Name = $"{x.File_Name} ({x.File_Name.tableNameGetDescription()})",
                        ErrorMsg = x.Error_Msg
                    }).ToList();
            }

            return data;
        }

        /// <summary>
        /// get SMF
        /// </summary>
        /// <param name="date">Report_Date</param>
        /// <param name="version">Version</param>
        /// <returns></returns>
        public List<string> getSMF(DateTime date, int version)
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Bond_Rating_Info.AsNoTracking().Where(x =>
                    x.Report_Date == date &&
                    x.Version == version &&
                    x.SMF != null)
                    .Select(x => x.SMF).Distinct()
                    .OrderBy(x => x).ToList();
            }
            
            return result;
        }

        /// <summary>
        /// 檢查是不是該基準日最後一版(重複執行信評轉檔最後一版時使用)
        /// </summary>
        /// <param name="date">Report_Date</param>
        /// <param name="version">Version</param>
        /// <returns></returns>
        public MSGReturnModel checkVersion(DateTime date, int version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var ver = db.Bond_Rating_Info.AsNoTracking()
                     .Where(x => x.Report_Date == date)
                     .Max(x => x.Version);
                if (ver.HasValue && ver.Value == version)
                    result.RETURN_FLAG = true;
            }
            return result;
        }

        #region 190222GetA59
        public Tuple<bool, List<A59ViewModel>> getA59(List<A58ViewModel> datatable, string reportdate)
        {
            DateTime startTime = DateTime.Now;
            DateTime rd = DateTime.MinValue;
            DateTime.TryParse(reportdate, out rd);

            #region 建立A59格式
            var rt = Rating_Type.A.GetDescription();
            List<A59ViewModel> A59Data =
                    (from q in datatable.Cast<A58ViewModel>()
                     group q by new
                     {
                         q.Reference_Nbr,
                         q.Report_Date,
                         q.Version,
                         q.Bond_Number,
                         q.Lots,
                         q.Origination_Date,
                         q.Portfolio_Name,
                         q.Portfolio,
                         q.SMF,
                         q.Issuer,
                         q.Security_Ticker,
                         q.RATING_AS_OF_DATE_OVERRIDE,
                         q.Rating_Type,
                         q.ISIN_Changed_Ind,
                         q.Bond_Number_Old,
                         q.Origination_Date_Old
                     } into item
                     select new A59ViewModel()
                     {
                         Reference_Nbr = item.Key.Reference_Nbr,
                         Report_Date = item.Key.Report_Date,
                         Version = item.Key.Version,
                         Bond_Number = item.Key.Bond_Number,
                         Lots = item.Key.Lots,
                         Origination_Date =
                         item.Key.Rating_Type == rt ?
                         item.Key.ISIN_Changed_Ind == "Y" ?
                         item.Key.Origination_Date_Old : item.Key.Origination_Date : item.Key.Origination_Date,
                         Portfolio_Name = item.Key.Portfolio_Name,
                         Portfolio = item.Key.Portfolio,
                         SMF = item.Key.SMF,
                         Issuer = item.Key.Issuer,
                         Security_Ticker = item.Key.Security_Ticker,
                         RATING_AS_OF_DATE_OVERRIDE = item.Key.RATING_AS_OF_DATE_OVERRIDE,
                         Rating_Type = item.Key.Rating_Type,
                         ISIN_Changed_Ind = item.Key.ISIN_Changed_Ind,
                         Bond_Number_Old = item.Key.Bond_Number_Old
                     }).ToList();
            #endregion

            try
            {
                #region 串投資平台DB資料
                //19.02.14測試直接連線方式直接抓Linked Server 的SP
                DataTable dttempIFRS9 = new DataTable();
                List<A59ViewModel> A59dataIFRS = new List<A59ViewModel>();
                //"Data Source=10.42.71.139,3301;Initial Catalog=IFRS9;User ID=FBLKRISK;Password=Password1"
                using (SqlConnection ConnIFRS9 = new SqlConnection(ConfigurationManager.ConnectionStrings["Apex"].ConnectionString))
                {
                    using (SqlCommand cmdIFRS9 = new SqlCommand($"exec [FBL_DB].[FBL_DB_SIT].dbo.GetCounterPartyCreditRatingForIFRS9 '{reportdate.Replace("/", "-")}'", ConnIFRS9))
                    {
                        using (SqlDataAdapter adaptIFRS9 = new SqlDataAdapter(cmdIFRS9))
                        {
                            cmdIFRS9.CommandType = CommandType.Text;
                            ConnIFRS9.Open();
                            adaptIFRS9.Fill(dttempIFRS9);
                            if (dttempIFRS9.Rows.Count == 0)
                            {
                                common.saveTransferCheck(Table_Type.A59Apex.ToString(), false, rd, 0, startTime, DateTime.Now, Message_Type.not_Find_CounterPartyCreditRating.GetDescription());
                                return new Tuple<bool, List<A59ViewModel>>(false, new List<A59ViewModel>());
                            }
                            #region 將SP取出資料放入A59ViewModel
                            foreach (DataRow dr in dttempIFRS9.Rows)
                            {
                                A59ViewModel obj = new A59ViewModel();

                                obj.Bond_Number = dr["ISINCode"].ToString();
                                obj.Origination_Date = dr["TradeDate"].ToString();

                                obj.RTG_SP = dr["RTG_SP"].ToString();
                                obj.SP_EFF_DT = dr["SP_EFF_DT"].ToString();
                                obj.RTG_TRC = dr["RTG_TRC"].ToString();
                                obj.TRC_EFF_DT = dr["TRC_EFF_DT"].ToString();
                                obj.RTG_MOODY = dr["RTG_MOODY"].ToString();
                                obj.MOODY_EFF_DT = dr["MOODY_EFF_DT"].ToString();
                                obj.RTG_FITCH = dr["RTG_FITCH"].ToString();
                                obj.FITCH_EFF_DT = dr["FITCH_EFF_DT"].ToString();
                                obj.RTG_FITCH_NATIONAL = dr["RTG_FITCH_NATIONAL"].ToString();
                                obj.RTG_FITCH_NATIONAL_DT = dr["RTG_FITCH_NATIONAL_DT"].ToString();

                                obj.RTG_SP_LT_FC_ISSUER_CREDIT = dr["RTG_SP_LT_FC_ISSUER_CREDIT"].ToString();
                                obj.RTG_SP_LT_FC_ISS_CRED_RTG_DT = dr["RTG_SP_LT_FC_ISS_CRED_RTG_DT"].ToString();
                                obj.RTG_SP_LT_LC_ISSUER_CREDIT = dr["RTG_SP_LT_LC_ISSUER_CREDIT"].ToString();
                                obj.RTG_SP_LT_LC_ISS_CRED_RTG_DT = dr["RTG_SP_LT_FC_ISS_CRED_RTG_DT"].ToString();
                                obj.RTG_MDY_ISSUER = dr["RTG_MDY_ISSUER"].ToString();
                                obj.RTG_MDY_ISSUER_RTG_DT = dr["RTG_MDY_ISSUER_RTG_DT"].ToString();
                                obj.RTG_MOODY_LONG_TERM = dr["RTG_MOODY_LONG_TERM"].ToString();
                                obj.RTG_MOODY_LONG_TERM_DATE = dr["RTG_MOODY_LONG_TERM_DATE"].ToString();
                                obj.RTG_MDY_SEN_UNSECURED_DEBT = dr["RTG_MDY_SEN_UNSECURED_DEBT"].ToString();
                                obj.RTG_MDY_SEN_UNSEC_RTG_DT = dr["RTG_MDY_SEN_UNSEC_RTG_DT"].ToString();

                                obj.RTG_MDY_FC_CURR_ISSUER_RATING = dr["RTG_MDY_FC_CURR_ISSUER_RATING"].ToString();
                                obj.RTG_MDY_FC_CURR_ISSUER_RTG_DT = dr["RTG_MDY_FC_CURR_ISSUER_RTG_DT"].ToString();
                                obj.RTG_FITCH_LT_ISSUER_DEFAULT = dr["RTG_FITCH_LT_ISSUER_DEFAULT"].ToString();
                                obj.RTG_FITCH_LT_ISSUER_DFLT_RTG_DT = dr["RTG_FITCH_LT_ISSUER_DFLT_RTG_DT"].ToString();
                                obj.RTG_FITCH_SEN_UNSECURED = dr["RTG_FITCH_SEN_UNSECURED"].ToString();
                                obj.RTG_FITCH_SEN_UNSEC_RTG_DT = dr["RTG_FITCH_SEN_UNSEC_RTG_DT"].ToString();
                                obj.RTG_FITCH_LT_FC_ISSUER_DEFAULT = dr["RTG_FITCH_LT_FC_ISSUER_DEFAULT"].ToString();
                                obj.RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT = dr["RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT"].ToString();
                                obj.RTG_FITCH_LT_LC_ISSUER_DEFAULT = dr["RTG_FITCH_LT_LC_ISSUER_DEFAULT"].ToString();
                                obj.RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT = dr["RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT"].ToString();

                                obj.RTG_FITCH_NATIONAL_LT = dr["RTG_FITCH_NATIONAL_LT"].ToString();
                                obj.RTG_FITCH_NATIONAL_LT_DT = dr["RTG_FITCH_NATIONAL_LT_DT"].ToString();
                                obj.RTG_TRC_LONG_TERM = dr["RTG_TRC_LONG_TERM"].ToString();
                                obj.RTG_TRC_LONG_TERM_RTG_DT = dr["RTG_TRC_LONG_TERM_RTG_DT"].ToString();
                                obj.G_RTG_SP_LT_FC_ISSUER_CREDIT = dr["G_RTG_SP_LT_FC_ISSUER_CREDIT"].ToString();
                                obj.G_RTG_SP_LT_FC_ISS_CRED_RTG_DT = dr["G_RTG_SP_LT_FC_ISS_CRED_RTG_DT"].ToString();
                                obj.G_RTG_SP_LT_LC_ISSUER_CREDIT = dr["G_RTG_SP_LT_LC_ISSUER_CREDIT"].ToString();
                                obj.G_RTG_SP_LT_LC_ISS_CRED_RTG_DT = dr["G_RTG_SP_LT_LC_ISS_CRED_RTG_DT"].ToString();
                                obj.G_RTG_MDY_ISSUER = dr["G_RTG_MDY_ISSUER"].ToString();
                                obj.G_RTG_MDY_ISSUER_RTG_DT = dr["G_RTG_MDY_ISSUER_RTG_DT"].ToString();

                                obj.G_RTG_MOODY_LONG_TERM = dr["G_RTG_MOODY_LONG_TERM"].ToString();
                                obj.G_RTG_MOODY_LONG_TERM_DATE = dr["G_RTG_MOODY_LONG_TERM_DATE"].ToString();
                                obj.G_RTG_MDY_SEN_UNSECURED_DEBT = dr["G_RTG_MDY_SEN_UNSECURED_DEBT"].ToString();
                                obj.G_RTG_MDY_SEN_UNSEC_RTG_DT = dr["G_RTG_MDY_SEN_UNSEC_RTG_DT"].ToString();
                                obj.G_RTG_MDY_FC_CURR_ISSUER_RATING = dr["G_RTG_MDY_FC_CURR_ISSUER_RATING"].ToString();
                                obj.G_RTG_MDY_FC_CURR_ISSUER_RTG_DT = dr["G_RTG_MDY_FC_CURR_ISSUER_RTG_DT"].ToString();
                                obj.G_RTG_FITCH_LT_ISSUER_DEFAULT = dr["G_RTG_FITCH_LT_ISSUER_DEFAULT"].ToString();
                                obj.G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT = dr["G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT"].ToString();
                                obj.G_RTG_FITCH_SEN_UNSECURED = dr["G_RTG_FITCH_SEN_UNSECURED"].ToString();
                                obj.G_RTG_FITCH_SEN_UNSEC_RTG_DT = dr["G_RTG_FITCH_SEN_UNSEC_RTG_DT"].ToString();

                                obj.G_RTG_FITCH_LT_FC_ISSUER_DEFAULT = dr["G_RTG_FITCH_LT_FC_ISSUER_DEFAULT"].ToString();
                                obj.G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT = dr["G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT"].ToString();
                                obj.G_RTG_FITCH_LT_LC_ISSUER_DEFAULT = dr["G_RTG_FITCH_LT_LC_ISSUER_DEFAULT"].ToString();
                                obj.G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT = dr["G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT"].ToString();
                                obj.G_RTG_FITCH_NATIONAL_LT = dr["G_RTG_FITCH_NATIONAL_LT"].ToString();
                                obj.G_RTG_FITCH_NATIONAL_LT_DT = dr["G_RTG_FITCH_NATIONAL_LT_DT"].ToString();
                                obj.G_RTG_TRC_LONG_TERM = dr["G_RTG_TRC_LONG_TERM"].ToString();
                                obj.G_RTG_TRC_LONG_TERM_RTG_DT = dr["G_RTG_TRC_LONG_TERM_RTG_DT"].ToString();

                                //補入60個欄位

                                A59dataIFRS.Add(obj);
                            }
                            #endregion

                        }
                    }
                }

                #region Linq join兩個list
                var templist2 = (from left in A59Data
                                 join right in A59dataIFRS
                                  on new { left.Bond_Number, left.Origination_Date } equals new { right.Bond_Number, right.Origination_Date }
                                 into grouping
                                 from right2 in grouping.DefaultIfEmpty()
                                 select new A59ViewModel
                                 {
                                     Reference_Nbr = left.Reference_Nbr,
                                     Report_Date = left.Report_Date,
                                     Version = left.Version,
                                     Bond_Number = left.Bond_Number,
                                     Lots = left.Lots,
                                     Origination_Date = left.Origination_Date,
                                     Portfolio_Name = left.Portfolio_Name,
                                     Portfolio = left.Portfolio,
                                     SMF = left.SMF,
                                     Issuer = left.Issuer,
                                     Security_Ticker = left.Security_Ticker,
                                     RATING_AS_OF_DATE_OVERRIDE = left.RATING_AS_OF_DATE_OVERRIDE,
                                     Rating_Type = left.Rating_Type,
                                     ISIN_Changed_Ind = left.ISIN_Changed_Ind,
                                     Bond_Number_Old = left.Bond_Number_Old,

                                     RTG_SP = right2 == null || right2.RTG_SP == null ? "" : right2.RTG_SP,
                                     SP_EFF_DT = right2 == null || right2.SP_EFF_DT == null ? "" : right2.SP_EFF_DT,
                                     RTG_TRC = right2 == null || right2.RTG_TRC == null ? "" : right2.RTG_TRC,
                                     TRC_EFF_DT = right2 == null || right2.TRC_EFF_DT == null ? "" : right2.TRC_EFF_DT,
                                     RTG_MOODY = right2 == null || right2.RTG_MOODY == null ? "" : right2.RTG_MOODY,
                                     MOODY_EFF_DT = right2 == null || right2.MOODY_EFF_DT == null ? "" : right2.MOODY_EFF_DT,
                                     RTG_FITCH = right2 == null || right2.RTG_FITCH == null ? "" : right2.RTG_FITCH,
                                     FITCH_EFF_DT = right2 == null || right2.FITCH_EFF_DT == null ? "" : right2.FITCH_EFF_DT,
                                     RTG_FITCH_NATIONAL = right2 == null || right2.RTG_FITCH_NATIONAL == null ? "" : right2.RTG_FITCH_NATIONAL,
                                     RTG_FITCH_NATIONAL_DT = right2 == null || right2.RTG_FITCH_NATIONAL_DT == null ? "" : right2.RTG_FITCH_NATIONAL_DT,

                                     RTG_SP_LT_FC_ISSUER_CREDIT = right2 == null || right2.RTG_SP_LT_FC_ISSUER_CREDIT == null ? "" : right2.RTG_SP_LT_FC_ISSUER_CREDIT,
                                     RTG_SP_LT_FC_ISS_CRED_RTG_DT = right2 == null || right2.RTG_SP_LT_FC_ISS_CRED_RTG_DT == null ? "" : right2.RTG_SP_LT_FC_ISS_CRED_RTG_DT,
                                     RTG_SP_LT_LC_ISSUER_CREDIT = right2 == null || right2.RTG_SP_LT_LC_ISSUER_CREDIT == null ? "" : right2.RTG_SP_LT_LC_ISSUER_CREDIT,
                                     RTG_SP_LT_LC_ISS_CRED_RTG_DT = right2 == null || right2.RTG_SP_LT_LC_ISS_CRED_RTG_DT == null ? "" : right2.RTG_SP_LT_LC_ISS_CRED_RTG_DT,
                                     RTG_MDY_ISSUER = right2 == null || right2.RTG_MDY_ISSUER == null ? "" : right2.RTG_MDY_ISSUER,
                                     RTG_MDY_ISSUER_RTG_DT = right2 == null || right2.RTG_MDY_ISSUER_RTG_DT == null ? "" : right2.RTG_MDY_ISSUER_RTG_DT,
                                     RTG_MOODY_LONG_TERM = right2 == null || right2.RTG_MOODY_LONG_TERM == null ? "" : right2.RTG_MOODY_LONG_TERM,
                                     RTG_MOODY_LONG_TERM_DATE = right2 == null || right2.RTG_MOODY_LONG_TERM_DATE == null ? "" : right2.RTG_MOODY_LONG_TERM_DATE,
                                     RTG_MDY_SEN_UNSECURED_DEBT = right2 == null || right2.RTG_MDY_SEN_UNSECURED_DEBT == null ? "" : right2.RTG_MDY_SEN_UNSECURED_DEBT,
                                     RTG_MDY_SEN_UNSEC_RTG_DT = right2 == null || right2.RTG_MDY_SEN_UNSEC_RTG_DT == null ? "" : right2.RTG_MDY_SEN_UNSEC_RTG_DT,

                                     RTG_MDY_FC_CURR_ISSUER_RATING = right2 == null || right2.RTG_MDY_FC_CURR_ISSUER_RATING == null ? "" : right2.RTG_MDY_FC_CURR_ISSUER_RATING,
                                     RTG_MDY_FC_CURR_ISSUER_RTG_DT = right2 == null || right2.RTG_MDY_FC_CURR_ISSUER_RTG_DT == null ? "" : right2.RTG_MDY_FC_CURR_ISSUER_RTG_DT,
                                     RTG_FITCH_LT_ISSUER_DEFAULT = right2 == null || right2.RTG_FITCH_LT_ISSUER_DEFAULT == null ? "" : right2.RTG_FITCH_LT_ISSUER_DEFAULT,
                                     RTG_FITCH_LT_ISSUER_DFLT_RTG_DT = right2 == null || right2.RTG_FITCH_LT_ISSUER_DFLT_RTG_DT == null ? "" : right2.RTG_FITCH_LT_ISSUER_DFLT_RTG_DT,
                                     RTG_FITCH_SEN_UNSECURED = right2 == null || right2.RTG_FITCH_SEN_UNSECURED == null ? "" : right2.RTG_FITCH_SEN_UNSECURED,
                                     RTG_FITCH_SEN_UNSEC_RTG_DT = right2 == null || right2.RTG_FITCH_SEN_UNSEC_RTG_DT == null ? "" : right2.RTG_FITCH_SEN_UNSEC_RTG_DT,
                                     RTG_FITCH_LT_FC_ISSUER_DEFAULT = right2 == null || right2.RTG_FITCH_LT_FC_ISSUER_DEFAULT == null ? "" : right2.RTG_FITCH_LT_FC_ISSUER_DEFAULT,
                                     RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT = right2 == null || right2.RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT == null ? "" : right2.RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT,
                                     RTG_FITCH_LT_LC_ISSUER_DEFAULT = right2 == null || right2.RTG_FITCH_LT_LC_ISSUER_DEFAULT == null ? "" : right2.RTG_FITCH_LT_LC_ISSUER_DEFAULT,
                                     RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT = right2 == null || right2.RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT == null ? "" : right2.RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT,

                                     RTG_FITCH_NATIONAL_LT = right2 == null || right2.RTG_FITCH_NATIONAL_LT == null ? "" : right2.RTG_FITCH_NATIONAL_LT,
                                     RTG_FITCH_NATIONAL_LT_DT = right2 == null || right2.RTG_FITCH_NATIONAL_LT_DT == null ? "" : right2.RTG_FITCH_NATIONAL_LT_DT,
                                     RTG_TRC_LONG_TERM = right2 == null || right2.RTG_TRC_LONG_TERM == null ? "" : right2.RTG_TRC_LONG_TERM,
                                     RTG_TRC_LONG_TERM_RTG_DT = right2 == null || right2.RTG_TRC_LONG_TERM_RTG_DT == null ? "" : right2.RTG_TRC_LONG_TERM_RTG_DT,
                                     G_RTG_SP_LT_FC_ISSUER_CREDIT = right2 == null || right2.G_RTG_SP_LT_FC_ISSUER_CREDIT == null ? "" : right2.G_RTG_SP_LT_FC_ISSUER_CREDIT,
                                     G_RTG_SP_LT_FC_ISS_CRED_RTG_DT = right2 == null || right2.G_RTG_SP_LT_FC_ISS_CRED_RTG_DT == null ? "" : right2.G_RTG_SP_LT_FC_ISS_CRED_RTG_DT,
                                     G_RTG_SP_LT_LC_ISSUER_CREDIT = right2 == null || right2.G_RTG_SP_LT_LC_ISSUER_CREDIT == null ? "" : right2.G_RTG_SP_LT_LC_ISSUER_CREDIT,
                                     G_RTG_SP_LT_LC_ISS_CRED_RTG_DT = right2 == null || right2.G_RTG_SP_LT_LC_ISS_CRED_RTG_DT == null ? "" : right2.G_RTG_SP_LT_LC_ISS_CRED_RTG_DT,
                                     G_RTG_MDY_ISSUER = right2 == null || right2.G_RTG_MDY_ISSUER == null ? "" : right2.G_RTG_MDY_ISSUER,
                                     G_RTG_MDY_ISSUER_RTG_DT = right2 == null || right2.G_RTG_MDY_ISSUER_RTG_DT == null ? "" : right2.G_RTG_MDY_ISSUER_RTG_DT,

                                     G_RTG_MOODY_LONG_TERM = right2 == null || right2.G_RTG_MOODY_LONG_TERM == null ? "" : right2.G_RTG_MOODY_LONG_TERM,
                                     G_RTG_MOODY_LONG_TERM_DATE = right2 == null || right2.G_RTG_MOODY_LONG_TERM_DATE == null ? "" : right2.G_RTG_MOODY_LONG_TERM_DATE,
                                     G_RTG_MDY_SEN_UNSECURED_DEBT = right2 == null || right2.G_RTG_MDY_SEN_UNSECURED_DEBT == null ? "" : right2.G_RTG_MDY_SEN_UNSECURED_DEBT,
                                     G_RTG_MDY_SEN_UNSEC_RTG_DT = right2 == null || right2.G_RTG_MDY_SEN_UNSEC_RTG_DT == null ? "" : right2.G_RTG_MDY_SEN_UNSEC_RTG_DT,
                                     G_RTG_MDY_FC_CURR_ISSUER_RATING = right2 == null || right2.G_RTG_MDY_FC_CURR_ISSUER_RATING == null ? "" : right2.G_RTG_MDY_FC_CURR_ISSUER_RATING,
                                     G_RTG_MDY_FC_CURR_ISSUER_RTG_DT = right2 == null || right2.G_RTG_MDY_FC_CURR_ISSUER_RTG_DT == null ? "" : right2.G_RTG_MDY_FC_CURR_ISSUER_RTG_DT,
                                     G_RTG_FITCH_LT_ISSUER_DEFAULT = right2 == null || right2.G_RTG_FITCH_LT_ISSUER_DEFAULT == null ? "" : right2.G_RTG_FITCH_LT_ISSUER_DEFAULT,
                                     G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT = right2 == null || right2.G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT == null ? "" : right2.G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT,
                                     G_RTG_FITCH_SEN_UNSECURED = right2 == null || right2.G_RTG_FITCH_SEN_UNSECURED == null ? "" : right2.G_RTG_FITCH_SEN_UNSECURED,
                                     G_RTG_FITCH_SEN_UNSEC_RTG_DT = right2 == null || right2.G_RTG_FITCH_SEN_UNSEC_RTG_DT == null ? "" : right2.G_RTG_FITCH_SEN_UNSEC_RTG_DT,

                                     G_RTG_FITCH_LT_FC_ISSUER_DEFAULT = right2 == null || right2.G_RTG_FITCH_LT_FC_ISSUER_DEFAULT == null ? "" : right2.G_RTG_FITCH_LT_FC_ISSUER_DEFAULT,
                                     G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT = right2 == null || right2.G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT == null ? "" : right2.G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT,
                                     G_RTG_FITCH_LT_LC_ISSUER_DEFAULT = right2 == null || right2.G_RTG_FITCH_LT_LC_ISSUER_DEFAULT == null ? "" : right2.G_RTG_FITCH_LT_LC_ISSUER_DEFAULT,
                                     G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT = right2 == null || right2.G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT == null ? "" : right2.G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT,
                                     G_RTG_FITCH_NATIONAL_LT = right2 == null || right2.G_RTG_FITCH_NATIONAL_LT == null ? "" : right2.G_RTG_FITCH_NATIONAL_LT,
                                     G_RTG_FITCH_NATIONAL_LT_DT = right2 == null || right2.G_RTG_FITCH_NATIONAL_LT_DT == null ? "" : right2.G_RTG_FITCH_NATIONAL_LT_DT,
                                     G_RTG_TRC_LONG_TERM = right2 == null || right2.G_RTG_TRC_LONG_TERM == null ? "" : right2.G_RTG_TRC_LONG_TERM,
                                     G_RTG_TRC_LONG_TERM_RTG_DT = right2 == null || right2.G_RTG_TRC_LONG_TERM_RTG_DT == null ? "" : right2.G_RTG_TRC_LONG_TERM_RTG_DT

                                 }).ToList();
                #endregion join兩個list
                //==========
                #endregion 串投資平台DB資料

                #region 測試調整匯出格式
                bool flag = false;
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (templist2.Any())
                    {
                        var first = templist2.First();
                        DateTime reportDate = DateTime.Parse(first.Report_Date);
                        int version = Int32.Parse(first.Version);
                        string Rating_Type = first.Rating_Type;

                        var A57s = db.Bond_Rating_Info.AsNoTracking()
                            .Where(x => x.Report_Date == reportDate &&
                                        x.Version == version &&
                                        x.Rating_Type == Rating_Type &&
                                        //x.RTG_Bloomberg_Field != null &&
                                        x.Rating != null).ToList();
                        var A57ns = db.Bond_Rating_Info.AsNoTracking()
                            .Where(x => x.Report_Date == reportDate &&
                                        x.Version == version &&
                                        x.Rating_Type == Rating_Type &&
                                        //x.RTG_Bloomberg_Field == " " &&
                                        x.Rating == null).ToList();
                        var PInfos = new A59ViewModel().GetType().GetProperties();  //excel titles
                        templist2.ForEach(x =>
                        {
                            var A57Data = A57s.Where(y =>
                            y.Reference_Nbr == x.Reference_Nbr).ToList();
                            if (A57Data.Any())
                            {
                                var A57f = A57Data.First();
                                x.ISSUER_EQUITY_TICKER = A57f.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57f.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57f.GUARANTOR_EQY_TICKER;
                                x.Fill_up_YN = A57Data.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                                if (flag)
                                {
                                    A57Data.Aggregate(x,
                                    (A59, A57) =>
                                    {
                                        var PInfo = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                       A57.RTG_Bloomberg_Field.Trim().ToUpper())
                                                          .FirstOrDefault();
                                        if (PInfo != null)
                                        {
                                            PInfo.SetValue(A59, A57.Rating);
                                            var rdt = TypeTransfer.dateTimeNToString(A57.Rating_Date);
                                            if (!rdt.IsNullOrWhiteSpace())
                                            {
                                                var DtFildes = PInfo.GetCustomAttributes(typeof(SetDateFiledAttribute), false);
                                                if (DtFildes.Length > 0)
                                                {
                                                    var DtFilde = ((SetDateFiledAttribute)DtFildes[0]).Description;
                                                    var PInfo2 = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                                    DtFilde.ToUpper()).FirstOrDefault();
                                                    if (PInfo2 != null)
                                                    {
                                                        PInfo2.SetValue(A59, rdt);
                                                    }
                                                }
                                            }
                                        }
                                        return A59;
                                    });
                                }
                            }
                            else
                            {
                                var _A57ns = A57ns.Where(y =>
                                        y.Reference_Nbr == x.Reference_Nbr).ToList();
                                var A57n = _A57ns.FirstOrDefault();
                                x.ISSUER_EQUITY_TICKER = A57n?.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57n?.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57n?.GUARANTOR_EQY_TICKER;
                                x.Fill_up_YN = _A57ns.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                            }
                        });
                    }
                }
                #endregion
                if (templist2.Any())
                {
                    //檢查log檔有無資料
                    bool A59Flag = common.checkTransferCheck(Table_Type.A59Apex.ToString(), "A59Apex", rd, 0);
                    //
                    if (A59Flag)
                    {
                        using (IFRS9DBEntities db = new IFRS9DBEntities())
                        {
                            string sql = string.Empty;
                            sql += $@"UPDATE Transfer_CheckTable Set TransferType = 'R' 
                    WHERE ReportDate = '{reportdate.Replace("/", "-")}'
                    AND Version =0
                    AND TransferType = 'Y'
                    AND File_Name ='{Table_Type.A59Apex.ToString()}' ;";
                            db.Database.ExecuteSqlCommand(sql);
                        }
                    }
                    string Msg_GetA59 = $"{Message_Type.Join_CounterPartyCreditRating_Success.GetDescription()}" + $"共接入{dttempIFRS9.Rows.Count}筆資料";
                    common.saveTransferCheck(Table_Type.A59Apex.ToString(), true, rd, 0, startTime, DateTime.Now, Msg_GetA59);
                    return new Tuple<bool, List<A59ViewModel>>(true, templist2);
                }
                else
                {
                    common.saveTransferCheck(Table_Type.A59Apex.ToString(), true, rd, 0, startTime, DateTime.Now, Message_Type.not_Find_CounterPartyCreditRating.GetDescription());
                    return new Tuple<bool, List<A59ViewModel>>(false, new List<A59ViewModel>());
                }
            }
            catch (Exception ex)
            {
                common.saveTransferCheck(Table_Type.A59Apex.ToString(), false, rd, 0, startTime, DateTime.Now, Message_Type.Find_CounterPartyCreditRating_Error.GetDescription(null, ex.Message));
                return new Tuple<bool, List<A59ViewModel>>(false, new List<A59ViewModel>());
            }
        }

        #endregion 190222GetA59

        #endregion Get Data

        #region Save Db

        /// <summary>
        /// Save A56
        /// </summary>
        /// <param name="data">新增 or 修改(限定只能設定失效)</param>
        /// <param name="IsActive"></param>
        /// <returns></returns>
        public MSGReturnModel saveA56(A56ViewModel data, bool IsActive)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var tableName = Table_Type.A56.tableNameGetDescription();
            if (data == null)
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(tableName);
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (IsActive) //生效 (新增)
                {
                    var _Replace_Object = data.Replace_Object.TrimEnd();
                    if (_Replace_Object.IsNullOrWhiteSpace())
                    {
                        result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                        return result;
                    }
                    if (db.Rating_Update_Info.Any(x => x.Replace_Object == _Replace_Object && x.IsActive == "Y"))
                    {
                        result.DESCRIPTION = "已有相同字元存在,請重新輸入或將已生效的字元變更為無效。";
                        return result;
                    }
                    List<Rating_Update_Method> methods = EnumUtil.GetValues<Rating_Update_Method>().ToList();
                    db.Rating_Update_Info.Add(new Rating_Update_Info()
                    {
                        IsActive = "Y",
                        Replace_Object = _Replace_Object,
                        Update_Method = data.Update_Method,
                        Method_MEMO = A56Method_MEMO(data.Update_Method, methods),
                        Replace_Word = data.Replace_Word,
                        Create_User = _UserInfo._user,
                        Create_Date = _UserInfo._date,
                        Create_Time = _UserInfo._time,
                        LastUpdate_User = _UserInfo._user,
                        LastUpdate_Date = _UserInfo._date,
                        LastUpdate_Time = _UserInfo._time
                    });
                }
                else //失效 
                {
                    int id = 0;
                    Int32.TryParse(data.ID, out id);
                    var A56 = db.Rating_Update_Info.Where(x => x.IsActive == "Y").FirstOrDefault(y => y.ID == id);
                    if (A56 != null)
                    {
                        A56.IsActive = "N";
                        A56.LastUpdate_User = _UserInfo._user;
                        A56.LastUpdate_Date = _UserInfo._date;
                        A56.LastUpdate_Time = _UserInfo._time;
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription(tableName);
                        return result;
                    }
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    if (IsActive)
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription(tableName);
                    else
                        result.DESCRIPTION = Message_Type.update_Success.GetDescription(tableName);
                }
                catch (Exception ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }

        /// <summary>
        /// save A59 (save A57 & A58)
        /// </summary>
        /// <param name="dataModel">A59ViewModel</param>
        /// <returns></returns>
        public MSGReturnModel saveA59(List<A59ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error
                    .GetDescription(Table_Type.A59.ToString());
                return result;
            }

            DateTime startTime = DateTime.Now;
            string startDt = startTime.ToString("yyyy/MM/dd");
            List<StringBuilder> sbs = new List<StringBuilder>();

            List<Bond_Rating_Info> A57Issues = new List<Bond_Rating_Info>(); //A57檢核到有特殊的Rating 

            string SP = RatingOrg.SP.GetDescription(); //SP
            string Fitch = RatingOrg.Fitch.GetDescription(); //Fitch
            string FitchTwn = RatingOrg.FitchTwn.GetDescription(); //Fitch(twn)
            string Moody = RatingOrg.Moody.GetDescription(); //Moody
            string CW = RatingOrg.CW.GetDescription(); //CW

            string Bonds = RatingObject.Bonds.GetDescription(); //債項
            string ISSUER = RatingObject.ISSUER.GetDescription(); //發行人
            string GUARANTOR = RatingObject.GUARANTOR.GetDescription(); //保證人

            string Domestic = "國內";
            string Foreign = "國外";

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                #region 新增或修改需使用的資料

                List<Bond_Rating_Parm> D60s = getParmIDs();
                if (!D60s.Any())
                {
                    result.DESCRIPTION = Table_Type.D60.tableNameGetDescription() + "參數設定有誤";
                    return result;
                }
                var _checkParm = checkParmID(D60s);
                if (!_checkParm.IsNullOrWhiteSpace())
                {
                    result.DESCRIPTION = _checkParm;
                    return result;
                }
                List<Bond_Rating_Info> A57s = new List<Bond_Rating_Info>();
                List<Bond_Rating_Summary> A58s = new List<Bond_Rating_Summary>();
                List<Grade_Moody_Info> A51s = new List<Grade_Moody_Info>();
                List<Grade_Mapping_Info> A52s = new List<Grade_Mapping_Info>();
                List<Rating_Update_Info> A56s = db.Rating_Update_Info.AsNoTracking()
                    .Where(x => x.IsActive == "Y").ToList();
                rule_1s = A56s.Where(x => x.Update_Method == "01").Select(x => x.Replace_Object).ToList();
                rule_2s = A56s.Where(x => x.Update_Method == "02")
                    .Select(x => new Tuple<string, string>(x.Replace_Object, x.Replace_Word)).ToList();
                var first = dataModel.First();
                List<DateTime> reportDates = dataModel.Where(x => !x.Report_Date.IsNullOrWhiteSpace())
                                                      .Select(x => x.Report_Date).Distinct()
                                                      .Select(x => TypeTransfer.stringToDateTime(x)).ToList();
                List<int?> versions = dataModel.Where(x => !x.Version.IsNullOrWhiteSpace())
                                              .Select(x => x.Version).Distinct()
                                              .Select(x => TypeTransfer.stringToIntN(x)).ToList();
                List<string> ratingTypes = dataModel.Where(x => !x.Rating_Type.IsNullOrWhiteSpace())
                                                    .Select(x => x.Rating_Type).Distinct().ToList();
                A57s = db.Bond_Rating_Info.AsNoTracking()
                         .Where(x => reportDates.Contains(x.Report_Date) &&
                                     versions.Contains(x.Version) &&
                                     ratingTypes.Contains(x.Rating_Type)).ToList();
                A58s = db.Bond_Rating_Summary.AsNoTracking()
                         .Where(x => reportDates.Contains(x.Report_Date) &&
                                     versions.Contains(x.Version) &&
                                     ratingTypes.Contains(x.Rating_Type)).ToList();
                A51s = db.Grade_Moody_Info.AsNoTracking()
                         .Where(x => x.Status == "1").ToList();
                A52s = db.Grade_Mapping_Info.AsNoTracking().Where(x=>x.IsActive == "Y").ToList();

                #endregion 新增或修改需使用的資料

                #region 新增 & 修改 暫存資料表

                List<Bond_Rating_Info> InsertA57s = new List<Bond_Rating_Info>();
                List<Bond_Rating_Info> UpdateA57s = new List<Bond_Rating_Info>();
                List<Bond_Rating_Summary> InsertA58s = new List<Bond_Rating_Summary>();
                List<Bond_Rating_Summary> UpdateA58s = new List<Bond_Rating_Summary>();

                #endregion 新增 & 修改 暫存資料表

                dataModel.ForEach(x =>
                {
                    List<A59Data> datas = new List<A59Data>();
                    string issureStr = x.ISSUER_EQUITY_TICKER;
                    string gNameStr = x.GUARANTOR_NAME;
                    string gEqyTickerStr = x.GUARANTOR_EQY_TICKER;

                    #region Excel 資料組成 A59資料

                    addData("RTG_SP", x.RTG_SP, x.SP_EFF_DT, Bonds, Foreign, SP, datas);
                    addData("RTG_TRC", x.RTG_TRC, x.TRC_EFF_DT, Bonds, Domestic, CW, datas);
                    addData("RTG_MOODY", x.RTG_MOODY, x.MOODY_EFF_DT, Bonds, Foreign, Moody, datas);
                    addData("RTG_FITCH", x.RTG_FITCH, x.FITCH_EFF_DT, Bonds, Foreign, Fitch, datas);
                    addData("RTG_FITCH_NATIONAL", x.RTG_FITCH_NATIONAL, x.RTG_FITCH_NATIONAL_DT, Bonds, Domestic, FitchTwn, datas);
                    addData("RTG_SP_LT_FC_ISSUER_CREDIT", x.RTG_SP_LT_FC_ISSUER_CREDIT, x.RTG_SP_LT_FC_ISS_CRED_RTG_DT, ISSUER, Foreign, SP, datas);
                    addData("RTG_SP_LT_LC_ISSUER_CREDIT", x.RTG_SP_LT_LC_ISSUER_CREDIT, x.RTG_SP_LT_LC_ISS_CRED_RTG_DT, ISSUER, Foreign, SP, datas);
                    addData("RTG_MDY_ISSUER", x.RTG_MDY_ISSUER, x.RTG_MDY_ISSUER_RTG_DT, ISSUER, Foreign, Moody, datas);
                    addData("RTG_MOODY_LONG_TERM", x.RTG_MOODY_LONG_TERM, x.RTG_MOODY_LONG_TERM_DATE, ISSUER, Foreign, Moody, datas);
                    addData("RTG_MDY_SEN_UNSECURED_DEBT", x.RTG_MDY_SEN_UNSECURED_DEBT, x.RTG_MDY_SEN_UNSEC_RTG_DT, ISSUER, Foreign, Moody, datas);
                    addData("RTG_MDY_FC_CURR_ISSUER_RATING", x.RTG_MDY_FC_CURR_ISSUER_RATING, x.RTG_MDY_FC_CURR_ISSUER_RTG_DT, ISSUER, Foreign, Moody, datas);
                    //addData("RTG_MDY_LOCAL_LT_BANK_DEPOSITS", x.RTG_MDY_LOCAL_LT_BANK_DEPOSITS, x.RTG_MDY_LT_LC_BANK_DEP_RTG_DT, ISSUER, Domestic, SP, datas);
                    addData("RTG_FITCH_LT_ISSUER_DEFAULT", x.RTG_FITCH_LT_ISSUER_DEFAULT, x.RTG_FITCH_LT_ISSUER_DFLT_RTG_DT, ISSUER, Foreign, Fitch, datas);
                    addData("RTG_FITCH_SEN_UNSECURED", x.RTG_FITCH_SEN_UNSECURED, x.RTG_FITCH_SEN_UNSEC_RTG_DT, ISSUER, Foreign, Fitch, datas);
                    addData("RTG_FITCH_LT_FC_ISSUER_DEFAULT", x.RTG_FITCH_LT_FC_ISSUER_DEFAULT, x.RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT, ISSUER, Foreign, Fitch, datas);
                    addData("RTG_FITCH_LT_LC_ISSUER_DEFAULT", x.RTG_FITCH_LT_LC_ISSUER_DEFAULT, x.RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT, ISSUER, Foreign, Fitch, datas);
                    addData("RTG_FITCH_NATIONAL_LT", x.RTG_FITCH_NATIONAL_LT, x.RTG_FITCH_NATIONAL_LT_DT, ISSUER, Domestic, FitchTwn, datas);
                    addData("RTG_TRC_LONG_TERM", x.RTG_TRC_LONG_TERM, x.RTG_TRC_LONG_TERM_RTG_DT, ISSUER, Domestic, CW, datas);
                    addData("G_RTG_SP_LT_FC_ISSUER_CREDIT", x.G_RTG_SP_LT_FC_ISSUER_CREDIT, x.G_RTG_SP_LT_FC_ISS_CRED_RTG_DT, GUARANTOR, Foreign, SP, datas);
                    addData("G_RTG_SP_LT_LC_ISSUER_CREDIT", x.G_RTG_SP_LT_LC_ISSUER_CREDIT, x.G_RTG_SP_LT_LC_ISS_CRED_RTG_DT, GUARANTOR, Foreign, SP, datas);
                    addData("G_RTG_MDY_ISSUER", x.G_RTG_MDY_ISSUER, x.G_RTG_MDY_ISSUER_RTG_DT, GUARANTOR, Foreign, Moody, datas);
                    addData("G_RTG_MOODY_LONG_TERM", x.G_RTG_MOODY_LONG_TERM, x.G_RTG_MOODY_LONG_TERM_DATE, GUARANTOR, Foreign, Moody, datas);
                    addData("G_RTG_MDY_SEN_UNSECURED_DEBT", x.G_RTG_MDY_SEN_UNSECURED_DEBT, x.G_RTG_MDY_SEN_UNSEC_RTG_DT, GUARANTOR, Foreign, Moody, datas);
                    addData("G_RTG_MDY_FC_CURR_ISSUER_RATING", x.G_RTG_MDY_FC_CURR_ISSUER_RATING, x.G_RTG_MDY_FC_CURR_ISSUER_RTG_DT, GUARANTOR, Foreign, Moody, datas);
                    //addData("G_RTG_MDY_LOCAL_LT_BANK_DEPOSITS", x.G_RTG_MDY_LOCAL_LT_BANK_DEPOSITS, x.G_RTG_MDY_LT_LC_BANK_DEP_RTG_DT, GUARANTOR, Domestic, Moody, datas);
                    addData("G_RTG_FITCH_LT_ISSUER_DEFAULT", x.G_RTG_FITCH_LT_ISSUER_DEFAULT, x.G_RTG_FITCH_LT_ISSUER_DFLT_RTG_DT, GUARANTOR, Foreign, Fitch, datas);
                    addData("G_RTG_FITCH_SEN_UNSECURED", x.G_RTG_FITCH_SEN_UNSECURED, x.G_RTG_FITCH_SEN_UNSEC_RTG_DT, GUARANTOR, Foreign, Fitch, datas);
                    addData("G_RTG_FITCH_LT_FC_ISSUER_DEFAULT", x.G_RTG_FITCH_LT_FC_ISSUER_DEFAULT, x.G_RTG_FITCH_LT_FC_ISS_DFLT_RTG_DT, GUARANTOR, Foreign, Fitch, datas);
                    addData("G_RTG_FITCH_LT_LC_ISSUER_DEFAULT", x.G_RTG_FITCH_LT_LC_ISSUER_DEFAULT, x.G_RTG_FITCH_LT_LC_ISS_DFLT_RTG_DT, GUARANTOR, Foreign, Fitch, datas);
                    addData("G_RTG_FITCH_NATIONAL_LT", x.G_RTG_FITCH_NATIONAL_LT, x.G_RTG_FITCH_NATIONAL_LT_DT, GUARANTOR, Domestic, FitchTwn, datas);
                    addData("G_RTG_TRC_LONG_TERM", x.G_RTG_TRC_LONG_TERM, x.G_RTG_TRC_LONG_TERM_RTG_DT, GUARANTOR, Domestic, CW, datas);

                    #endregion Excel 資料組成 A59資料

                    if (datas.Any())
                    {
                        DateTime reportDate = TypeTransfer.stringToDateTime(x.Report_Date);
                        int? version = TypeTransfer.stringToIntN(x.Version);
                        var ratingType = x.Rating_Type;
                        //var year = reportDate.Year;
                        var originationDate = TypeTransfer.stringToDateTimeN(x.Origination_Date);

                        List<Bond_Rating_Info> InsertA57 = new List<Bond_Rating_Info>();
                        List<Bond_Rating_Info> UpdateA57 = new List<Bond_Rating_Info>();

                        //現有的A57資料
                        var A57s2 = A57s.Where(y => y.Reference_Nbr == x.Reference_Nbr &&
                                                            y.Report_Date == reportDate &&
                                                            y.Version == version &&
                                                            y.Bond_Number == x.Bond_Number &&
                                                            y.Rating_Type == x.Rating_Type &&
                                                            y.Lots == x.Lots &&
                                                            y.Portfolio_Name == x.Portfolio_Name).ToList();
                        //現有的A58資料
                        var A58s2 = A58s.Where(z => z.Reference_Nbr == x.Reference_Nbr &&
                                                            z.Report_Date == reportDate &&
                                                            z.Bond_Number == x.Bond_Number &&
                                                            z.Rating_Type == x.Rating_Type &&
                                                            z.Version == version &&
                                                            z.Lots == x.Lots &&
                                                            z.Portfolio_Name == x.Portfolio_Name).ToList();

                        if (A57s2.Any() && A58s2.Any())
                        {
                            #region Save A57

                            foreach (var data in datas)
                            {
                                var RTG_Bloomberg_Field = data.RTG_Bloomberg_Field;
                                var Rating_Date = TypeTransfer.stringToDateTimeN(data.Rating_Date);
                                var D60 = getParmID(D60s, data.Rating_Object, data.Rating_Org_Area);

                                var A57 = A57s2.FirstOrDefault(y => y.RTG_Bloomberg_Field == data.RTG_Bloomberg_Field);
                                Grade_Moody_Info A51 = null;
                                if (data.Rating_Org == RatingOrg.Moody.GetDescription())
                                {
                                    A51 = A51s.FirstOrDefault(z => z.Rating == data.Rating);
                                }
                                else
                                {
                                    var A52 = A52s.FirstOrDefault(z => z.Rating_Org == data.Rating_Org &&
                                                                       z.Rating == data.Rating);
                                    if (A52 != null)
                                        A51 = A51s.FirstOrDefault(z => z.PD_Grade == A52.PD_Grade);
                                }

                                //Rating 有值 但是 A51 比對不到 就為錯誤
                                if (!data.Rating.IsNullOrWhiteSpace() && A51 == null)
                                {
                                    A57Issues.Add(new Bond_Rating_Info()
                                    {
                                        Report_Date = reportDate,
                                        Version = version,
                                        Bond_Number = x.Bond_Number,
                                        Rating_Type = x.Rating_Type,
                                        RTG_Bloomberg_Field = data.RTG_Bloomberg_Field,
                                        Rating = data.Rating
                                    });
                                }

                                if (!A57Issues.Any()) //參數有誤不新增
                                {
                                    //A57 有找到資料 Rating 有參數
                                    if (A57 != null && data.Rating != string.Empty)
                                    {
                                        //if (A57.Rating != data.Rating)
                                        //{
                                            A57.Rating = data.Rating;
                                            if (A51 != null)
                                            {
                                                A57.PD_Grade = A51.PD_Grade;
                                                A57.Grade_Adjust = A51.Grade_Adjust;
                                            }
                                            A57.Rating_Date = Rating_Date;
                                            if (!issureStr.IsNullOrWhiteSpace())
                                                A57.ISSUER_TICKER = issureStr;
                                            if (!gNameStr.IsNullOrWhiteSpace())
                                                A57.GUARANTOR_NAME = gNameStr;
                                            if (!gEqyTickerStr.IsNullOrWhiteSpace())
                                                A57.GUARANTOR_EQY_TICKER = gEqyTickerStr;
                                            A57.Fill_up_Date = startTime;
                                            A57.Fill_up_YN = "Y";
                                            UpdateA57.Add(A57);
                                        //}
                                    }
                                    //A57 有找到資料 Rating 沒有參數
                                    else if (A57 != null && data.Rating == string.Empty)
                                    {
                                        //找到的A57 Rating 有值要設為null
                                        if (A57.Rating != null)
                                        {
                                            A57.Rating = null;
                                            A57.PD_Grade = null;
                                            A57.Grade_Adjust = null;
                                            A57.Rating_Date = null;
                                            if (!issureStr.IsNullOrWhiteSpace())
                                                A57.ISSUER_TICKER = issureStr;
                                            if (!gNameStr.IsNullOrWhiteSpace())
                                                A57.GUARANTOR_NAME = gNameStr;
                                            if (!gEqyTickerStr.IsNullOrWhiteSpace())
                                                A57.GUARANTOR_EQY_TICKER = gEqyTickerStr;
                                            A57.Fill_up_Date = startTime;
                                            A57.Fill_up_YN = "Y";
                                            UpdateA57.Add(A57);
                                        }
                                    }
                                    //Rating 有參數
                                    else if (data.Rating != string.Empty) //有rating 才須新增
                                    {
                                        if (A51 == null)
                                            A51 = new Grade_Moody_Info();
                                        var copyA57 = A57s2.First();
                                        InsertA57.Add(new Bond_Rating_Info()
                                        {
                                            Reference_Nbr = x.Reference_Nbr,
                                            Bond_Number = x.Bond_Number,
                                            Lots = x.Lots,
                                            Rating_Type = x.Rating_Type,
                                            Rating_Object = data.Rating_Object,
                                            Rating_Org = data.Rating_Org,
                                            Rating = data.Rating,
                                            Rating_Date = Rating_Date,
                                            Rating_Org_Area = data.Rating_Org_Area,
                                            Fill_up_YN = "Y",
                                            Fill_up_Date = startTime,
                                            Grade_Adjust = A51.Grade_Adjust,
                                            PD_Grade = A51.PD_Grade,
                                            Portfolio_Name = x.Portfolio_Name,
                                            RTG_Bloomberg_Field = data.RTG_Bloomberg_Field,
                                            SMF = x.SMF,
                                            ISSUER = x.Issuer,
                                            Version = version,
                                            Security_Ticker = x.Security_Ticker,
                                            Report_Date = reportDate,
                                            Origination_Date = copyA57.Origination_Date,
                                            ISSUER_TICKER = issureStr,
                                            GUARANTOR_NAME = gNameStr,
                                            GUARANTOR_EQY_TICKER = gEqyTickerStr,
                                            Portfolio = copyA57.Portfolio,
                                            Segment_Name = copyA57.Segment_Name,
                                            Parm_ID = D60 != null ? D60.Parm_ID.ToString() : null,
                                            Bond_Type = copyA57.Bond_Type,
                                            Lien_position = copyA57.Lien_position,
                                            ISIN_Changed_Ind = copyA57.ISIN_Changed_Ind,
                                            Bond_Number_Old = copyA57.Bond_Number_Old,
                                            Lots_Old = copyA57.Lots_Old,
                                            Portfolio_Name_Old = copyA57.Portfolio_Name_Old,
                                            Origination_Date_Old = copyA57.Origination_Date_Old
                                        });
                                    }
                                }

                            }

                            InsertA57s.AddRange(InsertA57);
                            UpdateA57s.AddRange(UpdateA57);

                            #endregion Save A57

                            #region Save A58 
                            if (!A57Issues.Any()) //參數有誤不新增
                            {
                                foreach (var A58Data in datas.GroupBy(y =>
                                new { y.Rating_Object, y.Rating_Org_Area }
                                ))
                                {
                                    var Rating_Object = A58Data.Key.Rating_Object;
                                    var Rating_Org_Area = A58Data.Key.Rating_Org_Area;
                                    var A57 = A57s2.Where(z => z.Rating_Object == Rating_Object &&
                                                               z.Rating_Org_Area == Rating_Org_Area).ToList();
                                    var D60 = getParmID(D60s, Rating_Object, Rating_Org_Area);
                                    var newA57 = new List<Bond_Rating_Info>();
                                    A57.ForEach(z =>
                                    {
                                        //假設舊資料沒有被 update 過
                                        if (UpdateA57.FirstOrDefault(i => i.RTG_Bloomberg_Field == z.RTG_Bloomberg_Field) == null)
                                            newA57.Add(z);
                                    });
                                    newA57.AddRange(InsertA57);
                                    newA57.AddRange(UpdateA57);
                                    newA57 = newA57.Where(z => z.Rating_Object == Rating_Object &&
                                                               z.Rating_Org_Area == Rating_Org_Area &&
                                                               z.Grade_Adjust != null).ToList();
                                    var A58 = A58s2.FirstOrDefault(z => z.Rating_Object == Rating_Object &&
                                                                        z.Rating_Org_Area == Rating_Org_Area);
                                    if (A58 != null)
                                    {
                                        A58.Rating_Selection = D60.Rating_Selection;
                                        A58.Processing_Date = startTime;
                                        if (newA57.Any())
                                        {
                                            if (A58.Rating_Selection == "1")
                                                A58.Grade_Adjust = newA57.Min(z => z.Grade_Adjust);
                                            if (A58.Rating_Selection == "2")
                                                A58.Grade_Adjust = newA57.Max(z => z.Grade_Adjust);
                                            UpdateA58s.Add(A58);
                                        }
                                        else
                                        {
                                            A58.Grade_Adjust = null;
                                            UpdateA58s.Add(A58);
                                        }
                                    }
                                    else if (newA57.Any())
                                    {
                                        var copyA58 = A58s2.First();
                                        int? Grade_Adjust = newA57.Any() & D60 != null ?
                                                            D60.Rating_Selection == "1" ?
                                                            newA57.Min(z => z.Grade_Adjust) :
                                                            D60.Rating_Selection == "2" ?
                                                            newA57.Max(z => z.Grade_Adjust) :
                                                            null : null;
                                        InsertA58s.Add(new Bond_Rating_Summary()
                                        {
                                            Reference_Nbr = x.Reference_Nbr,
                                            Report_Date = reportDate,
                                            Origination_Date = copyA58.Origination_Date,
                                            Lots = x.Lots,
                                            Rating_Type = x.Rating_Type,
                                            Rating_Object = Rating_Object,
                                            Rating_Org_Area = Rating_Org_Area,
                                            Grade_Adjust = Grade_Adjust,
                                            Rating_Priority = D60?.Rating_Priority,
                                            Rating_Selection = D60?.Rating_Selection,
                                            Processing_Date = startTime,
                                            Version = version,
                                            Bond_Number = x.Bond_Number,
                                            Portfolio_Name = x.Portfolio_Name,
                                            SMF = x.SMF,
                                            ISSUER = x.Issuer,
                                            Parm_ID = copyA58.Parm_ID,
                                            Portfolio = copyA58.Portfolio,
                                            Bond_Type = copyA58.Bond_Type,
                                            ISIN_Changed_Ind = copyA58.ISIN_Changed_Ind,
                                            Bond_Number_Old = copyA58.Bond_Number_Old,
                                            Lots_Old = copyA58.Lots_Old,
                                            Portfolio_Name_Old = copyA58.Portfolio_Name_Old,
                                            Origination_Date_Old = copyA58.Origination_Date_Old
                                        });
                                    }
                                }
                            }
                            #endregion Save A58
                        }
                    }
                });

                #region 新增 & 修改 A57,A58

                if (A57Issues.Any())
                {
                    var _first = A57Issues.First();
                    new BondsCheckRepository<Bond_Rating_Info>(A57Issues, Check_Table_Type.Bonds_A59_Before_Check, _first.Report_Date, _first.Version);
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Check_Fail.GetDescription(null, "「轉檔內容有信評特殊值，資料未寫入，請修改後重新上傳。」"); //檢核錯誤
                    return result;
                }

                foreach (var item in InsertA57s)
                {
                    sbs.Add(new StringBuilder($@"
INSERT INTO [Bond_Rating_Info]
           ([Reference_Nbr]
           ,[Bond_Number]
           ,[Lots]
           ,[Portfolio]
           ,[Segment_Name]
           ,[Bond_Type]
           ,[Lien_position]
           ,[Origination_Date]
           ,[Report_Date]
           ,[Rating_Date]
           ,[Rating_Type]
           ,[Rating_Object]
           ,[Rating_Org]
           ,[Rating]
           ,[Rating_Org_Area]
           ,[Fill_up_YN]
           ,[Fill_up_Date]
           ,[PD_Grade]
           ,[Grade_Adjust]
           ,[ISSUER_TICKER]
           ,[GUARANTOR_NAME]
           ,[GUARANTOR_EQY_TICKER]
           ,[Parm_ID]
           ,[Portfolio_Name]
           ,[RTG_Bloomberg_Field]
           ,[SMF]
           ,[ISSUER]
           ,[Version]
           ,[Security_Ticker]
           ,[ISIN_Changed_Ind]
           ,[Bond_Number_Old]
           ,[Lots_Old]
           ,[Portfolio_Name_Old]
           ,[Origination_Date_Old]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time])
     VALUES
           ({item.Reference_Nbr.stringToStrSql()}
           ,{item.Bond_Number.stringToStrSql()}
           ,{item.Lots.stringToStrSql()}
           ,{item.Portfolio.stringToStrSql()}
           ,{item.Segment_Name.stringToStrSql()}
           ,{item.Bond_Type.stringToStrSql()}
           ,{item.Lien_position.stringToStrSql()}
           ,{item.Origination_Date.dateTimeNToStrSql()}
           ,{item.Report_Date.dateTimeToStrSql()}
           ,{item.Rating_Date.dateTimeNToStrSql()}
           ,{item.Rating_Type.stringToStrSql()}
           ,{item.Rating_Object.stringToStrSql()}
           ,{item.Rating_Org.stringToStrSql()}
           ,{item.Rating.stringToStrSql()}
           ,{item.Rating_Org_Area.stringToStrSql()}
           ,{item.Fill_up_YN.stringToStrSql()}
           ,{item.Fill_up_Date.dateTimeNToStrSql()}
           ,{item.PD_Grade.intNToStrSql()}
           ,{item.Grade_Adjust.intNToStrSql()}
           ,{item.ISSUER_TICKER.stringToStrSql()}
           ,{item.GUARANTOR_NAME.stringToStrSql()}
           ,{item.GUARANTOR_EQY_TICKER.stringToStrSql()}
           ,{item.Parm_ID.stringToStrSql()}
           ,{item.Portfolio_Name.stringToStrSql()}
           ,{item.RTG_Bloomberg_Field.stringToStrSql()}
           ,{item.SMF.stringToStrSql()}
           ,{item.ISSUER.stringToStrSql()}
           ,{item.Version.intNToStrSql()}
           ,{item.Security_Ticker.stringToStrSql()}
           ,{item.ISIN_Changed_Ind.stringToStrSql()}
           ,{item.Bond_Number_Old.stringToStrSql()}
           ,{item.Lots_Old.stringToStrSql()}
           ,{item.Portfolio_Name_Old.stringToStrSql()}
           ,{item.Origination_Date_Old.dateTimeNToStrSql()} 
           ,{_UserInfo._user.stringToStrSql()}
           ,{_UserInfo._date.dateTimeToStrSql()}
           ,{_UserInfo._time.timeSpanToStrSql()} );
                    "));
                }
                foreach (var item in UpdateA57s)
                {
                    string sql =
                        $@"
UPDATE [Bond_Rating_Info]
   SET [LastUpdate_User] = {_UserInfo._user.stringToStrSql()}
      ,[LastUpdate_Date] = {_UserInfo._date.dateTimeToStrSql()}
      ,[LastUpdate_Time] = {_UserInfo._time.timeSpanToStrSql()}
      ,[Rating] = {item.Rating.stringToStrSql()}
      ,[Fill_up_YN] = {item.Fill_up_YN.stringToStrSql()}
      ,[Fill_up_Date] = {item.Fill_up_Date.dateTimeNToStrSql()}
      ,[PD_Grade] = {item.PD_Grade.intNToStrSql()}
      ,[Grade_Adjust] = {item.Grade_Adjust.intNToStrSql()} ";
                    sql += $@" ,[Rating_Date] = {item.Rating_Date.dateTimeNToStrSql()} ";
                    if (!item.ISSUER_TICKER.IsNullOrWhiteSpace())
                        sql += $@" ,[ISSUER_TICKER] = {item.ISSUER_TICKER.stringToStrSql()} ";
                    if (!item.GUARANTOR_NAME.IsNullOrWhiteSpace())
                        sql += $@" ,[GUARANTOR_NAME] = {item.GUARANTOR_NAME.stringToStrSql()} ";
                    if (!item.GUARANTOR_EQY_TICKER.IsNullOrWhiteSpace())
                        sql += $@" ,[GUARANTOR_EQY_TICKER] = {item.GUARANTOR_EQY_TICKER.stringToStrSql()} ";
                    sql += $@"
WHERE Reference_Nbr = { item.Reference_Nbr.stringToStrSql() }
  AND Report_Date = { item.Report_Date.dateTimeToStrSql() }
  AND Rating_Type = { item.Rating_Type.stringToStrSql() }
  AND Bond_Number = { item.Bond_Number.stringToStrSql() }
  AND Lots = { item.Lots.stringToStrSql() }
  AND RTG_Bloomberg_Field = { item.RTG_Bloomberg_Field.stringToStrSql() }
  AND Portfolio_Name = { item.Portfolio_Name.stringToStrSql() }
  AND Version = { item.Version.intNToStrSql() };
                    ";
                    sbs.Add(new StringBuilder(sql));
                }
                foreach (var item in InsertA57s.GroupBy(z =>
                new { z.Reference_Nbr, z.Bond_Number, z.Lots, z.Rating_Type, z.Report_Date, z.Version, z.Portfolio_Name }))
                {
                    sbs.Add(new StringBuilder(
                        $@"
Delete Bond_Rating_Info
where Reference_Nbr = {item.Key.Reference_Nbr.stringToStrSql()}
and Bond_Number = {item.Key.Bond_Number.stringToStrSql()}
and Lots = {item.Key.Lots.stringToStrSql()}
and Rating_Type = {item.Key.Rating_Type.stringToStrSql()}
and Report_Date = {item.Key.Report_Date.dateTimeToStrSql()}
and Version = {item.Key.Version.intNToStrSql()}
and Portfolio_Name = {item.Key.Portfolio_Name.stringToStrSql()}
and Rating_Object = ' '
and Rating_Org is null
and Rating is null
and Rating_Org_Area = ' ' ;
"
                        ));
                }
                foreach (var item in InsertA58s)
                {
                    sbs.Add(new StringBuilder(
                        $@"
INSERT INTO [Bond_Rating_Summary]
           ([Reference_Nbr]
           ,[Report_Date]
           ,[Parm_ID]
           ,[Bond_Type]
           ,[Rating_Type]
           ,[Rating_Object]
           ,[Rating_Org_Area]
           ,[Rating_Selection]
           ,[Grade_Adjust]
           ,[Rating_Priority]
           ,[Processing_Date]
           ,[Version]
           ,[Bond_Number]
           ,[Lots]
           ,[Portfolio]
           ,[Origination_Date]
           ,[Portfolio_Name]
           ,[SMF]
           ,[ISSUER]
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
           ({item.Reference_Nbr.stringToStrSql()}
           ,{item.Report_Date.dateTimeToStrSql()}
           ,{item.Parm_ID.stringToStrSql()}
           ,{item.Bond_Type.stringToStrSql()}
           ,{item.Rating_Type.stringToStrSql()}
           ,{item.Rating_Object.stringToStrSql()}
           ,{item.Rating_Org_Area.stringToStrSql()}
           ,{item.Rating_Selection.stringToStrSql()}
           ,{item.Grade_Adjust.intNToStrSql()}
           ,{item.Rating_Priority.intNToStrSql()}
           ,{item.Processing_Date.dateTimeNToStrSql()}
           ,{item.Version.intNToStrSql()}
           ,{item.Bond_Number.stringToStrSql()}
           ,{item.Lots.stringToStrSql()}
           ,{item.Portfolio.stringToStrSql()}
           ,{item.Origination_Date.dateTimeNToStrSql()}
           ,{item.Portfolio_Name.stringToStrSql()}
           ,{item.SMF.stringToStrSql()}
           ,{item.ISSUER.stringToStrSql()}
           ,{item.ISIN_Changed_Ind.stringToStrSql()}
           ,{item.Bond_Number_Old.stringToStrSql()}
           ,{item.Lots_Old.stringToStrSql()}
           ,{item.Portfolio_Name_Old.stringToStrSql()}
           ,{item.Origination_Date_Old.dateTimeNToStrSql()}
           ,{_UserInfo._user.stringToStrSql()}
           ,{_UserInfo._date.dateTimeToStrSql()}
           ,{_UserInfo._time.timeSpanToStrSql()} );
"
                        ));
                }
                foreach (var item in UpdateA58s)
                {
                    string sql = $@"
UPDATE [Bond_Rating_Summary]
   SET [LastUpdate_User] = {_UserInfo._user.stringToStrSql()} ,
       [LastUpdate_Date] = {_UserInfo._date.dateTimeToStrSql()} ,
       [LastUpdate_Time] = {_UserInfo._time.timeSpanToStrSql()} ,
       [Grade_Adjust] = {item.Grade_Adjust.intNToStrSql()} ,
       [Rating_Selection] = {item.Rating_Selection.stringToStrSql()}
";
                    if (item.Grade_Adjust.HasValue)
                        sql += $@" ,[Processing_Date] = {item.Processing_Date.dateTimeNToStrSql()} ";
                    sql += $@"
WHERE Reference_Nbr = {item.Reference_Nbr.stringToStrSql()}
AND   Report_Date = {item.Report_Date.dateTimeToStrSql()}
AND   Rating_Type = {item.Rating_Type.stringToStrSql()}
AND   Rating_Object = {item.Rating_Object.stringToStrSql()}
AND   Rating_Org_Area = {item.Rating_Org_Area.stringToStrSql()}
AND   Bond_Number = {item.Bond_Number.stringToStrSql()}
AND   Version = {item.Version.intNToStrSql()}
AND   Lots = {item.Lots.stringToStrSql()}
AND   Portfolio_Name = {item.Portfolio_Name.stringToStrSql()}
";
                    sbs.Add(new StringBuilder(sql));
                }
                foreach (var item in InsertA58s.GroupBy(z =>
                new { z.Reference_Nbr, z.Bond_Number, z.Lots, z.Rating_Type, z.Report_Date, z.Origination_Date, z.Version, z.Portfolio_Name }))
                {
                    sbs.Add(new StringBuilder(
                        $@"
Delete Bond_Rating_Summary
where Reference_Nbr = {item.Key.Reference_Nbr.stringToStrSql()}
and Bond_Number = {item.Key.Bond_Number.stringToStrSql()}
and Lots = {item.Key.Lots.stringToStrSql()}
and Rating_Type = {item.Key.Rating_Type.stringToStrSql()}
and Report_Date = {item.Key.Report_Date.dateTimeToStrSql()}
and Origination_Date = {item.Key.Origination_Date.dateTimeNToStrSql()}
and Version = {item.Key.Version.intNToStrSql()}
and Portfolio_Name = {item.Key.Portfolio_Name.stringToStrSql()}
and Rating_Object = ' '
and Rating_Selection is null
and Rating_Org_Area = ' ' ;
"
                        ));
                }

                #endregion 新增 & 修改 A57,A58

                if (!sbs.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.Not_Find_Relating_Missing.GetDescription();   //190222 修改錯誤訊息內容
                    return result;
                }

                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        int size = 10000;
                        for (int q = 0; (sbs.Count() / size) >= q; q += 1)
                        {
                            StringBuilder sql = new StringBuilder();
                            sbs.Skip((q) * size).Take(size).ToList()
                                .ForEach(x =>
                                {
                                    sql.Append(x.ToString());
                                });
                            if(sql.Length >0)
                            db.Database.ExecuteSqlCommand(sql.ToString());
                        }
                        dbContextTransaction.Commit();
                    }
                    catch (Exception ex)
                    {
                        //dbContextTransaction.Rollback(); //Required according to MSDN article
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = ex.exceptionMessage();
                        return result;
                    }
                }
            }
            result.RETURN_FLAG = true;
            result.DESCRIPTION = Message_Type.save_Success.GetDescription();
            return result;
        }

        /// <summary>
        /// 手動轉換 A57 & A58 (執行信評轉檔)
        /// </summary>
        /// <param name="date">基準日</param>
        /// <param name="version">版本</param>
        /// <param name="complement">抓取上一版資料(評估日)</param>
        /// <param name="deleteFlag">重新執行最新版</param>
        /// <param name="A59Flag">A59補登Flag</param>
        /// <returns></returns>
        public MSGReturnModel saveA57A58(DateTime dt, int version, string complement, bool deleteFlag, bool A59Flag = false)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime startTime = DateTime.Now;
            DateTime endTime = DateTime.Now;
            List<Bond_Rating_Parm> parmIDs = getParmIDs(); //選取有效的D60
            var _checkParm = checkParmID(parmIDs);
            if (!_checkParm.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = _checkParm;
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var A41 = db.Bond_Account_Info.AsNoTracking()
                    .Where(x => x.Report_Date == dt && x.Version == version);
                var A53Data = db.Rating_Info.AsNoTracking().Where(x => x.Report_Date == dt);
                if (!A41.Any())
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A41.tableNameGetDescription());
                }
                else if (!A59Flag && !A53Data.Any())
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A53.tableNameGetDescription());
                }
                else if (!parmIDs.Any())
                {
                    result.DESCRIPTION = Table_Type.D60.tableNameGetDescription() + "參數設定有誤";
                }
                else
                {
                    string reportData = dt.ToString("yyyy/MM/dd");
                    string ver = version.ToString();

                    var _checkSql = $@"
select top 1 Report_Date,Version
from Bond_Rating_Summary
where 
(Report_Date = '{reportData}'
and Version < '{ver}'
and  Rating_Type = '{Rating_Type.A.GetDescription()}' ) OR
(Report_Date < '{reportData}'
and   Rating_Type = '{Rating_Type.A.GetDescription()}' )
group by Report_Date,Version
order by Report_Date desc,Version desc ;
";

                    var oldA57data = db.Database.DynamicSqlQuery(_checkSql);
                    DateTime? oldReportDate = null;
                    int? oldVersion = null;
                    if (oldA57data.Any())
                    {
                        oldReportDate = oldA57data[0].Report_Date;
                        oldVersion = oldA57data[0].Version;
                        var _oldRating_Type = Rating_Type.A.GetDescription();
                        var _oldDatas = db.Bond_Rating_Info.AsNoTracking()
                            .Where(x => x.Report_Date == oldReportDate &&
                                      x.Version == oldVersion &&
                                      x.Rating_Type == _oldRating_Type &&
                                      x.Rating != null &&
                                      x.Grade_Adjust == null).ToList();
                        if (_oldDatas.Any())
                        {
                            new BondsCheckRepository<Bond_Rating_Info>(_oldDatas, Check_Table_Type.Bonds_A58_Before_Check,dt,version);
                            result.DESCRIPTION = Message_Type.Check_Fail.GetDescription();
                            return result;
                        }
                    }

                    using (var dbContextTransaction = db.Database.BeginTransaction())
                    {
                        try
                        {
                            //重新執行該基準日最大版本 2017/12/22 新增需求 (刪掉目前最大版本的紀錄檔)
                            if (deleteFlag)
                            {
                                string sql = string.Empty;
                                sql += $@"
                    UPDATE Transfer_CheckTable
                    Set TransferType = 'R'
                    WHERE ReportDate = '{reportData}'
                    AND Version = {ver}
                    AND TransferType = 'Y'
                    AND File_Name IN ('{Table_Type.A57.ToString()}','{Table_Type.A58.ToString()}') ;

                    delete from Bond_Rating_Summary
                    where Report_Date = '{reportData}'
                    and Version = {ver} ;

                    delete from Bond_Rating_Info
                    where Report_Date = '{reportData}'
                    and Version = {ver} ; ";
                                db.Database.ExecuteSqlCommand(sql);
                            }
                            //判斷轉檔紀錄
                            if (!A59Flag && !deleteFlag &&
                                (!common.checkTransferCheck(Table_Type.A57.ToString(), Table_Type.A41.ToString(), dt, version) ||
                                !common.checkTransferCheck(Table_Type.A58.ToString(), Table_Type.A53.ToString(), dt, version)))
                            {
                                result.DESCRIPTION = $@"信評評估已執行過 => 評估日:{dt.ToString("yyyy/MM/dd")},版本:{version.ToString()}";
                            }
                            else if (A59Flag && !deleteFlag && (!common.checkTransferCheck(Table_Type.A57.ToString(), Table_Type.A41.ToString(), dt, version)))
                            {
                                result.DESCRIPTION = $@"信評評估已執行過 => 評估日:{dt.ToString("yyyy/MM/dd")},版本:{version.ToString()}";
                            }
                            else
                            {
                                string _key = Controllers.AccountController.CurrentUserInfo.Name + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss");

                                #region sql

                                List<StringBuilder> sbt = new List<StringBuilder>();
                                string sql = string.Empty; //原始
                                string sql1_2 = string.Empty; //評估日
                                StringBuilder sql_A = new StringBuilder(); //新增三張信評特殊表格
                                string sql_D = string.Empty; //排除多餘
                                string sql1_3 = string.Empty; //更新三張信評特殊表格
                                string sql2 = string.Empty; //A58

                                #region A57原始日 sql

                                sql = $@"
--原始
WITH temp AS
(
select RTG_Bloomberg_Field,
       Rating_Object,
       Bond_Number,
       Lots,
       Portfolio_Name,
	   Bond_Number_Old,
	   Lots_Old,
	   Portfolio_Name_Old,
       Rating,
       Rating_Date,
       Rating_Org,
       A57.Report_Date
FROM   Bond_Rating_Info A57 --,temp2
WHERE  A57.Report_Date =  {oldReportDate.dateTimeNToStrSql()}  --temp2.Report_Date 
AND    A57.Version = {oldVersion.intNToStrSql()}  --temp2.Version
AND    A57.Rating_Type = '{Rating_Type.A.GetDescription()}'
AND    A57.Grade_Adjust is not null
), --最後一版A57(原始投資信評)
A52 AS (
   SELECT * FROM Grade_Mapping_Info
   WHERE IsActive = 'Y'
),
A51 AS (
   SELECT * FROM Grade_Moody_Info
   WHERE Status = '1'
),
tempA57t AS (
SELECT _A57.*,
       _A51.Grade_Adjust,
       _A51.PD_Grade
FROM   (select * from temp where Rating_Org <> '{RatingOrg.Moody.GetDescription()}') AS _A57
left Join   A52 _A52
ON     _A57.Rating = _A52.Rating
AND    _A57.Rating_Org = _A52.Rating_Org
left JOIN   A51 _A51
ON     _A52.PD_Grade = _A51.PD_Grade
  UNION ALL
SELECT _A57.*,
       _A51_2.Grade_Adjust,
       _A51_2.PD_Grade
FROM   (select * from temp where Rating_Org = '{RatingOrg.Moody.GetDescription()}' ) AS _A57
left JOIN   A51 _A51_2
ON     _A57.Rating = _A51_2.Rating
),
--最後一版A57(原始投資信評) 加入信評
tempA AS(
   SELECT A41.Reference_Nbr,
          A41.Bond_Number,
          A41.Lots,
          A41.Portfolio,
          A41.Segment_Name,
          A41.Bond_Type,
          A41.Lien_position,
          A41.Origination_Date,
          A41.Report_Date,
		  A41.Portfolio_Name,
		  A41.PRODUCT,
		  A41.ISSUER,
		  A41.Version,
		  A41.ISIN_Changed_Ind,
          A41.Bond_Number_Old,
          A41.Lots_Old,
          A41.Portfolio_Name_Old,
          A41.Origination_Date_Old,
      	  _A53Sample.ISSUER_TICKER,
		  _A53Sample.GUARANTOR_NAME,
		  _A53Sample.GUARANTOR_EQY_TICKER
   FROM (Select *
   from Bond_Account_Info
   where Report_Date = '{reportData}'
   and   Version =  {ver})
   AS A41
   LEFT JOIN (select * from Rating_Info_SampleInfo where Report_Date = '{reportData}') AS _A53Sample
   ON  A41.Bond_Number = _A53Sample.Bond_Number
), --全部的A41
T1 AS (
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  oldA57.Rating_Date AS Rating_Date,
          CASE WHEN oldA57.Rating_Object is null
               THEN ' '
               ELSE oldA57.Rating_Object
          END AS Rating_Object,
          oldA57.Rating_Org AS Rating_Org,
		  oldA57.Rating AS Rating,
		  (CASE WHEN oldA57.Rating_Org in ('{RatingOrg.SP.GetDescription()}','{RatingOrg.Moody.GetDescription()}','{RatingOrg.Fitch.GetDescription()}') THEN '國外'
	            WHEN oldA57.Rating_Org in ('{RatingOrg.FitchTwn.GetDescription()}','{RatingOrg.CW.GetDescription()}') THEN '國內'
                ELSE ' '
	      END) AS Rating_Org_Area,
          oldA57.PD_Grade,
          oldA57.Grade_Adjust,
		  CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
		  CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
		  CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
          CASE WHEN oldA57.RTG_Bloomberg_Field is null
               THEN ' '
               ELSE oldA57.RTG_Bloomberg_Field
          END AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  (select * from tempA where ISIN_Changed_Ind is null) AS BA_Info --沒換券的A41
   JOIN  tempA57t oldA57 --oldA57
   ON    BA_Info.Bond_Number =  oldA57.Bond_Number
   AND   BA_Info.Lots = oldA57.Lots
   AND   BA_Info.Portfolio_Name = oldA57.Portfolio_Name 
UNION ALL
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  oldA57.Rating_Date AS Rating_Date,
          CASE WHEN oldA57.Rating_Object is null
               THEN ' '
               ELSE oldA57.Rating_Object
          END AS Rating_Object,
          oldA57.Rating_Org AS Rating_Org,
		  oldA57.Rating AS Rating,
		  (CASE WHEN oldA57.Rating_Org in ('{RatingOrg.SP.GetDescription()}','{RatingOrg.Moody.GetDescription()}','{RatingOrg.Fitch.GetDescription()}') THEN '國外'
	            WHEN oldA57.Rating_Org in ('{RatingOrg.FitchTwn.GetDescription()}','{RatingOrg.CW.GetDescription()}') THEN '國內'
                ELSE ' '
	      END) AS Rating_Org_Area,
          oldA57.PD_Grade,
          oldA57.Grade_Adjust,
		  CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
		  CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
		  CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
          CASE WHEN oldA57.RTG_Bloomberg_Field is null
               THEN ' '
               ELSE oldA57.RTG_Bloomberg_Field
          END AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  (select * from tempA where ISIN_Changed_Ind is null) AS BA_Info --沒換券的A41
   JOIN  (select * from tempA57t where Bond_Number_Old is not null) AS oldA57 --oldA57
   ON    BA_Info.Bond_Number =  oldA57.Bond_Number_Old
   AND   BA_Info.Lots = oldA57.Lots_Old
   AND   BA_Info.Portfolio_Name = oldA57.Portfolio_Name_Old
UNION ALL
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  oldA57.Rating_Date AS Rating_Date,
          CASE WHEN oldA57.Rating_Object is null
               THEN ' '
               ELSE oldA57.Rating_Object
          END AS Rating_Object,
          oldA57.Rating_Org AS Rating_Org,
		  oldA57.Rating AS Rating,
		  (CASE WHEN oldA57.Rating_Org in ('{RatingOrg.SP.GetDescription()}','{RatingOrg.Moody.GetDescription()}','{RatingOrg.Fitch.GetDescription()}') THEN '國外'
	            WHEN oldA57.Rating_Org in ('{RatingOrg.FitchTwn.GetDescription()}','{RatingOrg.CW.GetDescription()}') THEN '國內'
                ELSE ' '
	      END) AS Rating_Org_Area,
          oldA57.PD_Grade,
          oldA57.Grade_Adjust,
          CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
          CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
          CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
          CASE WHEN oldA57.RTG_Bloomberg_Field is null
               THEN ' '
               ELSE oldA57.RTG_Bloomberg_Field
          END AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  (select * from tempA where ISIN_Changed_Ind = 'Y') AS BA_Info --換券的A41
   JOIN  tempA57t oldA57 --oldA57
   ON    BA_Info.Bond_Number_Old = oldA57.Bond_Number
   AND   BA_Info.Lots_Old = oldA57.Lots
   AND   BA_Info.Portfolio_Name_Old = oldA57.Portfolio_Name
UNION ALL
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  oldA57.Rating_Date AS Rating_Date,
          CASE WHEN oldA57.Rating_Object is null
               THEN ' '
               ELSE oldA57.Rating_Object
          END AS Rating_Object,
          oldA57.Rating_Org AS Rating_Org,
		  oldA57.Rating AS Rating,
		  (CASE WHEN oldA57.Rating_Org in ('{RatingOrg.SP.GetDescription()}','{RatingOrg.Moody.GetDescription()}','{RatingOrg.Fitch.GetDescription()}') THEN '國外'
	            WHEN oldA57.Rating_Org in ('{RatingOrg.FitchTwn.GetDescription()}','{RatingOrg.CW.GetDescription()}') THEN '國內'
                ELSE ' '
	      END) AS Rating_Org_Area,
          oldA57.PD_Grade,
          oldA57.Grade_Adjust,
          CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
          CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
          CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
          CASE WHEN oldA57.RTG_Bloomberg_Field is null
               THEN ' '
               ELSE oldA57.RTG_Bloomberg_Field
          END AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  (select * from tempA where ISIN_Changed_Ind = 'Y') AS BA_Info --換券的A41
   JOIN  (select * from tempA57t where Bond_Number_Old is not null) AS oldA57 --oldA57
   ON    BA_Info.Bond_Number_Old =  oldA57.Bond_Number_Old
   AND   BA_Info.Lots_Old = oldA57.Lots_Old
   AND   BA_Info.Portfolio_Name_Old = oldA57.Portfolio_Name_Old
),
T1s AS(
Select BA_Info.Reference_Nbr AS Reference_Nbr ,
    BA_Info.Bond_Number AS Bond_Number,
    BA_Info.Lots AS Lots,
    BA_Info.Portfolio AS Portfolio,
    BA_Info.Segment_Name AS Segment_Name,
    BA_Info.Bond_Type AS Bond_Type,
    BA_Info.Lien_position AS Lien_position,
    BA_Info.Origination_Date AS Origination_Date,
    BA_Info.Report_Date AS Report_Date,
    null AS Rating_Date,
    ' '  AS Rating_Object,
    null AS Rating_Org,
    null AS Rating,
    ' '  AS Rating_Org_Area,
    null AS PD_Grade,
    null AS Grade_Adjust,
    CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
    CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
    CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
    BA_Info.Portfolio_Name AS Portfolio_Name,
    ' '  AS RTG_Bloomberg_Field,
    BA_Info.PRODUCT AS SMF,
    BA_Info.ISSUER AS ISSUER,
    BA_Info.Version AS Version,
    CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
         THEN BA_Info.Bond_Number + ' Mtge'
	    ELSE BA_Info.Bond_Number + ' Corp'
    END AS Security_Ticker,
    BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
    BA_Info.Bond_Number_Old AS Bond_Number_Old,
    BA_Info.Lots_Old AS Lots_Old,
    BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
    BA_Info.Origination_Date_Old AS Origination_Date_Old
from tempA BA_Info
),
T1all AS(
  select T1.*, D60.Parm_ID AS Parm_ID from T1
  left join
  (select * from Bond_Rating_Parm where IsActive = 'Y') D60
  on T1.Rating_Object = D60.Rating_Object
  and T1.Rating_Org_Area = 
      case when D60.Rating_Org_Area is null
  	     then T1.Rating_Org_Area
  		 else D60.Rating_Org_Area
      end
  UNION ALL
  select *, null AS Parm_ID from T1s
)
INSERT INTO Bond_Rating_Info
           (Reference_Nbr,
           Bond_Number,
           Lots,
           Portfolio,
           Segment_Name,
           Bond_Type,
           Lien_position,
           Origination_Date,
           Report_Date,
           Rating_Date,
           Rating_Type,
           Rating_Object,
           Rating_Org,
           Rating,
           Rating_Org_Area,
           PD_Grade,
           Grade_Adjust,
           ISSUER_TICKER,
           GUARANTOR_NAME,
           GUARANTOR_EQY_TICKER,
           Parm_ID,
           Portfolio_Name,
           RTG_Bloomberg_Field,
           SMF,
           ISSUER,
           Version,
           Security_Ticker,
           ISIN_Changed_Ind,
           Bond_Number_Old,
           Lots_Old,
           Portfolio_Name_Old,
           Origination_Date_Old,
           Create_User,
           Create_Date,
           Create_Time)
SELECT     Reference_Nbr,
		   Bond_Number,
		   Lots,
           Portfolio,
           Segment_Name,
           Bond_Type,
           Lien_position,
           Origination_Date,
           Report_Date,
           Rating_Date,
           '{Rating_Type.A.GetDescription()}',
           Rating_Object,
           Rating_Org,
           Rating,
           Rating_Org_Area,
           PD_Grade,
           Grade_Adjust,
           ISSUER_TICKER,
           GUARANTOR_NAME,
           GUARANTOR_EQY_TICKER,
           Parm_ID,
           Portfolio_Name,
           RTG_Bloomberg_Field,
           SMF,
           ISSUER,
           Version,
           Security_Ticker,
           ISIN_Changed_Ind,
           Bond_Number_Old,
           Lots_Old,
           Portfolio_Name_Old,
           Origination_Date_Old,
           {_UserInfo._user.stringToStrSql()},
           {_UserInfo._date.dateTimeToStrSql()},
           {_UserInfo._time.timeSpanToStrSql()}
		   From
		   T1all ; ";

                                //刪除A57(原始)預設
                                sql_D += $@"
with tempDeleA as
(
Select distinct Reference_Nbr from Bond_Rating_Info
where  Report_Date = '{reportData}'
and Version = {ver}
and Rating_Type = '{Rating_Type.A.GetDescription()}'
and Rating is not null
)
delete Bond_Rating_Info
where  Report_Date = '{reportData}'
and Version = {ver}
and Rating is null
and Rating_Type = '{Rating_Type.A.GetDescription()}'
and Reference_Nbr in (select Reference_Nbr from tempDeleA); ";

                                #endregion A57原始日 sql

                                #region A57評估日 sql

                                sql1_2 = $@"
 --評估日

WITH A52 AS (
   SELECT * FROM Grade_Mapping_Info
   WHERE IsActive = 'Y'
),
A51 AS (
   SELECT * FROM Grade_Moody_Info
   WHERE Status = '1'
),
A53 AS (
  SELECT  Bond_Number,
          Rating_Date,
          Rating_Org,
		  Rating,
		  Rating_Object,
		  RTG_Bloomberg_Field
    FROM  Rating_Info
    WHERE Report_Date = '{reportData}'
    AND   RTG_Bloomberg_Field not in ('G_RTG_MDY_LOCAL_LT_BANK_DEPOSITS','RTG_MDY_LOCAL_LT_BANK_DEPOSITS') --穆迪長期本國銀行存款評等,在寫A57時不要寫進去放成空值,風控暫時不用它來判斷熟調或熟低。
),
A53t AS (
SELECT _A53.*,
       _A51.Grade_Adjust,
       _A51.PD_Grade
FROM  (select * from A53 where Rating_Org <> '{RatingOrg.Moody.GetDescription()}') AS _A53
left Join   A52 _A52
ON     _A53.Rating_Org <> '{RatingOrg.Moody.GetDescription()}'
AND    _A53.Rating_Org = _A52.Rating_Org
AND    _A53.Rating = _A52.Rating
left JOIN   A51 _A51
ON     _A52.PD_Grade = _A51.PD_Grade
  UNION ALL
SELECT _A53.*,
       _A51_2.Grade_Adjust,
       _A51_2.PD_Grade
FROM   (select * from A53 where Rating_Org = '{RatingOrg.Moody.GetDescription()}') AS _A53
left JOIN   A51 _A51_2
ON     _A53.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
AND    _A53.Rating = _A51_2.Rating
),
tempC AS (
   SELECT   A41.Reference_Nbr,
            A41.Bond_Number,
            A41.Lots,
            A41.Portfolio,
            A41.Segment_Name,
            A41.Bond_Type,
            A41.Lien_position,
            A41.Origination_Date,
            A41.Report_Date,
		    A41.Portfolio_Name,
		    A41.PRODUCT,
		    A41.ISSUER,
		    A41.Version,
		    A41.ISIN_Changed_Ind,
            A41.Bond_Number_Old,
            A41.Lots_Old,
            A41.Portfolio_Name_Old,
            A41.Origination_Date_Old,
	        _A53Sample.ISSUER_TICKER,
			_A53Sample.GUARANTOR_NAME,
			_A53Sample.GUARANTOR_EQY_TICKER
	 FROM (select * from Bond_Account_Info
     where Report_Date = '{reportData}'
     and VERSION = {ver}
     ) AS A41
	 LEFT JOIN (select * from Rating_Info_SampleInfo where Report_Date = '{reportData}') AS _A53Sample
	 ON   A41.Bond_Number = _A53Sample.Bond_Number
),
T0 AS (
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  RA_Info.Rating_Date AS Rating_Date,
          CASE WHEN RA_Info.Rating_Object is null
               THEN ' '
               ELSE RA_Info.Rating_Object
          END AS Rating_Object,
          RA_Info.Rating_Org AS Rating_Org,
		  RA_Info.Rating AS Rating,
		  (CASE WHEN RA_Info.Rating_Org in ('{RatingOrg.SP.GetDescription()}','{RatingOrg.Moody.GetDescription()}','{RatingOrg.Fitch.GetDescription()}') THEN '國外'
	            WHEN RA_Info.Rating_Org in ('{RatingOrg.FitchTwn.GetDescription()}','{RatingOrg.CW.GetDescription()}') THEN '國內'
                ELSE ' '
	      END) AS Rating_Org_Area,
          RA_Info.PD_Grade,
          RA_Info.Grade_Adjust,
		  CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
		  CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
		  CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
		  CASE WHEN RA_Info.RTG_Bloomberg_Field is null
               THEN ' '
               ELSE RA_Info.RTG_Bloomberg_Field
          END AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  ( Select * from tempC) BA_Info --A41  
   JOIN A53t RA_Info --A53
   ON BA_Info.Bond_Number = RA_Info.Bond_Number
   UNION ALL
   SELECT BA_Info.Reference_Nbr AS Reference_Nbr ,
          BA_Info.Bond_Number AS Bond_Number,
		  BA_Info.Lots AS Lots,
		  BA_Info.Portfolio AS Portfolio,
		  BA_Info.Segment_Name AS Segment_Name,
		  BA_Info.Bond_Type AS Bond_Type,
		  BA_Info.Lien_position AS Lien_position,
		  BA_Info.Origination_Date AS Origination_Date,
          BA_Info.Report_Date AS Report_Date,
		  null,
          ' ' AS Rating_Object,
          null,
		  null,
		  ' ' AS Rating_Org_Area,
          null,
          null,
		  CASE WHEN BA_Info.ISSUER_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.ISSUER_TICKER END AS ISSUER_TICKER,
		  CASE WHEN BA_Info.GUARANTOR_NAME in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_NAME END AS GUARANTOR_NAME,
		  CASE WHEN BA_Info.GUARANTOR_EQY_TICKER in ('N.S.', 'N.A.') THEN null Else BA_Info.GUARANTOR_EQY_TICKER END AS GUARANTOR_EQY_TICKER,
		  BA_Info.Portfolio_Name AS Portfolio_Name,
		  ' ' AS RTG_Bloomberg_Field,
		  BA_Info.PRODUCT AS SMF,
		  BA_Info.ISSUER AS ISSUER,
		  BA_Info.Version AS Version,
		  (CASE WHEN BA_Info.PRODUCT like 'A11%' OR BA_Info.PRODUCT like '932%'
		  THEN BA_Info.Bond_Number + ' Mtge' ELSE
		    BA_Info.Bond_Number + ' Corp' END) AS Security_Ticker,
          BA_Info.ISIN_Changed_Ind AS ISIN_Changed_Ind,
          BA_Info.Bond_Number_Old AS Bond_Number_Old,
          BA_Info.Lots_Old AS Lots_Old,
          BA_Info.Portfolio_Name_Old AS Portfolio_Name_Old,
          BA_Info.Origination_Date_Old AS Origination_Date_Old
   FROM  tempC BA_Info --A41
)
Insert into Bond_Rating_Info
           (Reference_Nbr,
           Bond_Number,
           Lots,
           Portfolio,
           Segment_Name,
           Bond_Type,
           Lien_position,
           Origination_Date,
           Report_Date,
           Rating_Date,
           Rating_Type,
           Rating_Object,
           Rating_Org,
           Rating,
           Rating_Org_Area,
           PD_Grade,
           Grade_Adjust,
           ISSUER_TICKER,
           GUARANTOR_NAME,
           GUARANTOR_EQY_TICKER,
           Parm_ID,
           Portfolio_Name,
           RTG_Bloomberg_Field,
           SMF,
           ISSUER,
           Version,
           Security_Ticker,
           ISIN_Changed_Ind,
           Bond_Number_Old,
           Lots_Old,
           Portfolio_Name_Old,
           Origination_Date_Old,
           Create_User,
           Create_Date,
           Create_Time)
SELECT     T0.Reference_Nbr,
		   T0.Bond_Number,
		   T0.Lots,
           T0.Portfolio,
           T0.Segment_Name,
           T0.Bond_Type,
           T0.Lien_position,
           T0.Origination_Date,
           T0.Report_Date,
           T0.Rating_Date,
           '{Rating_Type.B.GetDescription()}',
           T0.Rating_Object,
           T0.Rating_Org,
           T0.Rating,
           T0.Rating_Org_Area,
           T0.PD_Grade,
           T0.Grade_Adjust,
           T0.ISSUER_TICKER,
           T0.GUARANTOR_NAME,
           T0.GUARANTOR_EQY_TICKER,
           D60.Parm_ID,
           T0.Portfolio_Name,
           T0.RTG_Bloomberg_Field,
           T0.SMF,
           T0.ISSUER,
           T0.Version,
           T0.Security_Ticker,
           T0.ISIN_Changed_Ind,
           T0.Bond_Number_Old,
           T0.Lots_Old,
           T0.Portfolio_Name_Old,
           T0.Origination_Date_Old,
           {_UserInfo._user.stringToStrSql()},
           {_UserInfo._date.dateTimeToStrSql()},
           {_UserInfo._time.timeSpanToStrSql()}
		   From T0
  left join
  (select * from Bond_Rating_Parm where  IsActive = 'Y') D60
  on T0.Rating_Object = D60.Rating_Object
  and T0.Rating_Org_Area = 
      case when D60.Rating_Org_Area is null
  	     then T0.Rating_Org_Area
  		 else D60.Rating_Org_Area
      end
; ";

                                //刪除A57(評估日)預設
                                sql_D += $@"
with tempDeleB as
(
Select distinct Reference_Nbr from Bond_Rating_Info
where  Report_Date = '{reportData}'
and Version = {ver}
and Rating_Type = '{Rating_Type.B.GetDescription()}'
and Rating is not null
)
delete Bond_Rating_Info
where  Report_Date = '{reportData}'
and Version = {ver}
and Rating is null
and Rating_Type = '{Rating_Type.B.GetDescription()}'
and Reference_Nbr in (select Reference_Nbr from tempDeleB);
";

                                #endregion A57評估日 sql

                                #region 三張特殊表單更新 sql

                                sql1_3 = $@"
                                -- update Bond_Rating(債項信評) 的設定
                                update Bond_Rating_Info
                                set Rating =
                                      case
                                	    when Bond_Rating_Info.Rating_Org = '{RatingOrg.SP.GetDescription()}'
                                	     then Bond_Rating.S_And_P
                                        when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                	     then Bond_Rating.Moodys
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.Fitch.GetDescription()}'
                                		 then Bond_Rating.Fitch
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.FitchTwn.GetDescription()}'
                                		 then Bond_Rating.Fitch_TW
                                		when  Bond_Rating_Info.Rating_Org = '{RatingOrg.CW.GetDescription()}'
                                		 then Bond_Rating.TRC
                                	  else Bond_Rating_Info.Rating
                                	  end,
                                    Rating_Date = null
                                from Bond_Rating
                                where  Bond_Rating_Info.Bond_Number = Bond_Rating.Bond_Number
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.Bonds.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;

                                update Bond_Rating_Info
                                set PD_Grade =
                                       case
                                         when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                         then (select top 1 PD_Grade from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                	     else (select top 1 PD_Grade from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                       end ,
                                    Grade_Adjust =
                                         case
                                	     when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                		 then (select top 1 Grade_Adjust from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                		 else (select top 1 Grade_Adjust from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                      end
                                from Bond_Rating
                                where  Bond_Rating_Info.Bond_Number = Bond_Rating.Bond_Number
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.Bonds.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;

                                -- update Issuer_Rating(發行者信評) 的設定
                                update Bond_Rating_Info
                                set Rating =
                                      case
                                	    when Bond_Rating_Info.Rating_Org = '{RatingOrg.SP.GetDescription()}' 
                                	     then Issuer_Rating.S_And_P
                                        when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}' 
                                	     then Issuer_Rating.Moodys
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.Fitch.GetDescription()}' 
                                		 then Issuer_Rating.Fitch
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.FitchTwn.GetDescription()}' 
                                		 then Issuer_Rating.Fitch_TW
                                		when  Bond_Rating_Info.Rating_Org = '{RatingOrg.CW.GetDescription()}' 
                                		 then Issuer_Rating.TRC
                                	  else Bond_Rating_Info.Rating
                                	  end,
                                    Rating_Date = null
                                from Issuer_Rating
                                where  Bond_Rating_Info.ISSUER = Issuer_Rating.Issuer
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.ISSUER.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;

                                update Bond_Rating_Info
                                set PD_Grade =
                                       case
                                         when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                         then (select top 1 PD_Grade from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                	     else (select top 1 PD_Grade from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                       end ,
                                    Grade_Adjust =
                                         case
                                	     when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                		 then (select top 1 Grade_Adjust from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                		 else (select top 1 Grade_Adjust from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                      end
                                from Issuer_Rating
                                where  Bond_Rating_Info.ISSUER = Issuer_Rating.Issuer
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.ISSUER.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;

                                -- update Guarantor_Rating(擔保者信評) 的設定
                                update Bond_Rating_Info
                                set Rating =
                                      case
                                	    when Bond_Rating_Info.Rating_Org = '{RatingOrg.SP.GetDescription()}' 
                                	     then Guarantor_Rating.S_And_P
                                        when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}' 
                                	     then Guarantor_Rating.Moodys
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.Fitch.GetDescription()}' 
                                		 then Guarantor_Rating.Fitch
                                		when Bond_Rating_Info.Rating_Org = '{RatingOrg.FitchTwn.GetDescription()}' 
                                		 then Guarantor_Rating.Fitch_TW
                                		when  Bond_Rating_Info.Rating_Org = '{RatingOrg.CW.GetDescription()}'
                                		 then Guarantor_Rating.TRC
                                	  else Bond_Rating_Info.Rating
                                	  end,
                                    Rating_Date = null
                                from Guarantor_Rating
                                where  Bond_Rating_Info.ISSUER = Guarantor_Rating.Issuer
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.GUARANTOR.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;

                                update Bond_Rating_Info
                                set PD_Grade =
                                       case
                                         when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                         then (select top 1 PD_Grade from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                	     else (select top 1 PD_Grade from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                       end ,
                                    Grade_Adjust =
                                         case
                                	     when Bond_Rating_Info.Rating_Org = '{RatingOrg.Moody.GetDescription()}'
                                		 then (select top 1 Grade_Adjust from Grade_Moody_Info A51 where A51.Status = '1' and A51.Rating = Bond_Rating_Info.Rating)
                                		 else (select top 1 Grade_Adjust from
                                	       Grade_Moody_Info A51
                                		   where A51.Status = '1' and A51.PD_Grade =
                                	    (select top 1 PD_Grade from Grade_Mapping_Info A52 where A52.Rating_Org = Bond_Rating_Info.Rating_Org and A52.Rating = Bond_Rating_Info.Rating and A52.IsActive = 'Y'))
                                      end
                                from Guarantor_Rating
                                where  Bond_Rating_Info.ISSUER = Guarantor_Rating.Issuer
                                and Bond_Rating_Info.Report_Date = '{reportData}'
                                and Bond_Rating_Info.Version = {ver}
                                and Bond_Rating_Info.Rating_Object = '{RatingObject.GUARANTOR.GetDescription()}'
                                and Bond_Rating_Info.Rating_Type = '{Rating_Type.B.GetDescription()}' ;
                                ";

                                #endregion 三張特殊表單更新 sql

                                #region A58 sql

                                sql2 = $@"
WITH A57 AS
(
  SELECT * FROM Bond_Rating_Info
  WHERE Report_Date = '{reportData}'
  AND Version = {ver}
),
T4 AS
(
Select BR_Info.Reference_Nbr,
       BR_Info.Report_Date,
	   BR_Info.Parm_ID,
	   BR_Info.Bond_Type,
	   BR_Info.Rating_Type,
	   BR_Info.Rating_Object,
	   BR_Info.Rating_Org_Area,
	   BR_Parm.Rating_Selection,
	   (CASE WHEN BR_Parm.Rating_Selection = '1'
	         THEN Min(BR_Info.Grade_Adjust)
			 WHEN BR_Parm.Rating_Selection = '2'
			 THEN Max(BR_Info.Grade_Adjust)
			 ELSE null  END) AS Grade_Adjust,
	   BR_Parm.Rating_Priority,
       '{startTime.ToString("yyyy/MM/dd")}' AS Processing_Date,
	   BR_Info.Version,
	   BR_Info.Bond_Number,
	   BR_Info.Lots,
	   BR_Info.Portfolio,
	   BR_Info.Origination_Date,
	   BR_Info.Portfolio_Name,
	   BR_Info.SMF,
	   BR_Info.ISSUER,
       BR_Info.ISIN_Changed_Ind,
       BR_Info.Bond_Number_Old,
       BR_Info.Lots_Old,
       BR_Info.Portfolio_Name_Old,
       BR_Info.Origination_Date_Old
From   A57 BR_Info
LEFT JOIN (select * from  Bond_Rating_Parm where IsActive = 'Y') BR_Parm
ON      BR_Info.Rating_Object = BR_Parm.Rating_Object  
AND     BR_Info.Parm_ID =  BR_Parm.Parm_ID
AND     BR_Info.Rating_Org_Area =
        CASE WHEN BR_Parm.Rating_Org_Area is null
             THEN BR_Info.Rating_Org_Area
             ELSE BR_Parm.Rating_Org_Area
        END 
--on         BR_Info.Parm_ID = BR_Parm.Parm_ID --2018/03/09 Parm_ID 改為流水號
--AND        BR_Info.Rating_Object = BR_Parm.Rating_Object
--AND        BR_Info.Rating_Org_Area = BR_Parm.Rating_Org_Area  --2017/11/01修改為不分國內外
GROUP BY BR_Info.Reference_Nbr,
         BR_Info.Report_Date,
		 BR_Info.Bond_Type,
	     BR_Info.Rating_Type,
	     BR_Info.Rating_Object,
	     BR_Info.Rating_Org_Area,
		 BR_Info.Parm_ID,
		 BR_Info.Version,
		 BR_Info.Bond_Number,
		 BR_Info.Lots,
	     BR_Info.Portfolio,
	     BR_Info.Origination_Date,
	     BR_Info.Portfolio_Name,
	     BR_Info.SMF,
	     BR_Info.ISSUER,
		 BR_Parm.Rating_Selection,
		 BR_Parm.Rating_Priority,
         BR_Info.ISIN_Changed_Ind,
         BR_Info.Bond_Number_Old,
         BR_Info.Lots_Old,
         BR_Info.Portfolio_Name_Old,
         BR_Info.Origination_Date_Old
)
Insert Into Bond_Rating_Summary
            (
			  Reference_Nbr,
              Report_Date,
              Parm_ID,
              Bond_Type,
              Rating_Type,
              Rating_Object,
              Rating_Org_Area,
              Rating_Selection,
              Grade_Adjust,
              Rating_Priority,
              Processing_Date,
              Version,
              Bond_Number,
              Lots,
              Portfolio,
              Origination_Date,
              Portfolio_Name,
              SMF,
              ISSUER,
              ISIN_Changed_Ind,
              Bond_Number_Old,
              Lots_Old,
              Portfolio_Name_Old,
              Origination_Date_Old,
              Create_User,
              Create_Date,
              Create_Time
			)
select
			  Reference_Nbr,
              Report_Date,
              Parm_ID,
              Bond_Type,
              Rating_Type,
              Rating_Object,
              Rating_Org_Area,
              Rating_Selection,
              Grade_Adjust,
              Rating_Priority,
              Processing_Date,
              Version,
              Bond_Number,
              Lots,
              Portfolio,
              Origination_Date,
              Portfolio_Name,
              SMF,
              ISSUER,
              ISIN_Changed_Ind,
              Bond_Number_Old,
              Lots_Old,
              Portfolio_Name_Old,
              Origination_Date_Old,
              {_UserInfo._user.stringToStrSql()},
              {_UserInfo._date.dateTimeToStrSql()},
              {_UserInfo._time.timeSpanToStrSql()}
 from T4;
                        ";

                                #endregion A58 sql

                                #endregion sql

                                //insert A57 原始日
                                db.Database.ExecuteSqlCommand(sql);
                                //insert A57 評估日
                                db.Database.ExecuteSqlCommand(sql1_2);

                                #region 共用

                                List<Grade_Mapping_Info> A52s = db.Grade_Mapping_Info.AsNoTracking()
                                    .Where(x => x.IsActive == "Y").ToList();
                                List<Grade_Moody_Info> A51s = db.Grade_Moody_Info.AsNoTracking()
                                    .Where(x => x.Status == "1").ToList();
                                Grade_Moody_Info A51 = null;
                                string Domestic = "國內";
                                string Foreign = "國外";

                                #endregion 共用

                                #region 新增3張特殊表單rating

                                var _ratingType = Rating_Type.B.GetDescription();
                                var _Bond_Ratings = db.Bond_Rating.AsNoTracking().ToList();
                                var _Issuer_Ratings = db.Issuer_Rating.AsNoTracking().ToList();
                                var _Guarantor_Ratings = db.Guarantor_Rating.AsNoTracking().ToList();
                                List<insertRating> _data = new List<insertRating>();
                                RatingObject _Rating_Object = RatingObject.Bonds;
                                string _RObject = string.Empty;
                                if (_Bond_Ratings.Any())
                                {
                                    _Rating_Object = RatingObject.Bonds;
                                    _RObject = _Rating_Object.GetDescription();
                                    var _Bonds = _Bond_Ratings.Select(x => x.Bond_Number).ToList();
                                    db.Bond_Rating_Info.AsNoTracking()
                                      .Where(x => x.Report_Date == dt &&
                                                 x.Version == version &&
                                                 x.Rating_Type == _ratingType &&
                                                 _Bonds.Contains(x.Bond_Number)).ToList()
                                                 .GroupBy(x => new { x.Reference_Nbr, x.Bond_Number })
                                      .ToList().ForEach(x =>
                                      {
                                          _data = new List<insertRating>();
                                          var _first = x.First();
                                          var _Bond_Rating = _Bond_Ratings.First(y => y.Bond_Number == x.Key.Bond_Number);
                                          if (!_Bond_Rating.S_And_P.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.SP.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.SP,
                                                  _Rating = _Bond_Rating.S_And_P,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_SP"
                                              });
                                          }
                                          if (!_Bond_Rating.Moodys.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Moody.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Moody,
                                                  _Rating = _Bond_Rating.Moodys,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_MOODY"
                                              });
                                          }
                                          if (!_Bond_Rating.Fitch.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Fitch.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Fitch,
                                                  _Rating = _Bond_Rating.Fitch,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_FITCH"
                                              });
                                          }
                                          if (!_Bond_Rating.Fitch_TW.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.FitchTwn.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.FitchTwn,
                                                  _Rating = _Bond_Rating.Fitch_TW,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "RTG_FITCH_NATIONAL"
                                              });
                                          }
                                          if (!_Bond_Rating.TRC.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.CW.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.CW,
                                                  _Rating = _Bond_Rating.TRC,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "RTG_TRC"
                                              });
                                          }
                                          insertA57Rating(_data, _first, sql_A, _Rating_Object, A52s, A51s, parmIDs);
                                      });
                                }
                                if (_Issuer_Ratings.Any())
                                {
                                    _Rating_Object = RatingObject.ISSUER;
                                    _RObject = _Rating_Object.GetDescription();
                                    var _Issuers = _Issuer_Ratings.Select(x => x.Issuer).ToList();
                                    db.Bond_Rating_Info.AsNoTracking()
                                      .Where(x => x.Report_Date == dt &&
                                                 x.Version == version &&
                                                 x.Rating_Type == _ratingType &&
                                                 _Issuers.Contains(x.ISSUER)).ToList()
                                                 .GroupBy(x => new { x.Reference_Nbr, x.ISSUER })
                                      .ToList().ForEach(x =>
                                      {
                                          _data = new List<insertRating>();
                                          var _first = x.First();
                                          var _Issuer_Rating = _Issuer_Ratings.First(y => y.Issuer == x.Key.ISSUER);
                                          if (!_Issuer_Rating.S_And_P.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.SP.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.SP,
                                                  _Rating = _Issuer_Rating.S_And_P,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_SP_LT_FC_ISSUER_CREDIT"
                                              });
                                          }
                                          if (!_Issuer_Rating.Moodys.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Moody.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Moody,
                                                  _Rating = _Issuer_Rating.Moodys,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_MDY_ISSUER"
                                              });
                                          }
                                          if (!_Issuer_Rating.Fitch.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Fitch.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Fitch,
                                                  _Rating = _Issuer_Rating.Fitch,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "RTG_FITCH_LT_ISSUER_DEFAULT"
                                              });
                                          }
                                          if (!_Issuer_Rating.Fitch_TW.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.FitchTwn.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.FitchTwn,
                                                  _Rating = _Issuer_Rating.Fitch_TW,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "RTG_FITCH_NATIONAL_LT"
                                              });
                                          }
                                          if (!_Issuer_Rating.TRC.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.CW.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.CW,
                                                  _Rating = _Issuer_Rating.TRC,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "RTG_TRC_LONG_TERM"
                                              });
                                          }
                                          insertA57Rating(_data, _first, sql_A, _Rating_Object, A52s, A51s, parmIDs);
                                      });
                                }
                                if (_Guarantor_Ratings.Any())
                                {
                                    _Rating_Object = RatingObject.GUARANTOR;
                                    _RObject = _Rating_Object.GetDescription();
                                    var _Guarantors = _Guarantor_Ratings.Select(x => x.Issuer).ToList();
                                    db.Bond_Rating_Info.AsNoTracking()
                                      .Where(x => x.Report_Date == dt &&
                                                 x.Version == version &&
                                                 x.Rating_Type == _ratingType &&
                                                 _Guarantors.Contains(x.ISSUER)).ToList()
                                                 .GroupBy(x => new { x.Reference_Nbr, x.ISSUER })
                                      .ToList().ForEach(x =>
                                      {
                                          _data = new List<insertRating>();
                                          var _first = x.First();
                                          var _Guarantor_Rating = _Guarantor_Ratings.First(y => y.Issuer == x.Key.ISSUER);
                                          if (!_Guarantor_Rating.S_And_P.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.SP.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.SP,
                                                  _Rating = _Guarantor_Rating.S_And_P,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "G_RTG_SP_LT_FC_ISSUER_CREDIT"
                                              });
                                          }
                                          if (!_Guarantor_Rating.Moodys.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Moody.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Moody,
                                                  _Rating = _Guarantor_Rating.Moodys,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "G_RTG_MDY_ISSUER"
                                              });
                                          }
                                          if (!_Guarantor_Rating.Fitch.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.Fitch.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.Fitch,
                                                  _Rating = _Guarantor_Rating.Fitch,
                                                  _RatingOrgArea = Foreign,
                                                  _RTG_Bloomberg_Field = "G_RTG_FITCH_LT_ISSUER_DEFAULT"
                                              });
                                          }
                                          if (!_Guarantor_Rating.Fitch_TW.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.FitchTwn.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.FitchTwn,
                                                  _Rating = _Guarantor_Rating.Fitch_TW,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "G_RTG_FITCH_NATIONAL_LT"
                                              });
                                          }
                                          if (!_Guarantor_Rating.TRC.IsNullOrWhiteSpace() &&
                                              !x.Any(z =>
                                              z.Rating_Object == _RObject &&
                                              z.Rating_Org == RatingOrg.CW.GetDescription()))
                                          {
                                              _data.Add(new insertRating()
                                              {
                                                  _Rating_Org = RatingOrg.CW,
                                                  _Rating = _Guarantor_Rating.TRC,
                                                  _RatingOrgArea = Domestic,
                                                  _RTG_Bloomberg_Field = "G_RTG_TRC_LONG_TERM"
                                              });
                                          }

                                          insertA57Rating(_data, _first, sql_A, _Rating_Object, A52s, A51s, parmIDs);
                                      });
                                }
                                if (sql_A.Length > 0)
                                    db.Database.ExecuteSqlCommand(sql_A.ToString());

                                #endregion 新增3張特殊表單rating
                              
                                //刪除A57 預設
                                db.Database.ExecuteSqlCommand(sql_D);

                                bool D60Flag = true; //D60設定是否正確

                                #region issuer='GOV-TW-CEN' or 'GOV-Kaohsiung' or 'GOV-TAIPEI' then他們的債項評等放他們發行人的評等(PS:發行人的評等複製一份給債項評等)

                                var ISSUERstr = RatingObject.ISSUER.GetDescription();
                                var Bondstr = RatingObject.Bonds.GetDescription();
                                List<string> ISSUERs = new List<string>() { "GOV-TW-CEN", "GOV-Kaohsiung", "GOV-TAIPEI" };
                                StringBuilder sb2 = new StringBuilder();
                                
                                var A57ISSUERs =
                                    db.Bond_Rating_Info.AsNoTracking().Where(x =>
                                     x.Report_Date == dt &&
                                     x.Version == version &&
                                     x.Grade_Adjust != null &&
                                     x.PD_Grade != null &&
                                     ISSUERs.Contains(x.ISSUER) &&
                                     x.Rating_Object == ISSUERstr).ToList();
                                var A57Bonds =
                                     db.Bond_Rating_Info.AsNoTracking().Where(x =>
                                     x.Report_Date == dt &&
                                     x.Version == version &&
                                     ISSUERs.Contains(x.ISSUER) &&
                                     x.Rating_Object == Bondstr).ToList();
                                foreach (var A57group in A57ISSUERs.GroupBy(x => new
                                {
                                    x.Reference_Nbr,
                                    x.Bond_Number,
                                    x.Lots,
                                    x.Portfolio,
                                    x.Segment_Name,
                                    x.Bond_Type,
                                    x.Lien_position,
                                    x.Origination_Date,
                                    x.Report_Date,
                                    x.Rating_Type,
                                    x.Rating_Org,
                                    x.Rating_Object,
                                    x.ISSUER_TICKER,
                                    x.GUARANTOR_NAME,
                                    x.GUARANTOR_EQY_TICKER,
                                    x.Portfolio_Name,
                                    x.SMF,
                                    x.ISSUER,
                                    x.Version,
                                    x.Security_Ticker,
                                    x.ISIN_Changed_Ind,
                                    x.Bond_Number_Old,
                                    x.Lots_Old,
                                    x.Portfolio_Name_Old,
                                    x.Origination_Date_Old
                                }))
                                {
                                    var first = A57group.First();
                                    var _D60Foreign = getParmID(parmIDs, first.Rating_Object, Foreign); //國外
                                    var _D60Domestic = getParmID(parmIDs, first.Rating_Object, Domestic); //國內                                    
                                    if (_D60Foreign == null || _D60Domestic == null)
                                    {
                                        D60Flag = false;
                                    }
                                    if (D60Flag)
                                    {
                                        var _D60 = new Bond_Rating_Parm();
                                        if (_D60Foreign.Rating_Selection == _D60Domestic.Rating_Selection)
                                            _D60 = _D60Foreign;
                                        else
                                            _D60 = _D60Foreign.Rating_Priority <= _D60Domestic.Rating_Priority ? _D60Foreign : _D60Domestic;
                                        int? _PD_Grade = null;
                                        if (_D60 != null)
                                        {
                                            if (_D60.Rating_Selection == "1")
                                            {
                                                _PD_Grade = A57group.Where(z => z.PD_Grade != null).Min(z => z.PD_Grade);
                                            }
                                            if (_D60.Rating_Selection == "2")
                                            {
                                                _PD_Grade = A57group.Where(z => z.PD_Grade != null).Max(z => z.PD_Grade);
                                            }
                                        }
                                        if (_PD_Grade.HasValue)
                                        {
                                            var _A57ISSUERcopy = A57group.First(z => z.PD_Grade == _PD_Grade.Value);
                                            var _A57Bond = A57Bonds.FirstOrDefault(z =>
                                                                  z.Reference_Nbr == _A57ISSUERcopy.Reference_Nbr &&
                                                                  z.Rating_Type == _A57ISSUERcopy.Rating_Type &&
                                                                  z.Rating_Org == _A57ISSUERcopy.Rating_Org);
                                            if (_A57Bond == null) //沒有債項才須新增 (有債項就保持不變)
                                            {
                                                var RTGFiled = string.Empty;
                                                var RatingOrgArea = string.Empty;
                                                switch (_A57ISSUERcopy.Rating_Org)
                                                {
                                                    case "SP":
                                                        RTGFiled = "RTG_SP";
                                                        RatingOrgArea = Foreign;
                                                        break;

                                                    case "Moody":
                                                        RTGFiled = "RTG_MOODY";
                                                        RatingOrgArea = Foreign;
                                                        break;

                                                    case "Fitch":
                                                        RTGFiled = "RTG_FITCH";
                                                        RatingOrgArea = Foreign;
                                                        break;

                                                    case "Fitch(twn)":
                                                        RTGFiled = "RTG_FITCH_NATIONAL";
                                                        RatingOrgArea = Domestic;
                                                        break;

                                                    case "CW":
                                                        RTGFiled = "RTG_TRC";
                                                        RatingOrgArea = Domestic;
                                                        break;
                                                }
                                                if (!RTGFiled.IsNullOrWhiteSpace())
                                                {
                                                    var _parm = getParmID(parmIDs, Bondstr, RatingOrgArea);
                                                    _A57ISSUERcopy.Parm_ID = _parm?.Parm_ID.ToString();
                                                    _A57ISSUERcopy.Rating_Object = Bondstr;
                                                    _A57ISSUERcopy.Rating_Org_Area = RatingOrgArea;
                                                    _A57ISSUERcopy.Fill_up_YN = null;
                                                    _A57ISSUERcopy.Fill_up_Date = null;
                                                    _A57ISSUERcopy.RTG_Bloomberg_Field = RTGFiled;
                                                    sb2.Append(insertA57(_A57ISSUERcopy));
                                                }
                                            }
                                        }
                                    }
                                }
                                if (D60Flag && sb2.Length > 0)
                                {
                                    db.Database.ExecuteSqlCommand(sb2.ToString());
                                }

                                #endregion issuer='GOV-TW-CEN' or 'GOV-Kaohsiung' or 'GOV-TAIPEI' then他們的債項評等放他們發行人的評等(PS:發行人的評等複製一份給債項評等)

                                //三張特殊表單更新
                                if(D60Flag)
                                    db.Database.ExecuteSqlCommand(sql1_3);

                                #region 複寫前一版本已補登之信評

                                if (D60Flag && complement == "Y")
                                {
                                    StringBuilder sb = new StringBuilder();

                                    List<Bond_Rating_Info> InsertA57 = new List<Bond_Rating_Info>();
                                    List<Bond_Rating_Info> UpdateA57 = new List<Bond_Rating_Info>();

                                    var _Rating_Type = Rating_Type.B.GetDescription();
                                    var A57 = db.Bond_Rating_Info.AsNoTracking().Where(x => x.Report_Date == dt && x.Rating_Type == _Rating_Type);
                                    var _ver = A57.Where(x => x.Version != version).Max(x => x.Version);
                                    if (_ver != null)
                                    {
                                        //目前版本
                                        var A57nall = A57.Where(x => x.Version == version).ToList();
                                        //目前版本
                                        //上一次最後一版
                                        var A57s = A57.Where(x => x.Version == _ver && x.Fill_up_YN == "Y").ToList();
                                        //上一次最後一版
                                        var datas = A57s
                                            .GroupBy(x => new
                                            {
                                                x.Reference_Nbr,
                                                x.Bond_Number,
                                                x.Lots,
                                                x.Rating_Type,
                                                x.Report_Date,
                                                x.Version,
                                                x.Portfolio_Name
                                            });
                                        foreach (var item in datas)
                                        {
                                            bool deleteSqlFlag = false;
                                            var deleteA57 = new Bond_Rating_Info();
                                            var first = item.First();
                                            //A57n => 這一版A57變更共用資料(新增A57用)
                                            var A57ns = A57nall
                                                .Where(x => x.Bond_Number == first.Bond_Number &&
                                                            x.Lots == first.Lots &&
                                                            x.Portfolio_Name == first.Portfolio_Name).ToList();

                                            if (A57ns.Any()) //目前版本沒有比對到符合資料不需要繼續
                                            {
                                                var A57n = A57ns.First().ModelConvert<Bond_Rating_Info, Bond_Rating_Info>();
                                                //上一版有修改過的A57
                                                item.ToList().ForEach(
                                                x =>
                                                {
                                                    A51 = null;
                                                    if (x.Rating_Org == RatingOrg.Moody.GetDescription())
                                                    {
                                                        A51 = A51s.FirstOrDefault(z => z.Rating == x.Rating);
                                                    }
                                                    else
                                                    {
                                                        var A52 = A52s.FirstOrDefault(z => z.Rating_Org == x.Rating_Org &&
                                                                                           z.Rating == x.Rating);
                                                        if (A52 != null)
                                                            A51 = A51s.FirstOrDefault(z => z.PD_Grade == A52.PD_Grade);
                                                    }
                                                    //找尋目前版本存不存在相同的資料
                                                    var A57o = new Bond_Rating_Info();
                                                    A57o = A57ns.FirstOrDefault(y =>
                                                    y.Rating_Object == x.Rating_Object &&
                                                    y.RTG_Bloomberg_Field == x.RTG_Bloomberg_Field);
                                                    //不存在新增 防呆
                                                    if (A57o == null)
                                                    {
                                                        var _parm = getParmID(parmIDs, x.Rating_Object, x.Rating_Org_Area);
                                                        A57n.Parm_ID = _parm?.Parm_ID.ToString();
                                                        A57n.Rating_Date = x.Rating_Date;
                                                        A57n.Rating_Object = x.Rating_Object;
                                                        A57n.Rating_Org = x.Rating_Org;
                                                        A57n.Rating = x.Rating;
                                                        A57n.Rating_Org_Area = x.Rating_Org_Area;
                                                        A57n.Fill_up_YN = "Y";
                                                        A57n.Fill_up_Date = startTime;
                                                        A57n.PD_Grade = A51?.PD_Grade;
                                                        A57n.Grade_Adjust = A51?.Grade_Adjust;
                                                        A57n.RTG_Bloomberg_Field = x.RTG_Bloomberg_Field;
                                                        sb.Append(insertA57(A57n));
                                                        if (A51 != null)
                                                            deleteSqlFlag = true;
                                                    }
                                                    //存在修改
                                                    else if (A57o != null)
                                                    {
                                                        sb.Append($@"
                                                    UPDATE [Bond_Rating_Info]
                                                       SET [Rating] = {x.Rating.stringToStrSql()}
                                                          ,[PD_Grade] = {(A51?.PD_Grade).intNToStrSql()}
                                                          ,[Grade_Adjust] = {(A51?.Grade_Adjust).intNToStrSql()}
                                                          ,[Rating_Date] = {x.Rating_Date.dateTimeNToStrSql()}
                                                          ,[Fill_up_YN] = 'Y'
                                                          ,[Fill_up_Date] = '{startTime.ToString("yyyy/MM/dd")}'
                                                    WHERE Reference_Nbr = { A57o.Reference_Nbr.stringToStrSql() }
                                                      AND Report_Date = { A57o.Report_Date.dateTimeToStrSql() }
                                                      AND Rating_Type = { A57o.Rating_Type.stringToStrSql() }
                                                      AND Bond_Number = { A57o.Bond_Number.stringToStrSql() }
                                                      AND Lots = { A57o.Lots.stringToStrSql() }
                                                      AND RTG_Bloomberg_Field = { A57o.RTG_Bloomberg_Field.stringToStrSql() }
                                                      AND Portfolio_Name = { A57o.Portfolio_Name.stringToStrSql() }
                                                      AND Version = { A57o.Version.intNToStrSql() }; ");
                                                    }
                                                });
                                                if (deleteSqlFlag)
                                                {
                                                    sb.Append($@"
                            Delete Bond_Rating_Info
                            where Reference_Nbr = {A57n.Reference_Nbr.stringToStrSql()}
                            and Bond_Number = {A57n.Bond_Number.stringToStrSql()}
                            and Lots = {A57n.Lots.stringToStrSql()}
                            and Rating_Type = {A57n.Rating_Type.stringToStrSql()}
                            and Report_Date = {A57n.Report_Date.dateTimeToStrSql()}
                            and Version = {A57n.Version.intNToStrSql()}
                            and Portfolio_Name = {A57n.Portfolio_Name.stringToStrSql()}
                            and Rating_Object = ' '
                            and Rating_Org is null
                            and Rating is null
                            and Rating_Org_Area = ' ' ;  ");
                                                }
                                            }
                                        }
                                        if (sb.Length > 0)
                                        {
                                            db.Database.ExecuteSqlCommand(sb.ToString());
                                        }
                                    }
                                }

                                #endregion 複寫前一版本已補登之信評

                                if (D60Flag)
                                {
                                    //insert A58
                                    db.Database.ExecuteSqlCommand(sql2);
                                    dbContextTransaction.Commit();
                                    endTime = DateTime.Now;

                                    #region 執行信評轉檔後驗證

                                    new BondsCheckRepository<Bond_Rating_Summary>(
                                        db.Bond_Rating_Summary.AsNoTracking()
                                        .Where(x => x.Report_Date == dt &&
                                        x.Version == version).AsEnumerable(), Check_Table_Type.Bonds_A58_Transfer_Check, dt, version);

                                    #endregion 執行信評轉檔後驗證

                                    result.RETURN_FLAG = true;
                                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                                }
                                else
                                {
                                    result.DESCRIPTION = "D60參數設定未完全,請至D60設定參數。";
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            //dbContextTransaction.Rollback(); //Required according to MSDN article
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = ex.exceptionMessage();
                        }
                    }
                }
            }
            if (result.RETURN_FLAG)
            {
                //加入轉檔紀錄
                common.saveTransferCheck(
                    Table_Type.A57.ToString(),
                    true,
                    dt,
                    version,
                    startTime,
                    endTime);
                common.saveTransferCheck(
                    Table_Type.A58.ToString(),
                    true,
                    dt,
                    version,
                    startTime,
                    endTime);
            }
            else
            {
                //加入轉檔紀錄
                common.saveTransferCheck(
                    Table_Type.A57.ToString(),
                    false,
                    dt,
                    version,
                    startTime,
                    DateTime.Now,
                    result.DESCRIPTION);
                common.saveTransferCheck(
                    Table_Type.A58.ToString(),
                    false,
                    dt,
                    version,
                    startTime,
                    DateTime.Now,
                    result.DESCRIPTION);
            }
            return result;
        }

        /// <summary>
        /// A52(信評主標尺對應檔_其他) 複核
        /// </summary>
        /// <param name="Id"></param>
        /// <param name="Status"></param>
        /// <param name="Auditor_Reply"></param>
        /// <returns></returns>
        public MSGReturnModel A52Audit(int? Id, string Status, string Auditor_Reply)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            if (Id != null)
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var A52 = db.Grade_Mapping_Info.FirstOrDefault(x => x.Id == Id);
                    if (A52 != null)
                    {
                        if (Status == "1") //複核結果 1:接受 0:退回
                        {
                            switch (A52.Change_Status)
                            {
                                case "I":
                                    A52.IsActive = "Y";
                                    break;
                                case "D":
                                    A52.IsActive = "N";
                                    break;
                            }
                        }
                        A52.Status = Status;
                        A52.LastUpdate_User = _UserInfo._user;
                        A52.LastUpdate_Date = _UserInfo._date;
                        A52.LastUpdate_Time = _UserInfo._time;
                        A52.Auditor_Reply = Auditor_Reply;
                        A52.Processing_Date = _UserInfo._date;
                        A52.Audit_date = _UserInfo._date;
                        try
                        {
                            db.SaveChanges();
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();
                        }
                        catch (Exception ex)
                        {
                            result.DESCRIPTION = ex.exceptionMessage();
                        }
                    }
                }
            }
            return result;
        }

        #region saveA52

        public MSGReturnModel saveA52(string actionType, A52ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        var query = db.Grade_Mapping_Info
                                      .Where(x => x.Rating_Org == dataModel.Rating_Org
                                               && x.PD_Grade.ToString() == dataModel.PD_Grade
                                               && x.Rating == dataModel.Rating)
                                      .FirstOrDefault();
                        if (query != null && query.IsActive == "Y")
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：您輸入的資料已存在";
                            return result;
                        }
                        else if (query != null && query.IsActive == "N")
                        {
                            if (query.Status == null)
                            {
                                result.RETURN_FLAG = false;
                                result.DESCRIPTION = "資料重複：已有重複的資料等待複核";
                                return result;
                            }
                            query.Editor = _UserInfo._user;
                            query.Processing_Date = _UserInfo._date;
                            query.Auditor = dataModel.Auditor;
                            query.Auditor_Reply = null;
                            query.Audit_date = null;
                            query.Change_Status = "I"; //新增 I
                            query.LastUpdate_User = _UserInfo._user;
                            query.LastUpdate_Date = _UserInfo._date;
                            query.LastUpdate_Time = _UserInfo._time;
                            query.Status = null;
                        }
                        else
                        {
                            Grade_Mapping_Info addData = new Grade_Mapping_Info();

                            addData.Rating_Org = dataModel.Rating_Org;
                            addData.PD_Grade = int.Parse(dataModel.PD_Grade);
                            addData.Rating = dataModel.Rating;
                            addData.Editor = _UserInfo._user;
                            addData.Processing_Date = _UserInfo._date;
                            addData.Auditor = dataModel.Auditor;
                            addData.Change_Status = "I"; //新增 I
                            addData.IsActive = "N";
                            addData.Create_User = _UserInfo._user;
                            addData.Create_Date = _UserInfo._date;
                            addData.Create_Time = _UserInfo._time;
                            db.Grade_Mapping_Info.Add(addData);
                        }
                    }
                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "新增註記成功等待複核";
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }

        #endregion saveA52

        #region deleteA52

        public MSGReturnModel deleteA52(A52ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    int _PD_Grade = 0;
                    Int32.TryParse(dataModel.PD_Grade, out _PD_Grade);
                    var query = db.Grade_Mapping_Info
                                  .FirstOrDefault(x =>
                                  x.Rating_Org == dataModel.Rating_Org &&
                                  x.PD_Grade == _PD_Grade &&
                                  x.Rating == dataModel.Rating);
                    if (query == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.delete_Fail.GetDescription();
                        return result;
                    }

                    query.Change_Status = "D"; //刪除 D
                    query.Editor = _UserInfo._user;
                    query.Processing_Date = _UserInfo._date;
                    query.Auditor = dataModel.Auditor;
                    query.Status = null;
                    query.Auditor_Reply = null;
                    query.Audit_date = null;
                    query.LastUpdate_User = _UserInfo._user;
                    query.LastUpdate_Date = _UserInfo._date;
                    query.LastUpdate_Time = _UserInfo._time;

                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "刪除註記成功等待複核";
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }

        #endregion deleteA52

        #endregion Save Db

        #region Excel 部分

        #region 下載 Excel

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(A59)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <param name="flag">Data Flag</param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data, bool flag = false)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            var rt = Rating_Type.A.GetDescription();
            if (Excel_DownloadName.A59.ToString().Equals(type))
            {
                List<A59ViewModel> A59Data =
                    (from q in data.Cast<A58ViewModel>()
                     group q by new
                     {
                         q.Reference_Nbr,
                         q.Report_Date,
                         q.Version,
                         q.Bond_Number,
                         q.Lots,
                         q.Origination_Date,
                         q.Portfolio_Name,
                         q.Portfolio,
                         q.SMF,
                         q.Issuer,
                         q.Security_Ticker,
                         q.RATING_AS_OF_DATE_OVERRIDE,
                         q.Rating_Type,
                         q.ISIN_Changed_Ind,
                         q.Bond_Number_Old,
                         q.Origination_Date_Old
                     } into item
                     select new A59ViewModel()
                     {
                         Reference_Nbr = item.Key.Reference_Nbr,
                         Report_Date = item.Key.Report_Date,
                         Version = item.Key.Version,
                         Bond_Number = item.Key.Bond_Number,
                         Lots = item.Key.Lots,
                         Origination_Date =
                         item.Key.Rating_Type == rt ?
                         item.Key.ISIN_Changed_Ind == "Y" ?
                         item.Key.Origination_Date_Old : item.Key.Origination_Date : item.Key.Origination_Date,
                         Portfolio_Name = item.Key.Portfolio_Name,
                         Portfolio = item.Key.Portfolio,
                         SMF = item.Key.SMF,
                         Issuer = item.Key.Issuer,
                         Security_Ticker = item.Key.Security_Ticker,
                         RATING_AS_OF_DATE_OVERRIDE = item.Key.RATING_AS_OF_DATE_OVERRIDE,
                         Rating_Type = item.Key.Rating_Type,
                         ISIN_Changed_Ind = item.Key.ISIN_Changed_Ind,
                         Bond_Number_Old = item.Key.Bond_Number_Old
                     }).ToList();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (A59Data.Any())
                    {
                        var first = A59Data.First();
                        DateTime reportDate = DateTime.Parse(first.Report_Date);
                        int version = Int32.Parse(first.Version);
                        string Rating_Type = first.Rating_Type;
                        var A57s = db.Bond_Rating_Info.AsNoTracking()
                            .Where(x => x.Report_Date == reportDate &&
                                        x.Version == version &&
                                        x.Rating_Type == Rating_Type &&
                                        //x.RTG_Bloomberg_Field != null &&
                                        x.Rating != null).ToList();
                        var A57ns = db.Bond_Rating_Info.AsNoTracking()
                            .Where(x => x.Report_Date == reportDate &&
                                        x.Version == version &&
                                        x.Rating_Type == Rating_Type &&
                                        //x.RTG_Bloomberg_Field == " " &&
                                        x.Rating == null).ToList();
                        var PInfos = new A59ViewModel().GetType().GetProperties();
                        A59Data.ForEach(x =>
                        {
                            var A57Data = A57s.Where(y =>
                            y.Reference_Nbr == x.Reference_Nbr).ToList();
                            if (A57Data.Any())
                            {
                                var A57f = A57Data.First();
                                x.ISSUER_EQUITY_TICKER = A57f.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57f.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57f.GUARANTOR_EQY_TICKER;
                                x.Fill_up_YN = A57Data.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                                if (flag)
                                {
                                    A57Data.Aggregate(x,
                                    (A59, A57) =>
                                    {
                                        var PInfo = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                       A57.RTG_Bloomberg_Field.Trim().ToUpper())
                                                          .FirstOrDefault();
                                        if (PInfo != null)
                                        {
                                            PInfo.SetValue(A59, A57.Rating);
                                            var rdt = TypeTransfer.dateTimeNToString(A57.Rating_Date);
                                            if (!rdt.IsNullOrWhiteSpace())
                                            {
                                                var DtFildes = PInfo.GetCustomAttributes(typeof(SetDateFiledAttribute), false);
                                                if (DtFildes.Length > 0)
                                                {
                                                    var DtFilde = ((SetDateFiledAttribute)DtFildes[0]).Description;
                                                    var PInfo2 = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                                    DtFilde.ToUpper()).FirstOrDefault();
                                                    if (PInfo2 != null)
                                                    {
                                                        PInfo2.SetValue(A59, rdt);
                                                    }
                                                }
                                            }
                                        }
                                        return A59;
                                    });
                                }
                            }
                            else
                            {
                                var _A57ns = A57ns.Where(y =>
                                        y.Reference_Nbr == x.Reference_Nbr).ToList();
                                var A57n = _A57ns.FirstOrDefault();
                                x.ISSUER_EQUITY_TICKER = A57n?.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57n?.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57n?.GUARANTOR_EQY_TICKER;
                                x.Fill_up_YN = _A57ns.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                            }
                        });
                    }
                }

                DataTable dt = A59Data.ToDataTable();
                result.DESCRIPTION = FileRelated.DataTableToExcel(dt, path, Excel_DownloadName.A59);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (Excel_DownloadName.A59b.ToString().Equals(type))
            {
                List<A59bViewModel> A59bData = new List<A59bViewModel>();
                var _Rating_Selection_1 = "孰高";
                var _Rating_Selection_2 = "孰低";
                try
                {
                    //var D60s = getParmIDs();

                    foreach (var item in data.Cast<A58ViewModel>().GroupBy(x => new
                    {
                        x.Reference_Nbr,
                        x.Report_Date,
                        x.Version,
                        x.Bond_Number,
                        x.Lots,
                        x.Origination_Date,
                        x.Portfolio_Name,
                        x.Portfolio,
                        x.SMF,
                        x.Issuer,
                        x.Security_Ticker,
                        x.RATING_AS_OF_DATE_OVERRIDE,
                        x.Rating_Type,
                        x.ISIN_Changed_Ind,
                        x.Bond_Number_Old,
                        x.Origination_Date_Old
                    }))
                    {
                        //var _Rating_Selection = item.First().Rating_Selection;
                        //var _Rating_Selection_bonds = _Rating_Selection;
                        //var _Rating_Selection_issuer = _Rating_Selection;
                        //var _Rating_Selection_guarantor = _Rating_Selection;
                        var Listbonds = item.Where(x => x.Rating_Object == RatingObject.Bonds.GetDescription()).ToList();
                        var Listissuer = item.Where(x => x.Rating_Object == RatingObject.ISSUER.GetDescription()).ToList();
                        var Listguarantor = item.Where(x => x.Rating_Object == RatingObject.GUARANTOR.GetDescription()).ToList();
                        var _Rating_Selection_bonds = Listbonds.FirstOrDefault()?.Rating_Selection;
                        var _Rating_Selection_issuer = Listissuer.FirstOrDefault()?.Rating_Selection;
                        var _Rating_Selection_guarantor = Listguarantor.FirstOrDefault()?.Rating_Selection;
                        var _Rating_Priority_bonds = Listbonds.Any() ? Listbonds.Min(x => TypeTransfer.stringToInt(x.Rating_Priority)) : 0;
                        var _Rating_Priority_issuer = Listissuer.Any() ? Listissuer.Min(x => TypeTransfer.stringToInt(x.Rating_Priority)) : 0;
                        var _Rating_Priority_guarantor = Listguarantor.Any() ? Listguarantor.Min(x => TypeTransfer.stringToInt(x.Rating_Priority)) : 0;
                        Dictionary<string, int> Rating_Priority = new Dictionary<string, int>();
                        if (_Rating_Priority_bonds != 0)
                            Rating_Priority.Add(RatingObject.Bonds.GetDescription(), _Rating_Priority_bonds);
                        if (_Rating_Priority_issuer != 0)
                            Rating_Priority.Add(RatingObject.ISSUER.GetDescription(), _Rating_Priority_issuer);
                        if (_Rating_Priority_guarantor != 0)
                            Rating_Priority.Add(RatingObject.GUARANTOR.GetDescription(), _Rating_Priority_guarantor);
                        Rating_Priority = Rating_Priority
                            .OrderBy(Data => Data.Value)
                            .ToDictionary(keyvalue => keyvalue.Key, keyvalue => keyvalue.Value);
                        var bonds = item.Where(x => x.Rating_Object == RatingObject.Bonds.GetDescription());
                        var _bonds = bonds.Any(x=> TypeTransfer.stringToInt(x.Grade_Adjust) != 0) ?
                                     _Rating_Selection_bonds == _Rating_Selection_1 ? bonds.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Min(x => x) :
                                     _Rating_Selection_bonds == _Rating_Selection_2 ? bonds.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Max(x => x) : 0 : 0;
                        var issuer = item.Where(x => x.Rating_Object == RatingObject.ISSUER.GetDescription());
                        var _issuer = issuer.Any(x => TypeTransfer.stringToInt(x.Grade_Adjust) != 0) ?
                                      _Rating_Selection_issuer == _Rating_Selection_1 ? issuer.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Min(x => x) :
                                      _Rating_Selection_issuer == _Rating_Selection_2 ? issuer.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Max(x => x) : 0 : 0;
                        var guarantor = item.Where(x => x.Rating_Object == RatingObject.GUARANTOR.GetDescription());
                        var _guarantor = guarantor.Any(x => TypeTransfer.stringToInt(x.Grade_Adjust) != 0) ?
                                         _Rating_Selection_guarantor == _Rating_Selection_1 ? guarantor.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Min(x => x) :
                                         _Rating_Selection_guarantor == _Rating_Selection_2 ? guarantor.Select(x => TypeTransfer.stringToInt(x.Grade_Adjust)).Where(x => x != 0).Max(x => x) : 0 : 0;
                        string _all = null;
                        foreach (var i in Rating_Priority)
                        {
                            if (_all == null && i.Key == RatingObject.Bonds.GetDescription() && _bonds != 0)
                            {
                                _all = _bonds.ToString();
                            }
                            else if (_all == null && i.Key == RatingObject.ISSUER.GetDescription() && _issuer != 0)
                            {
                                _all = _issuer.ToString();
                            }
                            else if (_all == null && i.Key == RatingObject.GUARANTOR.GetDescription() && _guarantor != 0)
                            {
                                _all = _guarantor.ToString();
                            }
                        }
                        //var _all = _bonds != 0 ? _bonds.ToString() :
                        //           _issuer != 0 ? _issuer.ToString() :
                        //          _guarantor != 0 ? _guarantor.ToString() : null;
                        A59bData.Add(new A59bViewModel()
                        {
                            Reference_Nbr = item.Key.Reference_Nbr,
                            Report_Date = item.Key.Report_Date,
                            Version = item.Key.Version,
                            Bond_Number = item.Key.Bond_Number,
                            Lots = item.Key.Lots,
                            Origination_Date =
                            item.Key.Rating_Type == rt ?
                            item.Key.ISIN_Changed_Ind == "Y" ?
                            item.Key.Origination_Date_Old : item.Key.Origination_Date : item.Key.Origination_Date,
                            Portfolio_Name = item.Key.Portfolio_Name,
                            SMF = item.Key.SMF,
                            Issuer = item.Key.Issuer,
                            Security_Ticker = item.Key.Security_Ticker,
                            RATING_AS_OF_DATE_OVERRIDE = item.Key.RATING_AS_OF_DATE_OVERRIDE,
                            Rating_Type = item.Key.Rating_Type,
                            Bonds_Rating = _bonds != 0 ? _bonds.ToString() : null,
                            ISSUER_Rating = _issuer != 0 ? _issuer.ToString() : null,
                            GUARANTOR_Rating = _guarantor != 0 ? _guarantor.ToString() : null,
                            All_Rating = _all
                        });
                    }
                }
                catch (Exception ex)
                {
                    ex.exceptionMessage();
                }
                if (A59bData.Any())
                {
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        var first = A59bData.First();
                        DateTime reportDate = DateTime.Parse(first.Report_Date);
                        int version = Int32.Parse(first.Version);
                        string Rating_Type = first.Rating_Type;
                        var A57s = db.Bond_Rating_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == reportDate &&
                                                x.Version == version &&
                                                x.Rating_Type == Rating_Type &&
                                                //x.RTG_Bloomberg_Field != null &&
                                                x.Rating != null).ToList();
                        var A57ns = db.Bond_Rating_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == reportDate &&
                                                x.Version == version &&
                                                x.Rating_Type == Rating_Type &&
                                                //x.RTG_Bloomberg_Field == " " &&
                                                x.Rating == null).ToList();
                        var PInfos = new A59bViewModel().GetType().GetProperties();
                        A59bData.ForEach(x =>
                        {
                            //int _Bonds_Rating = x.Bonds_Rating == null ? 0 : TypeTransfer.stringToInt(x.Bonds_Rating);
                            //int _ISSUER_Rating = x.ISSUER_Rating == null ? 0 : TypeTransfer.stringToInt(x.ISSUER_Rating);
                            //int _GUARANTOR_Rating = x.GUARANTOR_Rating == null ? 0 : TypeTransfer.stringToInt(x.GUARANTOR_Rating);
                            var A57Data = A57s.Where(y =>
                            y.Reference_Nbr == x.Reference_Nbr).ToList();
                            if (A57Data.Any())
                            {
                                var A57f = A57Data.First();
                                x.ISSUER_EQUITY_TICKER = A57f.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57f.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57f.GUARANTOR_EQY_TICKER;
                                //var _Fill_up_YN = A57Data
                                //.Where(z =>
                                //z.Rating_Object == RatingObject.Bonds.GetDescription() &&
                                //z.Grade_Adjust != null &&
                                //z.Grade_Adjust.Value == _Bonds_Rating,
                                //_Bonds_Rating != 0)
                                //.Where(z =>
                                //z.Rating_Object == RatingObject.ISSUER.GetDescription() &&
                                //z.Grade_Adjust != null &&
                                //z.Grade_Adjust.Value == _ISSUER_Rating,
                                //_Bonds_Rating == 0 && _ISSUER_Rating != 0)
                                //.Where(z =>
                                //z.Rating_Object == RatingObject.GUARANTOR.GetDescription() &&
                                //z.Grade_Adjust != null &&
                                //z.Grade_Adjust.Value == _GUARANTOR_Rating,
                                //_Bonds_Rating == 0 && _ISSUER_Rating == 0 && _GUARANTOR_Rating != 0)
                                //.FirstOrDefault()?.Fill_up_YN;
                                var _Fill_up_YN = A57Data.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                                x.Fill_up_YN = _Fill_up_YN;
                                if (flag)
                                {
                                    A57Data.Aggregate(x,
                                    (A59b, A57) =>
                                    {
                                        var PInfo = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                       A57.RTG_Bloomberg_Field.Trim().ToUpper())
                                                          .FirstOrDefault();
                                        if (PInfo != null)
                                        {
                                            PInfo.SetValue(A59b, A57.Rating);
                                            var rdt = TypeTransfer.dateTimeNToString(A57.Rating_Date);
                                            if (!rdt.IsNullOrWhiteSpace())
                                            {
                                                var DtFildes = PInfo.GetCustomAttributes(typeof(SetDateFiledAttribute), false);
                                                if (DtFildes.Length > 0)
                                                {
                                                    var DtFilde = ((SetDateFiledAttribute)DtFildes[0]).Description;
                                                    var PInfo2 = PInfos.Where(t => t.Name.Trim().ToUpper() ==
                                                                    DtFilde.ToUpper()).FirstOrDefault();
                                                    if (PInfo2 != null)
                                                    {
                                                        PInfo2.SetValue(A59b, rdt);
                                                    }
                                                }
                                            }
                                        }
                                        return A59b;
                                    });
                                }
                            }
                            else
                            {
                                var _A57ns = A57ns.Where(y =>
                                        y.Reference_Nbr == x.Reference_Nbr).ToList();
                                var A57n = _A57ns.FirstOrDefault();
                                x.ISSUER_EQUITY_TICKER = A57n?.ISSUER_TICKER;
                                x.GUARANTOR_NAME = A57n?.GUARANTOR_NAME;
                                x.GUARANTOR_EQY_TICKER = A57n?.GUARANTOR_EQY_TICKER;
                                x.Fill_up_YN = _A57ns.Any(z => z.Fill_up_YN == "Y") ? "Y" : string.Empty;
                            }
                        });
                    }
                }
                DataTable dt = A59bData.ToDataTable();
                List<FormateTitle> ft = new List<FormateTitle>()
                {
                    new FormateTitle() {
                        OldTitle = "Bonds_Rating",
                        NewTitle = "債項最終評等"
                    },
                    new FormateTitle() {
                        OldTitle = "ISSUER_Rating",
                        NewTitle = "發行者最終評等"
                    },
                    new FormateTitle() {
                        OldTitle = "GUARANTOR_Rating",
                        NewTitle = "擔保者最終評等"
                    },
                    new FormateTitle() {
                        OldTitle = "All_Rating",
                        NewTitle = "最終評等"
                    }
                };
                result.DESCRIPTION = FileRelated.DataTableToExcel(dt, path, Excel_DownloadName.A59b, ft);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (Excel_DownloadName.A57.ToString().Equals(type))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<A57ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.A57);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (Excel_DownloadName.A56.ToString().Equals(type))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<A56ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.A56,new A56ViewModel().GetFormateTitles());
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (Excel_DownloadName.A52.ToString().Equals(type) || (Excel_DownloadName.A52Audit.ToString().Equals(type)))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<A52ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.A52, new A52ViewModel().GetFormateTitles());
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }

        #endregion 下載 Excel

        #region Excel 資料轉成 A59ViewModel

        /// <summary>
        /// Excel 資料轉成 A59ViewModel
        /// </summary>
        /// <param name="pathType">Excel 副檔名</param>
        /// <param name="path">檔案路徑</param>
        /// <param name="action"></param>
        /// <returns></returns>
        public Tuple<string,List<A59ViewModel>> getA59Excel(string pathType, string path, string action)
        {
            List<A59ViewModel> dataModel = new List<A59ViewModel>();
            IWorkbook wb = null;
            string msg = string.Empty;
            try
            {
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
                    List<string> titles = dt.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName.Trim())
                                 .ToList();

                    if (!titles.Contains("Bond_Number_Old"))
                        msg = "上傳A59檔案需使用(下載A59Excel[非熟高,熟低]所下載之檔案)";
                    dataModel = dt.AsEnumerable()
                        .Where(x => !TypeTransfer.objToString(x[0]).IsNullOrWhiteSpace())
                        .Select((x, y) =>
                        {
                            return getA59ViewModelInExcel(x, titles);
                        }
                        ).ToList();
                }
            }
            catch (Exception ex)
            { }
            return new Tuple<string, List<A59ViewModel>>(msg,dataModel);
        }

        #endregion Excel 資料轉成 A59ViewModel

        #endregion Excel 部分

        #region Private Function

        #region 組成 A51ViewModel

        private A51ViewModel getA51ViewModel(Grade_Moody_Info item)
        {
            A51ViewModel data = new A51ViewModel();
            data.Data_Year = item.Data_Year;
            data.Rating = item.Rating;
            data.PD_Grade = item.PD_Grade?.ToString();
            data.Rating_Adjust = item.Rating_Adjust;
            data.Grade_Adjust = item.Grade_Adjust?.ToString();
            data.Moodys_PD = item.Moodys_PD?.ToString();
            data.Status = A51Status(item.Status);
            data.Auditor = common.getUserAccountAndName(users,item.Auditor);
            data.Auditor_Reply = item.Auditor_Reply;
            data.Audit_date = item.Audit_date?.ToString("yyyy/MM/dd");
            return data;
        }

        #endregion 組成 A51ViewModel

        #region 組成 A52ViewModel

        private A52ViewModel getA52ViewModel(Grade_Mapping_Info item)
        {
            return new A52ViewModel()
            {
                Id = item.Id.ToString(),
                Rating_Org = item.Rating_Org,
                PD_Grade = item.PD_Grade.ToString(),
                Rating = item.Rating,
                IsActive = item.IsActive,
                Change_Status = item.Change_Status == "I" ? "新增:I" : item.Change_Status == "D" ? "刪除:D" : "",
                Editor = common.getUserAccountAndName(users, item.Editor),
                Processing_Date = item.Processing_Date?.ToString("yyyy/MM/dd"),
                Auditor = common.getUserAccountAndName(users, item.Auditor),
                Status = item.Status == "1" ? "接受" : item.Status == "0" ? "退回" : "待複核",
                Auditor_Reply = item.Auditor_Reply,
                Audit_date = item.Audit_date?.ToString("yyyy/MM/dd")
            };
        }

        #endregion 組成 A52ViewModel

        public string A51Status(string status)
        {
            if (status == "1")
                return Audit_Type.Enable.GetDescription();
            if (status == "2")
                return Audit_Type.TempDisabled.GetDescription();
            if (status == "3")
                return Audit_Type.Disabled.GetDescription();
            return Audit_Type.None.GetDescription();          
        }

        #region Bond_Rating_Summary 組成 A58ViewModel

        /// <summary>
        /// Bond_Rating_Summary 組成 A58ViewModel
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>A58ViewModel</returns>
        private A58ViewModel getA58ViewModel(Bond_Rating_Summary item)
        {
            return new A58ViewModel()
            {
                Reference_Nbr = item.Reference_Nbr,
                Report_Date = TypeTransfer.dateTimeNToString(item.Report_Date),
                Bond_Number = item.Bond_Number,
                Lots = item.Lots,
                Origination_Date = TypeTransfer.dateTimeNToString(item.Origination_Date),
                Parm_ID = item.Parm_ID,
                Bond_Type = item.Bond_Type,
                Rating_Type = item.Rating_Type,
                Rating_Object = item.Rating_Object,
                Rating_Org_Area = item.Rating_Org_Area,
                Rating_Selection = item.Rating_Selection == "1" ? "孰高" : item.Rating_Selection == "2" ? "孰低" : " ",
                Grade_Adjust = TypeTransfer.intNToString(item.Grade_Adjust),
                Rating_Priority = TypeTransfer.intNToString(item.Rating_Priority),
                Processing_Date = TypeTransfer.dateTimeNToString(item.Processing_Date),
                Version = TypeTransfer.intNToString(item.Version),
                Portfolio_Name = item.Portfolio_Name,
                Portfolio = item.Portfolio,
                SMF = item.SMF,
                Issuer = item.ISSUER,
                Security_Ticker = getSecurityTicker(item.SMF, item.Bond_Number),
                ISIN_Changed_Ind = item.ISIN_Changed_Ind,
                Origination_Date_Old = TypeTransfer.dateTimeNToString(item.Origination_Date_Old),
                Bond_Number_Old = item.Bond_Number_Old,
                RATING_AS_OF_DATE_OVERRIDE = item.Rating_Type == Rating_Type.A.GetDescription() ?
                TypeTransfer.dateTimeNToString(item.Origination_Date, 8) :
                TypeTransfer.dateTimeNToString(item.Report_Date, 8)
            };
        }

        #endregion Bond_Rating_Summary 組成 A58ViewModel

        #region Excel 組成 A59ViewModel

        private A59ViewModel getA59ViewModelInExcel(DataRow item, List<string> titles)
        {
            var A59 = new A59ViewModel();
            if (!titles.Any())
                return A59;
            for (int i = 0; i < titles.Count; i++) //每一行所有資料
            {
                string data = TypeTransfer.objToString(item[i]);
                if (!data.IsNullOrWhiteSpace()) //資料有值
                {
                    var A59PInfo = A59.GetType().GetProperties()
                                      .Where(x => x.Name.Trim().ToLower() == titles[i].Trim().ToLower())
                                      .FirstOrDefault();
                    if (A59PInfo != null)
                    {
                        if (A59PInfo.Name.ToUpper().EndsWith("DT") || A59PInfo.Name.ToUpper().EndsWith("DATE"))
                            A59PInfo.SetValue(A59, excelDateToString(data));
                        else
                            if(A59PInfo.Name == "Reference_Nbr")
                                A59PInfo.SetValue(A59, data.PadLeft(10, '0'));
                            else
                                A59PInfo.SetValue(A59, data);
                    }
                }
            }
            return A59;
        }

        #endregion Excel 組成 A59ViewModel

        #region Db 組成 A57ViewModel

        /// <summary>
        /// Db 組成 A57ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private A57ViewModel DbToA57ViewModel(Bond_Rating_Info data)
        {
            return new A57ViewModel()
            {
                Reference_Nbr = data.Reference_Nbr,
                Bond_Number = data.Bond_Number,
                Lots = data.Lots,
                Portfolio = data.Portfolio,
                Segment_Name = data.Segment_Name,
                Bond_Type = data.Bond_Type,
                Lien_position = data.Lien_position,
                Origination_Date = TypeTransfer.dateTimeNToString(data.Origination_Date),
                Report_Date = TypeTransfer.dateTimeNToString(data.Report_Date),
                Rating_Date = TypeTransfer.dateTimeNToString(data.Rating_Date),
                Rating_Type = data.Rating_Type,
                Rating_Object = data.Rating_Object,
                Rating_Org = data.Rating_Org,
                Rating = data.Rating,
                Rating_Org_Area = data.Rating_Org_Area,
                Fill_up_YN = data.Fill_up_YN,
                Fill_up_Date = TypeTransfer.dateTimeNToString(data.Fill_up_Date),
                PD_Grade = TypeTransfer.intNToString(data.PD_Grade),
                Grade_Adjust = TypeTransfer.intNToString(data.Grade_Adjust),
                ISSUER_TICKER = data.ISSUER_TICKER,
                GUARANTOR_NAME = data.GUARANTOR_NAME,
                GUARANTOR_EQY_TICKER = data.GUARANTOR_EQY_TICKER,
                Parm_ID = data.Parm_ID,
                Portfolio_Name = data.Portfolio_Name,
                RTG_Bloomberg_Field = data.RTG_Bloomberg_Field,
                SMF = data.SMF,
                ISSUER = data.ISSUER,
                Version = TypeTransfer.intNToString(data.Version),
                Security_Ticker = data.Security_Ticker,
                ISIN_Changed_Ind = data.ISIN_Changed_Ind,
                Bond_Number_Old = data.Bond_Number_Old,
                Lots_Old = data.Lots_Old,
                Portfolio_Name_Old = data.Portfolio_Name_Old,
                Origination_Date_Old = TypeTransfer.dateTimeNToString(data.Origination_Date_Old)
            };
        }

        #endregion Db 組成 A57ViewModel

        /// <summary>
        /// get Security_Ticker
        /// </summary>
        /// <param name="SMF"></param>
        /// <param name="bondNumber"></param>
        /// <returns></returns>
        private string getSecurityTicker(string SMF, string bondNumber)
        {
            List<string> Mtges = new List<string>() { "A11", "932" };
            if (!SMF.IsNullOrWhiteSpace() && SMF.Trim().Length > 3)
                if (Mtges.Contains(SMF.Substring(0, 3)))
                    return string.Format("{0} {1}", bondNumber, "Mtge");
            return string.Format("{0} {1}", bondNumber, "Corp");
        }

        /// <summary>
        /// 修正 rating
        /// </summary>
        /// <param name="rating"></param>
        /// <returns></returns>
        private string forRating(string rating, string Org)
        {
            if (rating.IsNullOrWhiteSpace())
                return string.Empty;
            string value = rating.Trim();
            if ((Org != RatingOrg.Moody.GetDescription()) && (value.IndexOf("WR") > -1))
                return string.Empty;
            rule_2s.ForEach(x =>
            {
                if (value.Contains(x.Item1) && !x.Item2.IsNullOrWhiteSpace())
                {
                    value = value.Replace(x.Item1, x.Item2);
                    if (!value.IsNullOrWhiteSpace())
                        value = value.Trim();
                }
            });
            List<string> Liststrs = new List<string>();
            rule_1s.ForEach(x =>
            {
                if (value.Contains(x))
                {
                    //value = SplitFirst(value, x);
                    //value = value.Replace(x,string.Empty);
                    Liststrs.Add(x);
                }
            });
            if (Liststrs.Count > 0)
            {
                string[] strs = Liststrs.ToArray();
                ExchangeSort(strs);
                strs.Reverse().ToList().ForEach(x =>
                {
                    value = value.Replace(x, string.Empty);
                    if (!value.IsNullOrWhiteSpace())
                        value = value.Trim();
                });
            }
            return value;
        }

        //Exchange Sort(交換排序法)
        public static void ExchangeSort(string[] list)
        {
            int n = list.Length;
            string temp;
            for (int i = 0; i <= n - 1; i++)
            {
                for (int j = i + 1; j < n; j++)
                {
                    if (list[i].Contains(list[j]))
                    {  // 比較鄰近兩個物件，左邊包含右邊時就互換。	       
                        temp = list[j];
                        list[j] = list[i];
                        list[i] = temp;
                    }
                }
            }
        }

        /// <summary>
        /// get ParmID
        /// </summary>
        /// <returns></returns>
        private List<Bond_Rating_Parm> getParmIDs()
        {
            List<Bond_Rating_Parm> parmIDs = new List<Bond_Rating_Parm>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var parms = db.Bond_Rating_Parm.AsNoTracking()
                    .Where(j => j.IsActive == "Y").ToList(); //抓取所有有效資料
                foreach (var item in parms.GroupBy(x => new { x.Rating_Object ,x.Rating_Org_Area}))
                {
                    if (item.Any(y=>y.Rating_Org_Area == null) && item.Count() >= 2) //國內/外類別 兩種設定(1.單筆null,2.一筆國內一筆國外)
                        return parmIDs;
                }
                parmIDs = parms;
            }
            return parmIDs;
        }

        private string checkParmID(List<Bond_Rating_Parm> parms)
        {
            string result = string.Empty;
            if (!parms.Any())
            {
                result = "無有效的D60";
            }
            else if (parms.Where(x => !x.Rating_Object.IsNullOrWhiteSpace()).GroupBy(x => x.Rating_Object).Count() != 3)
            {
                result = "債項、發行人、保證人資料缺漏,請至D60維護";
            }
            else if (parms.Where(x => !x.Rating_Selection.IsNullOrWhiteSpace()).GroupBy(x => x.Rating_Selection).Count() != 1)
            {
                result = "孰高孰低資料維護不一致,請至D60維護";
            }
            return result;
        }

        /// <summary>
        /// get ParmID
        /// </summary>
        /// <returns></returns>
        private Bond_Rating_Parm getParmID(List<Bond_Rating_Parm> datas, string Rating_Object, string Rating_Org_Area)
        {
            Bond_Rating_Parm _parm = datas.FirstOrDefault(m => m.Rating_Org_Area != null && m.Rating_Object == Rating_Object && m.Rating_Org_Area == Rating_Org_Area);
            if (_parm == null)
                _parm = datas.FirstOrDefault(m => m.Rating_Object == Rating_Object);
            return _parm;
        }

        /// <summary>
        /// 抓取 sampleInfo
        /// </summary>
        /// <param name="sampleInfos"></param>
        /// <param name="info"></param>
        /// <param name="nullarr"></param>
        /// <returns></returns>
        private Rating_Info_SampleInfo formateSampleInfo(
            List<Rating_Info_SampleInfo> sampleInfos,
            Rating_Info info,
            List<string> nullarr)
        {
            if (RatingObject.Bonds.GetDescription().Equals(info.Rating_Object))
            {
                Rating_Info_SampleInfo s = sampleInfos.FirstOrDefault(j => info.Bond_Number.Equals(j.Bond_Number));
                if (s != null)
                {
                    return new Rating_Info_SampleInfo()
                    {
                        Bond_Number = s.Bond_Number,
                        ISSUER_TICKER = sampleInfoValue(s.ISSUER_TICKER, nullarr),
                        GUARANTOR_EQY_TICKER = sampleInfoValue(s.GUARANTOR_EQY_TICKER, nullarr),
                        GUARANTOR_NAME = sampleInfoValue(s.GUARANTOR_NAME, nullarr)
                    };
                }
            }
            return new Rating_Info_SampleInfo();
        }

        /// <summary>
        /// 判斷sampleInfo 參數
        /// </summary>
        /// <param name="value"></param>
        /// <param name="nullarr"></param>
        /// <returns></returns>
        private string sampleInfoValue(string value, List<string> nullarr)
        {
            if (value == null)
                return null;
            if (nullarr.Contains(value.Trim()))
                return null;
            return value.Trim();
        }

        private string excelDateToString(string val)
        {
            if (val.IsNullOrWhiteSpace())
                return null;
            DateTime dt = DateTime.MinValue;
            if (DateTime.TryParse(val, out dt))
                return dt.ToString("yyyy/MM/dd");
            double d = 0d;
            if (!double.TryParse(val, out d))
                return null;
            return DateTime.FromOADate(d).ToString("yyyy/MM/dd");
        }

        private void addData(
            string name,
            string val,
            string dt,
            string Object,
            string Org_Area,
            string Org,
            List<A59Data> datas
            )
        {
            string v = forRating(val, Org);
            if (!v.IsNullOrWhiteSpace())
                datas.Add(new A59Data()
                {
                    RTG_Bloomberg_Field = name,
                    Rating = v,
                    Rating_Date = dt.stringToDateTimeStr(),
                    Rating_Object = Object,
                    Rating_Org_Area = Org_Area,
                    Rating_Org = Org
                });
            else
            {
                datas.Add(new A59Data()
                {
                    RTG_Bloomberg_Field = name,
                    Rating = string.Empty,
                    Rating_Date = dt.stringToDateTimeStr(),
                    Rating_Object = Object,
                    Rating_Org_Area = Org_Area,
                    Rating_Org = Org
                });
            }
        }

        private string SplitFirst(string value, string splitStr)
        {
            if (value.IsNullOrWhiteSpace())
                return value;
            return value.Split(new string[] { splitStr }, StringSplitOptions.RemoveEmptyEntries)[0].Trim();
        }

        private string insertA57(Bond_Rating_Info A57)
        {
            return $@"
INSERT INTO [Bond_Rating_Info]
           ([Reference_Nbr]
           ,[Bond_Number]
           ,[Lots]
           ,[Portfolio]
           ,[Segment_Name]
           ,[Bond_Type]
           ,[Lien_position]
           ,[Origination_Date]
           ,[Report_Date]
           ,[Rating_Date]
           ,[Rating_Type]
           ,[Rating_Object]
           ,[Rating_Org]
           ,[Rating]
           ,[Rating_Org_Area]
           ,[Fill_up_YN]
           ,[Fill_up_Date]
           ,[PD_Grade]
           ,[Grade_Adjust]
           ,[ISSUER_TICKER]
           ,[GUARANTOR_NAME]
           ,[GUARANTOR_EQY_TICKER]
           ,[Parm_ID]
           ,[Portfolio_Name]
           ,[RTG_Bloomberg_Field]
           ,[SMF]
           ,[ISSUER]
           ,[Version]
           ,[Security_Ticker]
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
           ({A57.Reference_Nbr.stringToStrSql()}
           ,{A57.Bond_Number.stringToStrSql()}
           ,{A57.Lots.stringToStrSql()}
           ,{A57.Portfolio.stringToStrSql()}
           ,{A57.Segment_Name.stringToStrSql()}
           ,{A57.Bond_Type.stringToStrSql()}
           ,{A57.Lien_position.stringToStrSql()}
           ,{A57.Origination_Date.dateTimeNToStrSql()}
           ,{A57.Report_Date.dateTimeToStrSql()}
           ,{A57.Rating_Date.dateTimeNToStrSql()}
           ,{A57.Rating_Type.stringToStrSql()}
           ,{A57.Rating_Object.stringToStrSql()}
           ,{A57.Rating_Org.stringToStrSql()}
           ,{A57.Rating.stringToStrSql()}
           ,{A57.Rating_Org_Area.stringToStrSql()}
           ,{A57.Fill_up_YN.stringToStrSql()}
           ,{A57.Fill_up_Date.dateTimeNToStrSql()}
           ,{A57.PD_Grade.intNToStrSql()}
           ,{A57.Grade_Adjust.intNToStrSql()}
           ,{A57.ISSUER_TICKER.stringToStrSql()}
           ,{A57.GUARANTOR_NAME.stringToStrSql()}
           ,{A57.GUARANTOR_EQY_TICKER.stringToStrSql()}
           ,{A57.Parm_ID.stringToStrSql()}
           ,{A57.Portfolio_Name.stringToStrSql()}
           ,{A57.RTG_Bloomberg_Field.stringToStrSql()}
           ,{A57.SMF.stringToStrSql()}
           ,{A57.ISSUER.stringToStrSql()}
           ,{A57.Version.intNToStrSql()}
           ,{A57.Security_Ticker.stringToStrSql()}
           ,{A57.ISIN_Changed_Ind.stringToStrSql()}
           ,{A57.Bond_Number_Old.stringToStrSql()}
           ,{A57.Lots_Old.stringToStrSql()}
           ,{A57.Portfolio_Name_Old.stringToStrSql()}
           ,{A57.Origination_Date_Old.dateTimeNToStrSql()} 
           ,{_UserInfo._user.stringToStrSql()}
           ,{_UserInfo._date.dateTimeToStrSql()}
           ,{_UserInfo._time.timeSpanToStrSql()} ); ";
        }

        private void insertA57Rating(
            List<insertRating> _data,
            Bond_Rating_Info _first,
            StringBuilder sb,
            RatingObject _Rating_Object,
            List<Grade_Mapping_Info> A52s,
            List<Grade_Moody_Info> A51s,
            List<Bond_Rating_Parm> parmIDs)
        {
            foreach (var item in _data)
            {
                Bond_Rating_Info _copy = _first.ModelConvert<Bond_Rating_Info, Bond_Rating_Info>();
                Grade_Moody_Info A51 = null;
                if (item._Rating_Org == RatingOrg.Moody)
                {
                    A51 = A51s.FirstOrDefault(z => z.Rating == item._Rating);
                }
                else
                {
                    var A52 = A52s.FirstOrDefault(z => z.Rating_Org == item._Rating_Org.GetDescription() &&
                                                       z.Rating == item._Rating);
                    if (A52 != null)
                        A51 = A51s.FirstOrDefault(z => z.PD_Grade == A52.PD_Grade);
                }
                _copy.Rating_Object = _Rating_Object.GetDescription();
                _copy.Rating_Org_Area = item._RatingOrgArea;
                var _parm = getParmID(parmIDs, _copy.Rating_Object, _copy.Rating_Org_Area);
                _copy.Rating = item._Rating;
                _copy.Rating_Org = item._Rating_Org.GetDescription();
                _copy.RTG_Bloomberg_Field = item._RTG_Bloomberg_Field;
                _copy.Parm_ID = _parm?.Parm_ID.ToString();
                _copy.Fill_up_YN = null;
                _copy.Fill_up_Date = null;
                _copy.PD_Grade = A51?.PD_Grade;
                _copy.Grade_Adjust = A51?.Grade_Adjust;
                sb.Append(insertA57(_copy));
            }
        }

        private string A56Method_MEMO(string Update_Method, List<Rating_Update_Method> methods)
        {
            string result = string.Empty;
            int i = 0;
            Int32.TryParse(Update_Method, out i);
            methods.ForEach(x =>
            {
                if ((int)x == i)
                    result = x.GetDescription();
            });
            return result;
        }

        private class A59Data
        {
            public string RTG_Bloomberg_Field { get; set; }
            public string Rating { get; set; }
            public string Rating_Date { get; set; }
            public string Rating_Org { get; set; }
            public string Rating_Object { get; set; }
            public string Rating_Org_Area { get; set; }
        }

        private class insertRating
        {
            public RatingOrg _Rating_Org { get; set; }
            public string _Rating { get; set; }
            public string _RatingOrgArea { get; set; }
            public string _RTG_Bloomberg_Field { get; set; }
        }

        #endregion Private Function
    }
}