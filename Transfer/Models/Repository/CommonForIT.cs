using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class CommonForIT : ICommonForIT
    {
        #region save sqllog(IFRS9_Log)

        /// <summary>
        /// Log資料存到Sql(IFRS9_Log)
        /// </summary>
        /// <param name="tableType">table簡寫</param>
        /// <param name="tableName">table名</param>
        /// <param name="fileName">檔案名</param>
        /// <param name="programName">專案名</param>
        /// <param name="falg">成功失敗</param>
        /// <param name="deptType">B:債券 M:房貸 (共用同一table時做區分)</param>
        /// <param name="start">開始時間</param>
        /// <param name="end">結束時間</param>
        /// <param name="userAccount">執行帳號</param>
        /// <param name="version">版本</param>
        /// <param name="reportDate">基準日</param>
        /// <returns>回傳成功或失敗</returns>
        public bool saveLog(
            Table_Type table,
            string fileName,
            string programName,
            bool falg,
            string deptType,
            DateTime start,
            DateTime end,
            string userAccount,
            int version = 1,
            Nullable<DateTime> date = null)
        {
            bool flag = true;
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var tableName = table.GetDescription();
                    db.IFRS9_Log.Add(new IFRS9_Log() //寫入DB
                    {
                        Table_type = table.ToString(),
                        Table_name = tableName.Substring(0, (tableName.Length > 40 ? 40 : tableName.Length)),
                        File_name = fileName,
                        Program_name = programName,
                        Version = version,
                        Create_date = start.ToString("yyyyMMdd"),
                        Create_time = start.ToString("HH:mm:ss"),
                        End_date = end.ToString("yyyyMMdd"),
                        End_time = end.ToString("HH:mm:ss"),
                        TYPE = falg ? "Y" : "N",
                        Debt_Type = deptType,
                        User_Account = userAccount,
                        Report_Date = date
                    });
                    db.SaveChanges(); //DB SAVE
                }
            }
            catch
            {
                flag = false;
            }
            return flag;
        }

        #endregion save sqllog(IFRS9_Log)

        /// <summary>
        /// 判斷轉檔紀錄是否有存在
        /// </summary>
        /// <param name="fileNames">目前檔案名稱</param>
        /// <param name="checkName">要檢查的檔案名稱</param>
        /// <param name="reportDate">基準日</param>
        /// <param name="version">版本</param>
        /// <returns></returns>
        public bool checkTransferCheck(
            string fileName,
            string checkName,
            DateTime reportDate,
            int version)
        {
            if (fileName.IsNullOrWhiteSpace() || checkName.IsNullOrWhiteSpace())
                return false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var checkTable = db.Transfer_CheckTable.AsNoTracking();
                var _A53Version = 0;
                if (checkName == "A53")
                {
                    var _vers = checkTable.Where(x =>
                    x.File_Name == "A53" &&
                    x.TransferType != "R" &&
                    x.ReportDate == reportDate).ToList();
                    if (_vers.Any())
                        _A53Version = _vers.Max(x => x.Version);
                }
                //須符合有一筆"Y"(上一動作完成) 自己沒有"Y"(重複做) 才算符合
                if (//當 fileName,checkName 都為 A41 不用檢查(為最先動作)
                    ((fileName == Table_Type.A41.ToString() &&
                      checkName == Table_Type.A41.ToString()) ||
                    //檢查上一動作有無成功(A53 只有一版),前置動作檢查檔案為A53版本為A53最大版
                    checkTable.Any(x => x.ReportDate == reportDate &&
                                                    ((checkName == "A53" &&
                                                      x.Version == _A53Version) ||
                                                    (x.File_Name == checkName &&
                                                   x.Version == version)) &&
                                                   x.TransferType == "Y")) &&
                    //檢查本身有無重複執行(有成功就要下一版)
                    !checkTable.Any(x => x.File_Name == fileName &&
                                                  x.ReportDate == reportDate &&
                                                  x.Version == version &&
                                                  x.TransferType == "Y"))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// get EF connection to ADO.NET connection
        /// </summary>
        /// <param name="efConnection"></param>
        /// <returns></returns>
        public string RemoveEntityFrameworkMetadata(string efConnection)
        {
            string efstr = string.Empty;
            if (string.IsNullOrWhiteSpace(efConnection))
            {
                efstr = System.Configuration.ConfigurationManager.
                         ConnectionStrings["IFRS9Entities"].ConnectionString;
            }
            else
            {
                efstr = efConnection;
            }
            int start = efstr.IndexOf("\"", StringComparison.OrdinalIgnoreCase);
            int end = efstr.LastIndexOf("\"", StringComparison.OrdinalIgnoreCase);

            // We do not want to include the quotation marks
            start++;
            int length = end - start;

            return efstr.Substring(start, length);
        }

        /// <summary>
        /// 轉檔紀錄存到Sql(Transfer_CheckTable)
        /// </summary>
        /// <param name="fileName">檔案名稱 A41,A42...</param>
        /// <param name="flag">成功失敗</param>
        /// <param name="reportDate">基準日</param>
        /// <param name="version">版本</param>
        /// <param name="start">轉檔開始時間</param>
        /// <param name="end">轉檔結束時間</param>
        /// <returns></returns>
        public bool saveTransferCheck(
            string fileName,
            bool flag,
            DateTime reportDate,
            int version,
            DateTime start,
            DateTime end,
            string Msg = null)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (flag && db.Transfer_CheckTable.Any(x =>
                              x.ReportDate == reportDate &&
                              x.Version == version &&
                              x.File_Name == fileName &&
                              x.TransferType == "Y"))
                    return false;
                if (EnumUtil.GetValues<Transfer_Table_Type>()
                    .Select(x => x.ToString()).ToList().Contains(fileName))
                {
                    db.Transfer_CheckTable.Add(new Transfer_CheckTable()
                    {
                        File_Name = fileName,
                        ReportDate = reportDate,
                        Version = version,
                        TransferType = flag ? "Y" : "N",
                        Create_date = start.ToString("yyyyMMdd"),
                        Create_time = start.ToString("HH:mm:ss"),
                        End_date = end.ToString("yyyyMMdd"),
                        End_time = end.ToString("HH:mm:ss"),
                        Process = "Web",
                        Error_Msg = Msg
                    });
                    try
                    {
                        db.SaveChanges();
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
                return false;
            }
        }

        /// <summary>
        /// reportDate 選擇後 查詢目前有的版本
        /// </summary>
        /// <param name="date"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public List<string> getVersion(DateTime date, string tableName)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                return db.Transfer_CheckTable.AsNoTracking().Where(
                         x => x.File_Name == tableName &&
                         x.TransferType == "Y" &&
                         x.ReportDate == date)
                         .AsEnumerable().Select(x => x.Version.ToString()).ToList();
            }
        }

        /// <summary>
        /// get 複核者 or 評估者
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<IFRS9_User> getAssessmentInfo(string productCode, string tableId, SetAssessmentType type)
        {
            List<IFRS9_User> result = new List<IFRS9_User>();
            var _SystemRepository = new SystemRepository();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var set = db.IFRS9_Assessment_Config.AsNoTracking()
                    .SingleOrDefault(x => x.Group_Product_Code == productCode &&
                                          x.Table_Id == tableId &&
                                          x.Effective == "Y");
                var users = db.IFRS9_User.AsNoTracking().Where(x => x.Effective);
                if (set != null)
                {
                    if (((type & SetAssessmentType.Auditor) == SetAssessmentType.Auditor) ||
                        ((type & SetAssessmentType.Presented) == SetAssessmentType.Presented))
                    {
                        var _user = _SystemRepository.getAssessment(productCode, tableId, type)
                                    .Select(x => x.User_Account);
                        result = users.Where(x => x.Effective && _user.Contains(x.User_Account)).ToList();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// DataTable to ViewModel
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<IViewModel> getViewModel(DataTable dt, Table_Type type)
        {
            if (dt == null)
                return new List<IViewModel>();
            List<string> titles = dt.Columns
                                 .Cast<DataColumn>()
                                 .Where(x => x.ColumnName != null)
                                 .Select(x => x.ColumnName.Trim().Replace("\t", string.Empty))
                                 .ToList();
            return getViewModel(dt.AsEnumerable().Select(x => x), titles, type);
        }

        /// <summary>
        /// dataRow to ViewModel
        /// </summary>
        /// <param name="item"></param>
        /// <param name="titles"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<IViewModel> getViewModel(IEnumerable<DataRow> items, List<string> titles, Table_Type type)
        {
            List<IViewModel> datas = new List<IViewModel>();
            var obj = FactoryRegistry.GetInstance(type);
            if (obj == null)
                return datas;
            var pros = obj.GetType().GetProperties();
            if (!items.Any() || items.First().ItemArray.Count() != titles.Count)
                return datas;
            items.ToList().ForEach(item =>
            {
                obj = FactoryRegistry.GetInstance(type);
                for (int i = 0; i < titles.Count; i++) //每一行所有資料
                {
                    string data = null;
                    if (item[i].GetType().Name.Equals("DateTime"))
                        data = TypeTransfer.objDateToString(item[i]);
                    else
                        data = TypeTransfer.objToString(item[i]);
                    if (!data.IsNullOrWhiteSpace()) //資料有值
                    {
                        var PInfo = pros.Where(x => x.Name.Trim().ToLower() == titles[i].Trim().ToLower())
                                           .FirstOrDefault();
                        if (PInfo != null)
                            PInfo.SetValue(obj, data);
                    }
                }
                datas.Add(obj);
            });

            return datas;
        }

        public List<IViewModel> getViewModel<T>(IEnumerable<T> dbModel, Table_Type type)
            where T : class
        {
            List<IViewModel> datas = new List<IViewModel>();
            var obj = FactoryRegistry.GetInstance(type);
            if (dbModel == null || obj == null)
                return datas;
            var viewpros = obj.GetType().GetProperties();
            if (!dbModel.Any())
                return datas;
            var dbpros = dbModel.First().GetType().GetProperties();
            dbModel.ToList().ForEach(db => {
                obj = FactoryRegistry.GetInstance(type);
                foreach (var item in viewpros)
                {
                    var p = dbpros.FirstOrDefault(x => x.Name.ToUpper() == item.Name.ToUpper());
                    if (p != null)
                    {
                        var _pt = p.PropertyType;
                        if (_pt == typeof(string))
                        {
                            item.SetValue(obj, p.GetValue(db));
                        }
                        else if (_pt == typeof(DateTime) || _pt == typeof(Nullable<DateTime>))
                        {
                            if (p.GetValue(db) != null)
                                item.SetValue(obj, TypeTransfer.objDateToString(p.GetValue(db)));
                            else
                                item.SetValue(obj, string.Empty);
                        }
                        else
                        {
                            if (p.GetValue(db) != null)
                                item.SetValue(obj, p.GetValue(db).ToString());
                            else
                                item.SetValue(obj, string.Empty);
                        }
                    }
                }
                datas.Add(obj);
            });
            return datas;
        }

        public string getUserName(IEnumerable<IFRS9_User> users, string account)
        {
            return users.FirstOrDefault(x => x.User_Account == account)?.User_Name;
        }

        public List<IFRS9_User> getAllUsers(bool Effective = true)
        {
            List<IFRS9_User> results = new List<IFRS9_User>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                results = db.IFRS9_User.AsNoTracking().Where(x => x.Effective == Effective).ToList();
            }
            return results;
        }

        /// <summary>
        /// get user debt
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        public string getUserDebt(string account)
        {
            var result = string.Empty;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _user = db.IFRS9_User.AsNoTracking().FirstOrDefault(x => x.User_Account == account);
                if (_user != null)
                    result = _user.DebtType;
            }
            return result;
        }

        /// <summary>
        /// 查看Debt權限設定
        /// </summary>
        /// <param name="debt">A(All),B(Bonds),M(Mortgage)</param>
        /// <returns></returns>
        public List<string> getDebt(string debt)
        {
            List<string> result = new List<string>();
            switch (debt)
            {
                case "A": //A => (可以設定其他帳號(DebtType)為 A, B, M)
                    result = new List<string>() { "A", "B", "M" };
                    break;
                case "B": //B => (可以設定其他帳號(DebtType)為 B)
                    result = new List<string>() { "B" };
                    break;
                case "M": //M => (可以設定其他帳號(DebtType)為 M)
                    result = new List<string>() { "M" };
                    break;
            }
            return result;
        }

        /// <summary>
        /// get Debt SelectOption
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        public List<SelectOption> getDebtSelectOption(string account)
        {
            List<SelectOption> result = new List<SelectOption>() {
                new SelectOption() {
                    Text = " ",Value = " "
                } };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result.AddRange(getDebt("A")
                    .Select(x => debtSelectOption(x)));
            }
            return result;
        }

        private SelectOption debtSelectOption(string debt)
        {
            SelectOption result = new SelectOption();
            switch (debt)
            {
                case "A":
                    result.Value = "A";
                    result.Text = "A:(全部權限)";
                    break;
                case "B":
                    result.Value = "B";
                    result.Text = "B:(債券)";
                    break;
                case "M":
                    result.Value = "M";
                    result.Text = "M:(房貸)";
                    break;
            }
            return result;
        }

        public string getUserName()
        {
            return Controllers.AccountController.CurrentUserInfo.Name;
        }

        public DateTime getDate()
        {
            return DateTime.Now.Date;
        }

        public TimeSpan getTimeSpa()
        {
            return DateTime.Now.TimeOfDay;
        }

        public class User
        {
            public User()
            {
                _user = "";
                _date = DateTime.Now.Date;
                _time = DateTime.Now.TimeOfDay;
            }
            public string _user { get; private set; }
            public DateTime _date { get; set; }
            public TimeSpan _time { get; set; }
        }

        public string excelDateToString(string val)
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
    }
}