using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class SystemRepository : ISystemRepository
    {
        #region 其他

        public SystemRepository()
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

        #region GetData

        /// <summary>
        /// get menu in account
        /// </summary>
        /// <param name="userAccount"></param>
        /// <returns></returns>
        public List<CheckBoxListInfo> getMenu(string userAccount)
        {
            List<CheckBoxListInfo> resultData = new List<CheckBoxListInfo>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _debt = common.getDebt(common.getUserDebt(userAccount));
                _debt.Add("A");
                List<string> Menu_Ids = db.IFRS9_Menu_Set.AsNoTracking().AsEnumerable()
                        .Where(x => userAccount.Equals(x.User_Account) && x.Effective == "Y")
                        .Select(x => x.Menu_Id).ToList();
                resultData.AddRange(db.IFRS9_Menu_Sub.AsNoTracking().AsEnumerable()
                    .Where(x => !"System".Equals(x.Menu) && _debt.Contains(x.DebtType))
                    .Select(x =>
                    {
                        return new CheckBoxListInfo()
                        {
                            Value = x.Menu_Id,
                            DisplayText = x.Menu_Detail,
                            IsChecked = Menu_Ids.Contains(x.Menu_Id)
                        };
                    }));
            }

            return resultData;
        }

        /// <summary>
        /// get can modify menu 的 帳號
        /// </summary>
        /// <param name="type">menu=>設定畫面的帳號,log=>觀察log的帳號</param>
        /// <param name="searchAll">log時使用(查詢包含失效(刪除))</param>
        /// <param name="adminAccount">管理者帳號</param>
        /// <returns></returns>
        public List<Tuple<string, string>> getUser(string type,bool searchAll,string adminAccount)
        {
            List<Tuple<string, string>> result = new List<Tuple<string, string>>() { new Tuple<string, string>("", "") };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_User.Any(x => x.User_Account == adminAccount && x.AdminFlag))
                {
                    var _debt = common.getUserDebt(adminAccount);
                    var _debts = common.getDebt(_debt); //債券類別 A,M,B
                    if (type == "menu")
                    {
                        result.AddRange(
                              db.IFRS9_User.AsNoTracking().Where(x =>
                              !x.AdminFlag &&
                               x.Effective )
                               .Where(x => _debts.Contains(x.DebtType), _debt != "A")
                              .AsEnumerable()
                              .Select(x => new Tuple<string, string>(x.User_Account,
                              string.Format("{0} ({1})", x.User_Account, x.User_Name))));                           
                    }
                    if (type == "log")
                    {
                        result.AddRange(
                            db.IFRS9_User.AsNoTracking()
                            .Where(x => x.Effective, !searchAll)
                            .Where(x => _debts.Contains(x.DebtType), _debt != "A")
                            .AsEnumerable()
                            .Select(x => new Tuple<string, string>(x.User_Account,
                            string.Format("{0} ({1})", x.User_Account, x.User_Name))));
                    }
                }

            }
            return result;
        }

        /// <summary>
        /// get ProductCode
        /// </summary>
        /// <param name="adminAccount"></param>
        /// <returns></returns>
        public List<Tuple<string, string>> getProductCode(string adminAccount)
        {
            List<Tuple<string, string>> result =
                new List<Tuple<string, string>>() { new Tuple<string, string>("", "") };
            var _debt = common.getUserDebt(adminAccount);
            List<string> _Group_Product_Codes = new List<string>();
            switch (_debt)
            {
                case "A":
                    _Group_Product_Codes = new List<string>() { "001", "002" };
                    break;
                case "B":
                    _Group_Product_Codes = new List<string>() { "001" };
                    break;
                case "M":
                    _Group_Product_Codes = new List<string>() { "002" };
                    break;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Assessment.Any())
                {
                    result.AddRange(db.IFRS9_Assessment.AsNoTracking()
                        .Where(x=> _Group_Product_Codes.Contains(x.Group_Product_Code)).AsEnumerable()
                        .Select(x => new Tuple<string, string>(x.Group_Product_Code,
                        string.Format("{0} ({1})", x.Group_Product_Code, x.Group_Product_Name))));
                }
            }
            return result;
        }

        /// <summary>
        /// get Account
        /// </summary>
        /// <param name="account"></param>
        /// <param name="admin"></param>
        /// <returns></returns>
        public List<AccountViewModel> getAccount(string account, string admin)
        {
            List<AccountViewModel> datas = new List<AccountViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _admin = db.IFRS9_User.AsNoTracking().FirstOrDefault(x => x.User_Account == admin && x.Effective && x.DebtType != null);
                if (_admin == null)
                    return datas;
                var _debt = _admin.DebtType;
                var _debts = common.getDebt(_debt);
                var users = db.IFRS9_User.AsNoTracking()
                    .Where(x => x.User_Account != admin && x.Effective)
                    .Where(x => x.DebtType == null || _debts.Contains(x.DebtType), _debt != "A");
                users = users.Where(x => x.User_Account.Contains(account), !account.IsNullOrWhiteSpace());             
                datas.AddRange(
                    users.AsEnumerable()                  
                    .Select(x => new AccountViewModel()
                    {
                        User_Account = x.User_Account,
                        User_Name = x.User_Name,
                        AdminFlag = x.AdminFlag ? "Y" : "N",
                        LoginFlag = x.LoginFlag ? "Y" : "N",
                        DebtFlag = x.DebtType == "A" ? "A:(全部權限)" :
                                   x.DebtType == "B" ? "B:(債券)" :
                                   x.DebtType == "M" ? "M:(房貸)" : string.Empty
                    })
                );
            }
            return datas;
        }

        /// <summary>
        /// 查詢帳號使用紀錄
        /// </summary>
        /// <param name="account">查詢的條件</param>
        /// <param name="type">查詢的log,User=IFRS9_User_Log,Browser=IFRS9_Browse_Log,Event=IFRS9_Event_Log</param>
        /// <param name="range">查詢的範圍 今天,7天,全部 (tpye = user 使用)</param>
        /// <returns></returns>
        public List<AccountLogViewModel> getAccountLog(AccountLogViewModel model, string type,string range)
        {
            List<AccountLogViewModel> datas = new List<AccountLogViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (type == "User")
                {
                    var date = DateTime.Now.Date;
                    var date2 = date.AddDays(-7);
                    string account = model.User_Account;
                    var result = db.IFRS9_User_Log.AsNoTracking()
                        .Where(x => x.User_Account.Contains(account), !account.IsNullOrWhiteSpace())
                        .Where(x => x.Login_Date == date, range == "1")
                        .Where(x => x.Login_Date > date2,range == "7")
                        .AsEnumerable();
                    datas = result.OrderByDescending(x => x.Login_Time)
                             .Select(x => new AccountLogViewModel
                             {
                                 User_Account = x.User_Account,
                                 User_Name = x.IFRS9_User.User_Name,
                                 Login_Time = x.Login_Time.dateTimeToStr(),
                                 Logout_Time = x.Logout_Time.dateTimeToStr(),
                                 IP_Location = x.Login_IP
                             }).ToList();
                }
                if (type == "Browser")
                {
                    string account = model.User_Account;
                    string Login_Time = model.Login_Time;
                    datas = db.IFRS9_Browse_Log.AsNoTracking()
                            .Where(x => x.User_Account == account).AsEnumerable()
                            .Where(x => x.Login_Time.ToString("yyyy/MM/dd HH:mm:ss") == Login_Time)
                            .OrderByDescending(x => x.Browse_Time)
                            .Select(x => new AccountLogViewModel
                            {
                                User_Account = x.User_Account,
                                Login_Time = x.Login_Time.dateTimeToStr(),
                                Menu_Id = x.Menu_Id,
                                Menu_Detail = x.IFRS9_Menu_Sub.Menu_Detail,
                                Browse_Time = x.Browse_Time.dateTimeToStr()
                            }).ToList();
                }
                if (type == "Event")
                {
                    string account = model.User_Account;
                    string Login_Time = model.Login_Time;
                    string Menu_Id = model.Menu_Id;
                    string Browse_Time = model.Browse_Time;
                    datas = db.IFRS9_Event_Log.AsNoTracking()
                              .Where(x => x.User_Account == account &&
                                          x.Menu_Id == Menu_Id ).AsEnumerable()
                              .Where(x => x.Login_Time.ToString("yyyy/MM/dd HH:mm:ss") == Login_Time &&
                                          x.Browse_Time.ToString("yyyy/MM/dd HH:mm:ss") == Browse_Time)
                             .OrderByDescending(x => x.Event_Begin)
                             .Select(x => new AccountLogViewModel
                             {
                                 Menu_Id = x.Menu_Id,
                                 Browse_Time = x.Browse_Time.dateTimeToStr(),
                                 Action_Name = x.Action_Name,
                                 Event_Name = x.Event_Name,
                                 Event_Begin = x.Event_Begin.dateTimeToStr(),
                                 Event_Complete = x.Event_Complete.dateTimeToStr(),
                                 Event_Flag = x.Event_Flag != null ? ( x.Event_Flag.Value ? "Y" : "N" ): "N"
                             }).ToList();
                }
            }
            return datas;
        }

        /// <summary>
        /// get 權限資料資料
        /// </summary>
        /// <param name="productCode">001,002</param>
        /// <param name="tableId">D64,D66</param>
        /// <param name="type">SetAssessmentType enum</param>
        /// <param name="effective">true = 只找Effective = Y</param>
        /// <returns></returns>
        public List<SetAssessmentViewModel> getAssessment(string productCode, string tableId, SetAssessmentType type,bool effective = true)
        {
            List<SetAssessmentViewModel> datas = new List<SetAssessmentViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {

                if (type == SetAssessmentType.Assessment)
                {
                    datas = db.IFRS9_Assessment_Config.AsNoTracking()
                        .Where(x=>x.Group_Product_Code == productCode)
                        .Where(x=>x.Effective == "Y", effective)
                        .Select(x => new SetAssessmentViewModel()
                    {
                        Group_Product_Code = x.Group_Product_Code,
                        Group_Product_Name = x.IFRS9_Assessment.Group_Product_Name,
                        Table_Id = x.Table_Id
                    }).ToList();
                }
                if (type == SetAssessmentType.Auditor || type == SetAssessmentType.Presented)
                {
                    var users = db.IFRS9_User.AsNoTracking().Where(x=>x.Effective).ToList();
                    var data = db.IFRS9_Assessment_Config.AsNoTracking()
                        .Where(x=> x.Effective == "Y", effective)
                        .FirstOrDefault(x => x.Group_Product_Code == productCode
                        && x.Table_Id == tableId);
                    string sql = string.Empty;
                    string sql2 = " and  Effective = 'Y' ";
                    if (type == SetAssessmentType.Auditor && data != null)
                    {
                        sql = $@"
select *
from IFRS9_Assessment_Auditor_Config
where Group_Product_Code = {productCode.stringToStrSql()}
and  Table_Id =  {tableId.stringToStrSql()}
";
                        if (effective)
                            sql += sql2;
                        var _data = db.Database.DynamicSqlQuery(sql);
                        if (_data.Any())
                            datas = _data.Select(x => new SetAssessmentViewModel()
                            {
                                User_Account = x.User_Account,
                                User_Name = users.Where(y=>y.User_Account == x.User_Account)
                                .DefaultIfEmpty().First().User_Name,
                            }).ToList();           
                    }
                    if (type == SetAssessmentType.Presented && data != null && data.IFRS9_User1.Any())
                    {
                        sql = $@"
select *
from IFRS9_Assessment_Presented_Config
where Group_Product_Code = {productCode.stringToStrSql()}
and  Table_Id =  {tableId.stringToStrSql()}
";
                        if (effective)
                            sql += sql2;
                        var _data = db.Database.DynamicSqlQuery(sql);
                        if (_data.Any())
                            datas = _data.Select(x => new SetAssessmentViewModel()
                            {
                                User_Account = x.User_Account,
                                User_Name = users.Where(y => y.User_Account == x.User_Account)
                                .DefaultIfEmpty().First().User_Name,
                            }).ToList();
                    }
                }
            }
            return datas;
        }

        /// <summary>
        /// get assessment can add Auditor or Presented User
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public List<Tuple<string, string>> getAssessmentAddUser(string productCode, string tableId, SetAssessmentType type)
        {
            List<Tuple<string, string>> result = new List<Tuple<string, string>>() { new Tuple<string, string>("", "") };
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<string> users = new List<string>();

                if ((type & SetAssessmentType.Auditor) == SetAssessmentType.Auditor ||
                    (type & SetAssessmentType.Presented) == SetAssessmentType.Presented)
                {
                    IEnumerable<string> _users =
                        getAssessment(productCode, tableId, SetAssessmentType.Presented).Select(x=>x.User_Account);
                    IEnumerable<string> _users1 =
                        getAssessment(productCode, tableId, SetAssessmentType.Auditor).Select(x => x.User_Account);
                    users = _users.Union(_users1).ToList();
                }
                var _debt = productCode == "001" ? "B" : productCode == "001" ? "M" : string.Empty;
                var _debts = common.getDebt(_debt); //債券類別 A,M,B
                _debts.Add("A");
                _debts = _debts.Distinct().ToList();
                result.AddRange(db.IFRS9_User.AsNoTracking()
                    .Where(x => 
                    !users.Contains(x.User_Account) &&
                    _debts.Contains(x.DebtType) &&
                    !x.AdminFlag && 
                    x.Effective).AsEnumerable()
                    .Select(x => new Tuple<string, string>
                    (x.User_Account, string.Format("{0}({1})", x.User_Account, x.User_Name))));
            }
            return result;
        }

        #endregion GetData

        #region SaveData

        /// <summary>
        /// save menu
        /// </summary>
        /// <param name="menuSub"></param>
        /// <param name="userName"></param>
        /// <returns></returns>
        public MSGReturnModel saveMenu(List<CheckBoxListInfo> menuSub, string userAccount)
        {
            MSGReturnModel result = new MSGReturnModel();
            if (!menuSub.Any() && userAccount.IsNullOrWhiteSpace())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<IFRS9_Menu_Set> sets = db.IFRS9_Menu_Set.Where(x =>
                userAccount.Equals(x.User_Account)).ToList();
                foreach (CheckBoxListInfo item in menuSub)
                {
                    IFRS9_Menu_Set set = sets.FirstOrDefault(x => item.Value.Equals(x.Menu_Id));
                    if (set != null) //原本有設定
                    {
                        if (!item.IsChecked) //設定無權限
                        {
                            if (set.Effective == "Y") //目前有權限
                            {
                                set.Effective = "N";
                                set.LastUpdate_User = _UserInfo._user;
                                set.LastUpdate_Date = _UserInfo._date;
                                set.LastUpdate_Time = _UserInfo._time;
                            }
                        }
                        else //設定無權限
                        {
                            if (set.Effective != "Y")//目前無權限
                            {
                                set.Effective = "Y";
                                set.LastUpdate_User = _UserInfo._user;
                                set.LastUpdate_Date = _UserInfo._date;
                                set.LastUpdate_Time = _UserInfo._time;
                            }
                        }
                    }
                    else //原本無設定
                    {
                        if (item.IsChecked)
                            db.IFRS9_Menu_Set.Add(new IFRS9_Menu_Set()
                            {
                                User_Account = userAccount,
                                Menu_Id = item.Value,
                                Create_User = _UserInfo._user,
                                Create_Date = _UserInfo._date,
                                Create_Time = _UserInfo._time,
                                Effective = "Y"
                            });
                    }
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(null,
                        $"message: {ex.Message}" +
                        $", inner message {ex.InnerException?.InnerException?.Message}");
                }
            }
            return result;
        }

        /// <summary>
        /// save Account
        /// </summary>
        /// <param name="action"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel saveAccount(string action, IFRS9_User data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            bool changeFlag = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (action == Action_Type.Add.ToString())
                {
                    if (db.IFRS9_User.Any(x => x.User_Account == data.User_Account && !x.Effective))
                    {
                        result.DESCRIPTION = "剛帳號已存在,目前於失效(刪除過)狀態,如需復原請於資料庫動作!";
                        return result;
                    }
                    db.IFRS9_User.Add(new IFRS9_User()
                    {
                        User_Account = data.User_Account,
                        User_Password = data.User_Password.stringToSHA512(),
                        User_Name = data.User_Name,
                        AdminFlag = data.AdminFlag,
                        Effective = true,
                        LoginFlag = false,
                        DebtType = data.DebtType.IsNullOrWhiteSpace() ? null : data.DebtType,
                        Create_User = _UserInfo._user,
                        Create_Date = _UserInfo._date,
                        Create_Time = _UserInfo._time
                    });
                    changeFlag = true;
                }
                if (action == Action_Type.Edit.ToString())
                {
                    var change = db.IFRS9_User.SingleOrDefault(x => x.User_Account == data.User_Account);
                    if (change != null)
                    {
                        if (!data.User_Password.IsNullOrWhiteSpace())
                        {
                            change.User_Password = data.User_Password.stringToSHA512();
                        }
                        change.User_Name = data.User_Name;
                        change.AdminFlag = data.AdminFlag;
                        change.LoginFlag = data.LoginFlag;
                        change.DebtType = data.DebtType.IsNullOrWhiteSpace() ? null : data.DebtType;
                        change.LastUpdate_User = _UserInfo._user;
                        change.LastUpdate_Date = _UserInfo._date;
                        change.LastUpdate_Time = _UserInfo._time;
                        changeFlag = true;
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    }
                }
                if (action == Action_Type.Dele.ToString())
                {
                    var del = db.IFRS9_User.SingleOrDefault(x => x.User_Account == data.User_Account);
                    if (del != null)
                    {
                        del.LastUpdate_User = _UserInfo._user;
                        del.LastUpdate_Date = _UserInfo._date;
                        del.LastUpdate_Time = _UserInfo._time;
                        del.Effective = false;
                        changeFlag = true;
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    }
                }
                if (changeFlag)
                {
                    try
                    {
                        db.SaveChanges(); //Save
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                        if (action == Action_Type.Dele.ToString())
                            result.DESCRIPTION = Message_Type.delete_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// Assessment Add
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public MSGReturnModel AssessmentAdd(SetAssessmentViewModel model, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            string sql = string.Empty;
            bool updateFlag = true;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (type == SetAssessmentType.Assessment.ToString())
                {
                    if (db.IFRS9_Assessment_Config.AsNoTracking().Any(x =>
                     x.Group_Product_Code == model.Group_Product_Code &&
                     x.Table_Id == model.Table_Id && x.Effective == "Y"))
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_Save.GetDescription(null, model.Table_Id);
                        return result;
                    }
                    else
                    {
                        var _data = db.IFRS9_Assessment_Config.FirstOrDefault(x =>
                             x.Group_Product_Code == model.Group_Product_Code &&
                              x.Table_Id == model.Table_Id && x.Effective != "Y");
                        if (_data != null)
                        {
                            _data.LastUpdate_User = _UserInfo._user;
                            _data.LastUpdate_Date = _UserInfo._date;
                            _data.LastUpdate_Time = _UserInfo._time;
                            _data.Effective = "Y";
                        }
                        else
                            db.IFRS9_Assessment_Config.Add(
                                new IFRS9_Assessment_Config()
                                {
                                    Group_Product_Code = model.Group_Product_Code,
                                    Table_Id = model.Table_Id,
                                    Create_User = _UserInfo._user,
                                    Create_Date = _UserInfo._date,
                                    Create_Time = _UserInfo._time,
                                    Effective = "Y"
                                });
                    }
                }
                if (type == SetAssessmentType.Auditor.ToString())
                {
                    var saveDb = db.IFRS9_Assessment_Config.FirstOrDefault(x=>
                    x.Group_Product_Code == model.Group_Product_Code &&
                    x.Table_Id == model.Table_Id && x.Effective == "Y");
                    var userInfo = db.IFRS9_User.FirstOrDefault(x =>
                    x.User_Account == model.User_Account && x.Effective);
                    if (saveDb == null || userInfo == null ||
                        getAssessment(model.Group_Product_Code, model.Table_Id, SetAssessmentType.Auditor)
                        .Any(x=>x.User_Account == userInfo.User_Account))
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                        return result;
                    }
                    saveDb.IFRS9_User.Add(userInfo);
                    saveDb.LastUpdate_User = _UserInfo._user;
                    saveDb.LastUpdate_Date = _UserInfo._date;
                    saveDb.LastUpdate_Time = _UserInfo._time;
                    updateFlag = getAssessment(model.Group_Product_Code, model.Table_Id, SetAssessmentType.Auditor, false)
                        .Any(x => x.User_Account == userInfo.User_Account);
                    if (updateFlag)
                    {
                        sql = $@"
update IFRS9_Assessment_Auditor_Config
   SET LastUpdate_User = {_UserInfo._user.stringToStrSql()},
       LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
	   LastUpdate_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'Y'
where Group_Product_Code = {model.Group_Product_Code.stringToStrSql()}
and   Table_Id = {model.Table_Id.stringToStrSql()}
and   User_Account = {model.User_Account.stringToStrSql()} ; ";
                    }
                    else
                    {
                        sql = $@"
update IFRS9_Assessment_Auditor_Config
   SET Create_User = {_UserInfo._user.stringToStrSql()},
       Create_Date = {_UserInfo._date.dateTimeToStrSql()},
	   Create_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'Y'
where Group_Product_Code = {model.Group_Product_Code.stringToStrSql()}
and   Table_Id = {model.Table_Id.stringToStrSql()}
and   User_Account = {model.User_Account.stringToStrSql()} ; ";
                    }
                }
                if (type == SetAssessmentType.Presented.ToString())
                {
                    var saveDb = db.IFRS9_Assessment_Config.FirstOrDefault(x =>
                                    x.Group_Product_Code == model.Group_Product_Code &&
                                    x.Table_Id == model.Table_Id && x.Effective == "Y");
                    var userInfo = db.IFRS9_User.FirstOrDefault(x =>
                                    x.User_Account == model.User_Account && x.Effective);
                    if (saveDb == null || userInfo == null ||
                        getAssessment(model.Group_Product_Code, model.Table_Id, SetAssessmentType.Presented)
                        .Any(x => x.User_Account == userInfo.User_Account)
                        )
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                        return result;
                    }
                    saveDb.IFRS9_User1.Add(userInfo);
                    saveDb.LastUpdate_User = _UserInfo._user;
                    saveDb.LastUpdate_Date = _UserInfo._date;
                    saveDb.LastUpdate_Time = _UserInfo._time;
                    updateFlag = getAssessment(model.Group_Product_Code, model.Table_Id, SetAssessmentType.Presented, false)
                                .Any(x => x.User_Account == userInfo.User_Account);
                    if (updateFlag)
                    {
                        sql = $@"
update IFRS9_Assessment_Presented_Config
   SET LastUpdate_User = {_UserInfo._user.stringToStrSql()},
       LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
	   LastUpdate_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'Y'
where Group_Product_Code = {model.Group_Product_Code.stringToStrSql()}
and   Table_Id = {model.Table_Id.stringToStrSql()}
and   User_Account = {model.User_Account.stringToStrSql()} ; ";
                    }
                    else
                    {
                        sql = $@"
update IFRS9_Assessment_Presented_Config
   SET Create_User = {_UserInfo._user.stringToStrSql()},
       Create_Date = {_UserInfo._date.dateTimeToStrSql()},
	   Create_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'Y'
where Group_Product_Code = {model.Group_Product_Code.stringToStrSql()}
and   Table_Id = {model.Table_Id.stringToStrSql()}
and   User_Account = {model.User_Account.stringToStrSql()} ; ";
                    }
                }
                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        db.SaveChanges();
                        if(sql.Length > 0)
                            db.Database.ExecuteSqlCommand(sql);
                        dbContextTransaction.Commit();
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = ex.exceptionMessage();
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

        /// <summary>
        /// Assessment Del
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public MSGReturnModel AssessmentDel(List<SetAssessmentViewModel> model, string type)
        {
            int userCount = 0;
            MSGReturnModel result = new MSGReturnModel();
            if (!model.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.not_Find_Update_Data.GetDescription();
                return result;
            }  
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                string sql = string.Empty;
                if (type == SetAssessmentType.Assessment.ToString())
                {
                    List<string> errors = new List<string>();
                    List<IFRS9_Assessment_Config> removeDatas = new List<IFRS9_Assessment_Config>();
                    model.ForEach(x =>
                    {
                        var assessment = db.IFRS9_Assessment_Config.FirstOrDefault(y =>
                        y.Group_Product_Code == x.Group_Product_Code &&
                        y.Table_Id == x.Table_Id && y.Effective == "Y");
                        if (assessment == null)
                        {
                            errors.Add(x.Table_Id);
                        }
                        else
                        {
                            if (getAssessment(x.Group_Product_Code, x.Table_Id, SetAssessmentType.Presented).Any()
                            || getAssessment(x.Group_Product_Code, x.Table_Id, SetAssessmentType.Auditor).Any())
                                errors.Add(x.Table_Id);
                            else
                            {
                                assessment.LastUpdate_User = _UserInfo._user;
                                assessment.LastUpdate_Date = _UserInfo._date;
                                assessment.LastUpdate_Time = _UserInfo._time;
                                assessment.Effective = "N";
                            }
                        }
                    });
                    if (errors.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(null,
                            string.Format("{0} , {1}{2}", string.Join("、", errors),
                            "有關聯的覆核帳號or呈送帳號未刪除!",
                            "或者" + Message_Type.already_Change.GetDescription()
                            ));
                        return result;
                    }
                }
                if (type == SetAssessmentType.Auditor.ToString())
                {
                    var Group_Product_Code = model.First().Group_Product_Code;
                    var Table_Id = model.First().Table_Id;
                    var data = db.IFRS9_Assessment_Config.FirstOrDefault(x =>
                    x.Group_Product_Code == Group_Product_Code &&
                    x.Table_Id == Table_Id && x.Effective == "Y");
                    if (data == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                        return result;
                    }
                    else
                    {
                        var users = model.Select(z => z.User_Account);
                        userCount = users.Count();
                        var removeDatas = db.IFRS9_User
                            .Where(x => users.Contains(x.User_Account) && x.Effective).ToList();
                        if (getAssessment(Group_Product_Code, Table_Id, SetAssessmentType.Auditor)
                            .Count(x => users.Contains(x.User_Account))
                            != userCount || userCount != removeDatas.Count)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                            return result;
                        }
                        else
                        {
                            sql = $@"
update IFRS9_Assessment_Auditor_Config
   SET LastUpdate_User = {_UserInfo._user.stringToStrSql()},
       LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
	   LastUpdate_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'N'
where Group_Product_Code = {Group_Product_Code.stringToStrSql()}
and   Table_Id = {Table_Id.stringToStrSql()}
and   User_Account IN  ({users.ToList().stringListToInSql()}) ; ";                         
                                data.LastUpdate_User = _UserInfo._user;
                                data.LastUpdate_Date = _UserInfo._date;
                                data.LastUpdate_Time = _UserInfo._time;
                        }
                    }
                }
                if (type == SetAssessmentType.Presented.ToString())
                {
                    var Group_Product_Code = model.First().Group_Product_Code;
                    var Table_Id = model.First().Table_Id;
                    var data = db.IFRS9_Assessment_Config.FirstOrDefault(x =>
                                x.Group_Product_Code == Group_Product_Code &&
                                x.Table_Id == Table_Id && x.Effective == "Y");
                    if (data == null)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                        return result;
                    }
                    else
                    {
                        var users = model.Select(z => z.User_Account);
                        userCount = users.Count();
                        var removeDatas = db.IFRS9_User
                            .Where(x => users.Contains(x.User_Account)).ToList();
                        if (getAssessment(Group_Product_Code, Table_Id, SetAssessmentType.Presented)
                            .Count(x => users.Contains(x.User_Account)) != userCount ||
                            userCount != removeDatas.Count)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.already_Change.GetDescription();
                            return result;
                        }
                        else
                        {
                            sql = $@"
update IFRS9_Assessment_Presented_Config
   SET LastUpdate_User = {_UserInfo._user.stringToStrSql()},
       LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
	   LastUpdate_Time = {(_UserInfo._time.ToString(@"hh\:mm\:ss").stringToStrSql())},
       Effective = 'N'
where Group_Product_Code = {Group_Product_Code.stringToStrSql()}
and   Table_Id = {Table_Id.stringToStrSql()}
and   User_Account IN  ({users.ToList().stringListToInSql()}) ; ";
                            data.LastUpdate_User = _UserInfo._user;
                            data.LastUpdate_Date = _UserInfo._date;
                            data.LastUpdate_Time = _UserInfo._time;
                        }
                    }
                }
                using (var dbContextTransaction = db.Database.BeginTransaction())
                {
                    try
                    {
                        var updateCount = 0;
                        if (sql.Length > 0)
                            updateCount = db.Database.ExecuteSqlCommand(sql);
                        db.SaveChanges();
                        if (updateCount == userCount)
                            dbContextTransaction.Commit();
                        else
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.save_Fail.GetDescription();
                            return result;
                        }
                        result.RETURN_FLAG = true;
                        result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                    }
                    catch (DbUpdateException ex)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = ex.exceptionMessage();
                    }
                }
            }
            return result;
        }
        #endregion SaveData

        #region privateFunction



        #endregion

        #region getMenuMainAll
        public Tuple<bool, List<MenuMainViewModel>> getMenuMainAll()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Menu_Main.Any())
                {
                    return new Tuple<bool, List<MenuMainViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.IFRS9_Menu_Main.AsNoTracking().AsEnumerable()
                                        select DbToMenuMainModel(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<MenuMainViewModel>>(true, new List<MenuMainViewModel>());
        }
        #endregion

        #region DbToMenuMainModel
        private MenuMainViewModel DbToMenuMainModel(IFRS9_Menu_Main data)
        { 
            return new MenuMainViewModel()
            {
                Menu = data.Menu,
                Menu_Id = data.Menu_Id,
                Menu_Detail = data.Menu_Detail,
                Class = data.Class,
                Create_User = data.Create_User,
                Create_Date = TypeTransfer.dateTimeNToString(data.Create_Date),
                Create_Time = data.Create_Time.timeSpanNToStr(),
                LastUpdate_User = data.LastUpdate_User,
                LastUpdate_Date = TypeTransfer.dateTimeNToString(data.LastUpdate_Date),
                LastUpdate_Time = data.LastUpdate_Time.timeSpanNToStr()
            };
        }
        #endregion

        #region  getMenuMain
        public Tuple<bool, List<MenuMainViewModel>> getMenuMain(MenuMainViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Menu_Main.Any())
                {
                    var query = db.IFRS9_Menu_Main.AsNoTracking()
                                  .Where(x => x.Menu == dataModel.Menu, dataModel.Menu.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Menu_Id == dataModel.Menu_Id, dataModel.Menu_Id.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Menu_Detail.Contains(dataModel.Menu_Detail), dataModel.Menu_Detail.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Class == dataModel.Class, dataModel.Class.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<MenuMainViewModel>>(query.Any(), query.AsEnumerable()
                                                                                 .Select(x => { return DbToMenuMainModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<MenuMainViewModel>>(false, new List<MenuMainViewModel>());
        }
        #endregion

        #region saveMenuMain
        public MSGReturnModel saveMenuMain(string actionType, MenuMainViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    IFRS9_Menu_Main dataEdit = new IFRS9_Menu_Main();

                    if (actionType == "Add")
                    {
                        dataEdit.Menu = dataModel.Menu;
                        dataEdit.Menu_Id = dataModel.Menu + "Main";
                        dataEdit.Create_User = _UserInfo._user;
                        dataEdit.Create_Date = _UserInfo._date;
                        dataEdit.Create_Time = _UserInfo._time;
                    }
                    else if (actionType == "Modify")
                    {
                        dataEdit = db.IFRS9_Menu_Main
                                     .Where(x => x.Menu == dataModel.Menu)
                                     .FirstOrDefault();
                        dataEdit.LastUpdate_User = _UserInfo._user;
                        dataEdit.LastUpdate_Date = _UserInfo._date;
                        dataEdit.LastUpdate_Time = _UserInfo._time;
                    }

                    dataEdit.Menu_Detail = dataModel.Menu_Detail;
                    dataEdit.Class = dataModel.Class;

                    if (actionType == "Add")
                    {
                        db.IFRS9_Menu_Main.Add(dataEdit);
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteMenuMain
        public MSGReturnModel deleteMenuMain(string menu)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var menuMain = db.IFRS9_Menu_Main
                                     .Where(x => x.Menu == menu).FirstOrDefault();
                    string sql = "";
                    if (menuMain != null)
                    {
                        var menuSub = db.IFRS9_Menu_Sub.Where(x => x.Menu == menu).ToList();

                        string menuId = "";

                        for (int i=0; i < menuSub.Count;i++)
                        {
                            menuId += "'" + menuSub[i].Menu_Id + "',";
                        }

                        if (menuId.Length > 0)
                        {
                            menuId = menuId.Substring(0, menuId.Length - 1);

                            sql = $@"
                                         DELETE IFRS9_Event_Log WHERE Menu_Id IN ({menuId});   
                                         DELETE IFRS9_Browse_Log WHERE Menu_Id IN ({menuId});   
                                         DELETE IFRS9_Menu_Set WHERE Menu_Id IN ({menuId});   
                                         DELETE IFRS9_Menu_Sub WHERE Menu_Id IN ({menuId});
                                    ";
                        }

                        sql += $" DELETE IFRS9_Menu_Main WHERE Menu = '{menu}';";

                        db.Database.ExecuteSqlCommand(sql);
                    }

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region getMenuSubAll
        public Tuple<bool, List<MenuSubViewModel>> getMenuSubAll()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Menu_Sub.Any())
                {
                    return new Tuple<bool, List<MenuSubViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.IFRS9_Menu_Sub.AsNoTracking().AsEnumerable()
                                        select DbToMenuSubModel(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<MenuSubViewModel>>(true, new List<MenuSubViewModel>());
        }
        #endregion

        #region DbToMenuSubModel
        private MenuSubViewModel DbToMenuSubModel(IFRS9_Menu_Sub data)
        {
            return new MenuSubViewModel()
            {
                Menu = data.Menu,
                Menu_Id = data.Menu_Id,
                Menu_Detail = data.Menu_Detail,
                Class = data.Class,
                Href = data.Href,
                SortNum = data.SortNum.ToString(),
                DebtType = data.DebtType,
                Create_User = data.Create_User,
                Create_Date = TypeTransfer.dateTimeNToString(data.Create_Date),
                Create_Time = data.Create_Time.timeSpanNToStr(),
                LastUpdate_User = data.LastUpdate_User,
                LastUpdate_Date = TypeTransfer.dateTimeNToString(data.LastUpdate_Date),
                LastUpdate_Time = data.LastUpdate_Time.timeSpanNToStr()
            };
        }
        #endregion

        #region  getMenuSub
        public Tuple<bool, List<MenuSubViewModel>> getMenuSub(MenuSubViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Menu_Sub.Any())
                {
                    var query = db.IFRS9_Menu_Sub.AsNoTracking()
                                  .Where(x => x.Menu == dataModel.Menu, dataModel.Menu.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Menu_Id == dataModel.Menu_Id, dataModel.Menu_Id.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Menu_Detail.Contains(dataModel.Menu_Detail), dataModel.Menu_Detail.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Class == dataModel.Class, dataModel.Class.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<MenuSubViewModel>>(query.Any(), query.AsEnumerable()
                                                                                .Select(x => { return DbToMenuSubModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<MenuSubViewModel>>(false, new List<MenuSubViewModel>());
        }
        #endregion

        #region saveMenuSub
        public MSGReturnModel saveMenuSub(string actionType, MenuSubViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    IFRS9_Menu_Sub dataEdit = new IFRS9_Menu_Sub();

                    if (actionType == "Add")
                    {
                        dataEdit.Menu_Id = dataModel.Menu_Id;
                        dataEdit.Create_User = _UserInfo._user;
                        dataEdit.Create_Date = _UserInfo._date;
                        dataEdit.Create_Time = _UserInfo._time;
                    }
                    else if (actionType == "Modify")
                    {
                        dataEdit = db.IFRS9_Menu_Sub
                                     .Where(x => x.Menu_Id == dataModel.Menu_Id)
                                     .FirstOrDefault();
                        dataEdit.LastUpdate_User = _UserInfo._user;
                        dataEdit.LastUpdate_Date = _UserInfo._date;
                        dataEdit.LastUpdate_Time = _UserInfo._time;
                    }

                    dataEdit.Menu = dataModel.Menu;
                    dataEdit.Menu_Detail = dataModel.Menu_Detail;
                    dataEdit.Class = dataModel.Class;
                    dataEdit.Href = dataModel.Href;
                    dataEdit.SortNum = int.Parse(dataModel.SortNum);
                    dataEdit.DebtType = dataModel.DebtType;

                    if (actionType == "Add")
                    {
                        db.IFRS9_Menu_Sub.Add(dataEdit);
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteMenuSub
        public MSGReturnModel deleteMenuSub(string menuId)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    string sql = $@"
                                     DELETE IFRS9_Event_Log WHERE Menu_Id = '{menuId}';   
                                     DELETE IFRS9_Browse_Log WHERE Menu_Id = '{menuId}';   
                                     DELETE IFRS9_Menu_Set WHERE Menu_Id = '{menuId}';   
                                     DELETE IFRS9_Menu_Sub WHERE Menu_Id = '{menuId}';
                                   ";
                    db.Database.ExecuteSqlCommand(sql);

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion
    }
}