using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
using Transfer.Models;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    public class SystemForITController : CommonForITController
    {
        private ISystemForITRepository SystemForITRepository;
        private string adminAccount;

        public SystemForITController()
        {
            this.SystemForITRepository = new SystemForITRepository();
            this.adminAccount = "";
        }

        /// <summary>
        /// 設定主頁 (預留)
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 設定 Account 的權限
        /// </summary>
        /// <returns></returns>
        public ActionResult SetAccount()
        {
            var jqgridInfo = new AccountViewModel().TojqGridData(null, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            ViewBag.debt = GetDebtSelectOption();
            return View();
        }

        /// <summary>
        /// 設定 Menu 的權限
        /// </summary>
        /// <returns></returns>
        public ActionResult SetMenu()
        {
            ViewBag.users = new SelectList(
                 SystemForITRepository.getUser("menu", false, adminAccount)
                 .Select(x => new { Text = x.Item2, Value = x.Item1 }), "Value", "Text");
            return View();
        }

        /// <summary>
        /// 設定 Assessment 的權限
        /// </summary>
        /// <returns></returns>
        public ActionResult SetAssessment()
        {
            ViewBag.productCode = new SelectList(
                 SystemForITRepository.getProductCode(adminAccount)
                 .Select(x => new { Value = x.Item1, Text = x.Item2 }), "Value", "Text");
            return View();
        }

        /// <summary>
        /// 排程重啟
        /// </summary>
        /// <returns></returns>
        public ActionResult TaskSchedule()
        {
            return View();
        }

        /// <summary>
        /// 帳號使用紀錄
        /// </summary>
        /// <returns></returns>
        public ActionResult AccountLog()
        {
            List<RadioButton> data = new List<RadioButton>();
            data.Add(new RadioButton() { Name = "Effective", Checked = true, Value = "只查詢有效帳號", Id = "Effective" });
            data.Add(new RadioButton() { Name = "Effective", Checked = false, Value = "全部查詢(包含失效)", Id = "EffectiveAll" });
            ViewBag.radio = data;
            data = new List<RadioButton>();
            data.Add(new RadioButton() { Name = "Range", Checked = true, Value = "今天", Id = "1" });
            data.Add(new RadioButton() { Name = "Range", Checked = false, Value = "一周內", Id = "7" });
            data.Add(new RadioButton() { Name = "Range", Checked = false, Value = "全部查詢", Id = "All" });
            ViewBag.radio2 = data;
            ViewBag.users = new SelectList(SystemForITRepository.getUser("log", false, adminAccount)
                   .Select(x => new { Text = x.Item2, Value = x.Item1 }), "Value", "Text");
            return View();
        }

        /// <summary>
        /// 選單類別維護
        /// </summary>
        /// <returns></returns>
        public ActionResult MenuMain()
        {
            return View();
        }

        /// <summary>
        /// 選單明細維護
        /// </summary>
        /// <returns></returns>
        public ActionResult MenuSub()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((SystemForITRepository.getMenuMainAll().Item2)
                                    .Select(x => new SelectOption()
                                    { Text = (x.Menu + "：" + x.Menu_Detail), Value = x.Menu }));
            ViewBag.MenuClass = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "A：All(所有權限)", Value = "A" },
                                              new SelectOption() { Text = "B：Bonds(債券)", Value = "B" },
                                              new SelectOption() { Text = "M：Mortgage(房貸)", Value = "M" }
                                           };
            ViewBag.DebtType = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        public JsonResult changeEffective(bool searchAll)
        {
            return Json(SystemForITRepository.getUser("log", searchAll, adminAccount));
        }

        /// <summary>
        /// 抓取使用者的 menu 設定
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        public JsonResult GetUser(string name)
        {
            MSGReturnModel result = new MSGReturnModel();
            List<CheckBoxListInfo> data
                = SystemForITRepository.getMenu(name);
            if (data.Any())
            {
                result = new MSGReturnModel()
                {
                    RETURN_FLAG = true,
                    Datas = Json(Extension.CheckBoxString("menuSet", data, null, 3).ToString())
                };
            }
            else
                result = new MSGReturnModel()
                {
                    RETURN_FLAG = false,
                    DESCRIPTION = Message_Type.not_Find_Any.GetDescription()
                };
            return Json(result);
        }

        /// <summary>
        /// 儲存使用者的menu設定
        /// </summary>
        /// <param name="data">menu Data</param>
        /// <param name="userName"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SaveMenu(List<CheckBoxListInfo> data, string userName)
        {
            MSGReturnModel result = SystemForITRepository.saveMenu(data, userName);
            return Json(result);
        }

        /// <summary>
        /// 儲存帳號
        /// </summary>
        /// <param name="action">前端動作</param>
        /// <param name="account">帳號</param>
        /// <param name="pwd">密碼</param>
        /// <param name="name">暱稱</param>
        /// <param name="adminFlag">是否為管理者</param>
        /// <param name="loginFlag">登入狀態</param>
        /// <param name="debtFlag">權限類別</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SaveAccount(string action, string account, string pwd, string name, string adminFlag, string loginFlag, string debtFlag)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (account.IsNullOrWhiteSpace() ||
                !EnumUtil.GetValues<Action_Type>().Select(x => x.ToString()).Contains(action))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            IFRS9_User data = new IFRS9_User();
            data.User_Account = account.Trim();
            data.User_Name = name == null ? string.Empty : name.Trim();
            if (action == Action_Type.Edit.ToString())
            {
                data.AdminFlag = adminFlag == "Y" ? true : false;
                data.LoginFlag = loginFlag == "Y" ? true : false;
                data.DebtType = debtFlag;
                if (!pwd.IsNullOrWhiteSpace())
                    data.User_Password = pwd.Trim();
            }
            if (action == Action_Type.Add.ToString())
            {
                data.User_Password = pwd.Trim();
                data.DebtType = debtFlag;
                data.AdminFlag = adminFlag == "Y" ? true : false;
            }
            result = SystemForITRepository.saveAccount(action, data);
            if (result.RETURN_FLAG)
            {
                SetCacheData(SystemForITRepository.getAccount(string.Empty,""));
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢帳號
        /// </summary>
        /// <param name="account">查詢的條件</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SearchAccount(string account)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            List<AccountViewModel> datas =
                SystemForITRepository.getAccount(account, adminAccount);
            if (datas.Any())
            {
                SetCacheData(datas);
                result.RETURN_FLAG = true;
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢帳號使用紀錄
        /// </summary>
        /// <param name="account">查詢的條件</param>
        /// <param name="type">查詢的log,User=IFRS9_User_Log,Browser=IFRS9_Browse_Log,Event=IFRS9_Event_Log</param>
        /// <param name="range">查詢的範圍 今天,7天,全部 (tpye = user 使用)</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SearchAccountLog(AccountLogViewModel model, string type, string range)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            List<AccountLogViewModel> datas = SystemForITRepository.getAccountLog(model, type, range);
            if (datas.Any())
            {
                SetCacheDataInAccountLog(datas, type);
                result.RETURN_FLAG = true;
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢信評
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="assessmentFlag"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SearchAssessment(string productCode, string tableId, bool assessmentFlag)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = true;
            if (productCode.IsNullOrWhiteSpace())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            if (assessmentFlag)
            {
                SetCacheDataInAssessment(SystemForITRepository
                    .getAssessment(productCode, tableId, SetAssessmentType.Assessment), SetAssessmentType.Assessment);
            }
            else
            {
                if (tableId.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    return Json(result);
                }
                SetCacheDataInAssessment(SystemForITRepository
                    .getAssessment(productCode, tableId, SetAssessmentType.Auditor), SetAssessmentType.Auditor);
                SetCacheDataInAssessment(SystemForITRepository
                    .getAssessment(productCode, tableId, SetAssessmentType.Presented), SetAssessmentType.Presented);
            }
            return Json(result);
        }

        /// <summary>
        /// 新增信評設定
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult AssessmentSetAdd(SetAssessmentViewModel model, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            if (!EnumUtil.GetValues<SetAssessmentType>().Any(x => x.ToString() == type))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            result = SystemForITRepository.AssessmentAdd(model, type);
            if (result.RETURN_FLAG)
            {
                if (type == SetAssessmentType.Assessment.ToString())
                {
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(model.Group_Product_Code, model.Table_Id,
                        SetAssessmentType.Assessment), SetAssessmentType.Assessment);
                }
                if (type == SetAssessmentType.Auditor.ToString() ||
                    type == SetAssessmentType.Presented.ToString())
                {
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(model.Group_Product_Code, model.Table_Id,
                        SetAssessmentType.Auditor), SetAssessmentType.Auditor);
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(model.Group_Product_Code, model.Table_Id,
                        SetAssessmentType.Presented), SetAssessmentType.Presented);
                }
            }
            return Json(result);
        }

        /// <summary>
        /// 刪除信評設定
        /// </summary>
        /// <param name="models"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult AssessmentSetDel(List<SetAssessmentViewModel> models, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            if (!models.Any() || !EnumUtil.GetValues<SetAssessmentType>().Any(x => x.ToString() == type))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            result = SystemForITRepository.AssessmentDel(models, type);
            if (result.RETURN_FLAG)
            {
                if (type == SetAssessmentType.Assessment.ToString())
                {
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(models.First().Group_Product_Code, models.First().Table_Id,
                        SetAssessmentType.Assessment), SetAssessmentType.Assessment);
                }
                if (type == SetAssessmentType.Auditor.ToString() ||
                    type == SetAssessmentType.Presented.ToString())
                {
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(models.First().Group_Product_Code, models.First().Table_Id,
                        SetAssessmentType.Auditor), SetAssessmentType.Auditor);
                    SetCacheDataInAssessment(SystemForITRepository
                        .getAssessment(models.First().Group_Product_Code, models.First().Table_Id,
                        SetAssessmentType.Presented), SetAssessmentType.Presented);
                }
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢信評 可新增的 覆核帳號 or 呈送帳號
        /// </summary>
        /// <param name="productCode"></param>
        /// <param name="tableId"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SearchAssessmentAddUser(string productCode, string tableId, string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = true;
            if (productCode.IsNullOrWhiteSpace() || tableId.IsNullOrWhiteSpace())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            try
            {
                List<Tuple<string, string>> datas = new List<Tuple<string, string>>();
                if (type == SetAssessmentType.Auditor.ToString())
                {
                    datas = SystemForITRepository.getAssessmentAddUser(productCode, tableId, SetAssessmentType.Auditor);
                }
                if (type == SetAssessmentType.Presented.ToString())
                {
                    datas = SystemForITRepository.getAssessmentAddUser(productCode, tableId, SetAssessmentType.Presented);
                }
                result.Datas = Json(datas);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        //#region get Cache Data

        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata)
        {
            List<AccountViewModel> datas = new List<AccountViewModel>();
            if (CacheForIT.IsSet(CacheList.userDbData))
                datas = (List<AccountViewModel>)CacheForIT.Get(CacheList.userDbData);  //從Cache 抓資料
            return Json(jdata.modelToJqgridResult(datas)); //查詢資料
        }

        /// <summary>
        /// get user log data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheDataInUserLog(jqGridParam jdata, string type, string Menu_Id = null, string Browse_Time = null)
        {
            List<AccountLogViewModel> datas = new List<AccountLogViewModel>();
            if (type == "User")
            {
                if (CacheForIT.IsSet(CacheList.userLogInUserDbData))
                    datas = (List<AccountLogViewModel>)CacheForIT.Get(CacheList.userLogInUserDbData);  //從Cache 抓資料
            }
            if (type == "Browser")
            {
                if (CacheForIT.IsSet(CacheList.userLogInBrowserDbData))
                    datas = (List<AccountLogViewModel>)CacheForIT.Get(CacheList.userLogInBrowserDbData);  //從Cache 抓資料
            }
            if (type == "Event")
            {
                if (CacheForIT.IsSet(CacheList.userLogInEventDbData))
                    datas = (List<AccountLogViewModel>)CacheForIT.Get(CacheList.userLogInEventDbData);  //從Cache 抓資料
            }
            return Json(jdata.modelToJqgridResult(datas)); //查詢資料
        }

        /// <summary>
        /// get Assessment data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheDataInAssessment(jqGridParam jdata, string type)
        {
            List<SetAssessmentViewModel> datas = new List<SetAssessmentViewModel>();
            if (type == SetAssessmentType.Assessment.ToString())
            {
                if (CacheForIT.IsSet(CacheList.SetAssessment))
                    datas = (List<SetAssessmentViewModel>)CacheForIT.Get(CacheList.SetAssessment);  //從Cache 抓資料
            }
            if (type == SetAssessmentType.Auditor.ToString())
            {
                if (CacheForIT.IsSet(CacheList.SetAssessmentAuditor))
                    datas = (List<SetAssessmentViewModel>)CacheForIT.Get(CacheList.SetAssessmentAuditor);  //從Cache 抓資料
            }
            if (type == SetAssessmentType.Presented.ToString())
            {
                if (CacheForIT.IsSet(CacheList.SetAssessmentPresented))
                    datas = (List<SetAssessmentViewModel>)CacheForIT.Get(CacheList.SetAssessmentPresented);  //從Cache 抓資料
            }
            return Json(jdata.modelToJqgridResult(datas)); //查詢資料
        }

        //#endregion

        //#region set Cache Data

        /// <summary>
        /// 帳號資料Cache
        /// </summary>
        /// <param name="datas"></param>
        private void SetCacheData(List<AccountViewModel> datas)
        {
            CacheForIT.Invalidate(CacheList.userDbData); //清除
            CacheForIT.Set(CacheList.userDbData, datas); //把資料存到 Cache
        }

        /// <summary>
        /// 使用者logCache
        /// </summary>
        /// <param name="datas"></param>
        /// <param name="type"></param>
        private void SetCacheDataInAccountLog(List<AccountLogViewModel> datas, string type)
        {
            if (type == "User")
            {
                CacheForIT.Invalidate(CacheList.userLogInUserDbData); //清除
                CacheForIT.Set(CacheList.userLogInUserDbData, datas); //把資料存到 Cache
            }
            if (type == "Browser")
            {
                CacheForIT.Invalidate(CacheList.userLogInBrowserDbData); //清除
                CacheForIT.Set(CacheList.userLogInBrowserDbData, datas); //把資料存到 Cache
            }
            if (type == "Event")
            {
                CacheForIT.Invalidate(CacheList.userLogInEventDbData); //清除
                CacheForIT.Set(CacheList.userLogInEventDbData, datas); //把資料存到 Cache
            }
        }

        /// <summary>
        /// 權限設定Cache
        /// </summary>
        /// <param name="datas"></param>
        /// <param name="type"></param>
        private void SetCacheDataInAssessment(List<SetAssessmentViewModel> datas, SetAssessmentType type)
        {
            if (type == SetAssessmentType.Assessment)
            {
                CacheForIT.Invalidate(CacheList.SetAssessment); //清除
                CacheForIT.Set(CacheList.SetAssessment, datas); //把資料存到 Cache
            }
            if (type == SetAssessmentType.Auditor)
            {
                CacheForIT.Invalidate(CacheList.SetAssessmentAuditor); //清除
                CacheForIT.Set(CacheList.SetAssessmentAuditor, datas); //把資料存到 Cache
            }
            if (type == SetAssessmentType.Presented)
            {
                CacheForIT.Invalidate(CacheList.SetAssessmentPresented); //清除
                CacheForIT.Set(CacheList.SetAssessmentPresented, datas); //把資料存到 Cache
            }
        }
        //#endregion

        #region GetMenuMainAll
        [HttpPost]
        public JsonResult GetMenuMainAllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = SystemForITRepository.getMenuMainAll();

                result.RETURN_FLAG = queryResult.Item1;

                CacheForIT.Invalidate(CacheList.MenuMainDbfileData); //清除
                CacheForIT.Set(CacheList.MenuMainDbfileData, queryResult.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetCacheMenuMainData
        [HttpPost]
        public JsonResult GetCacheMenuMainData(jqGridParam jdata)
        {
            List<MenuMainViewModel> data = new List<MenuMainViewModel>();
            if (CacheForIT.IsSet(CacheList.MenuMainDbfileData))
            {
                data = (List<MenuMainViewModel>)CacheForIT.Get(CacheList.MenuMainDbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetMenuMainData
        [HttpPost]
        public JsonResult GetMenuMainData(MenuMainViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = SystemForITRepository.getMenuMain(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                CacheForIT.Invalidate(CacheList.MenuMainDbfileData); //清除
                CacheForIT.Set(CacheList.MenuMainDbfileData, queryResult.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveMenuMain
        [HttpPost]
        public JsonResult SaveMenuMain(string actionType, MenuMainViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = SystemForITRepository.saveMenuMain(actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription() + " " + resultSave.DESCRIPTION;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region DeleteMenuMain
        [HttpPost]
        public JsonResult DeleteMenuMain(string menu)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = SystemForITRepository.deleteMenuMain(menu);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetMenuSubAll
        [HttpPost]
        public JsonResult GetMenuSubAllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = SystemForITRepository.getMenuSubAll();

                result.RETURN_FLAG = queryResult.Item1;

                CacheForIT.Invalidate(CacheList.MenuSubDbfileData); //清除
                CacheForIT.Set(CacheList.MenuSubDbfileData, queryResult.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetCacheMenuSubData
        [HttpPost]
        public JsonResult GetCacheMenuSubData(jqGridParam jdata)
        {
            List<MenuSubViewModel> data = new List<MenuSubViewModel>();
            if (CacheForIT.IsSet(CacheList.MenuSubDbfileData))
            {
                data = (List<MenuSubViewModel>)CacheForIT.Get(CacheList.MenuSubDbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetMenuSubData
        [HttpPost]
        public JsonResult GetMenuSubData(MenuSubViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = SystemForITRepository.getMenuSub(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                CacheForIT.Invalidate(CacheList.MenuSubDbfileData); //清除
                CacheForIT.Set(CacheList.MenuSubDbfileData, queryResult.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveMenuSub
        [HttpPost]
        public JsonResult SaveMenuSub(string actionType, MenuSubViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = SystemForITRepository.saveMenuSub(actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription() + " " + resultSave.DESCRIPTION;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region DeleteMenuSub
        [HttpPost]
        public JsonResult DeleteMenuSub(string menuId)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = SystemForITRepository.deleteMenuSub(menuId);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetTaskSchedule
        public JsonResult GetTaskSchedule()
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = SystemForITRepository.getTaskSchedule(null);
            result.Datas = Json(data);
            return Json(result);
        }
        #endregion

        #region RestartTaskSchedule
        [HttpPost]
        public JsonResult RestartTaskSchedule(string triggerTask, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                switch (triggerTask)
                {
                    case "StartKRiskService":
                    case "StopKRiskService":
                    case "StartTomcatService":
                    case "StopTomcatService":
                        break;
                    //default:
                    //    reportDate = DateTime.Parse(reportDate).ToString("yyyyMMdd");
                    //break;
                }

                MSGReturnModel resultRestart = SystemForITRepository.restartTaskSchedule(triggerTask, reportDate);

                result.RETURN_FLAG = resultRestart.RETURN_FLAG;
                result.DESCRIPTION = "重啟完成";

                if (result.RETURN_FLAG == false)
                {
                    result.DESCRIPTION = "重啟失敗：" + resultRestart.DESCRIPTION;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion
    }
}