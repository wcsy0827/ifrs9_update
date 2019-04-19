using System.Web.Mvc;
using Transfer.Models;
using Transfer.Utility;
using System.Linq;
using Transfer.Controllers;
using System;
using Newtonsoft.Json;
using static Transfer.Enum.Ref;
using Transfer.ViewModels;

namespace Transfer.Infrastructure
{
    public class BrowserEventAttribute : ActionFilterAttribute
    {
        private string _eventName;
        private IFRS9_Event_Log _eventLog;
        private string _controllerName;
        private string _actionName;
        protected IFRS9DBEntities db
        {
            get;
            private set;
        }

        public BrowserEventAttribute(string eventName)
        {
            _eventName = eventName;
        }

        //在執行 Action 之前執行
        public override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            _controllerName = (string)filterContext.RouteData.Values["controller"];
            _actionName = (string)filterContext.RouteData.Values["action"];
            var loginTime = AccountController.CurrentUserInfo.Ticket.UserData;
            var dt = DateTime.Now.Date;
            db = new IFRS9DBEntities();
            var result = db.IFRS9_Browse_Log.AsNoTracking().Where(x => 
                          x.User_Account == AccountController.CurrentUserInfo.Name &&
                          System.Data.Entity.DbFunctions.TruncateTime(x.Login_Time).Value == dt)
                        .AsEnumerable()
                        .Where(x => x.Login_Time.ToString("yyyy/MM/dd HH:mm:ss") == loginTime)
                        .OrderByDescending(x => x.Browse_Time)
                        .FirstOrDefault();

            if (result != null)
            {
                var _Event_Name_Length = 200;
                var _Event_Name = formatEventName(filterContext, _eventName);
                if(!_Event_Name.IsNullOrWhiteSpace())
                _Event_Name = _Event_Name.Substring(0, _Event_Name.Length >= _Event_Name_Length ? _Event_Name_Length : _Event_Name.Length);
                _eventLog = new IFRS9_Event_Log()
                {
                    Menu_Id = result.Menu_Id,
                    User_Account = result.User_Account,
                    Login_Time = result.Login_Time,
                    Browse_Time = result.Browse_Time,
                    Action_Name = _actionName,
                    Event_Name = _Event_Name,
                    Event_Begin = DateTime.Now
                };
                db.IFRS9_Event_Log.Add(_eventLog);
                try
                {
                    db.SaveChanges();
                }
                catch (Exception ex)
                {
                    ex.exceptionMessage();
                }
            }
            
        }

        //在執行 Action Result 之前執行
        public override void OnResultExecuting(ResultExecutingContext filterContext)
        {
            JsonResult result = filterContext.Result as JsonResult;
            if (_eventLog != null)
            {
                if (_controllerName == "KRisk")
                {
                    KRiskController.KRiskFlowLoaderResult res = result.Data as KRiskController.KRiskFlowLoaderResult;
                    _eventLog.Event_Flag = res.result == "0" ? true : false;
                }
                else
                {
                    MSGReturnModel msg = result.Data as MSGReturnModel;
                    _eventLog.Event_Flag = msg.RETURN_FLAG;
                }
                _eventLog.Event_Complete = DateTime.Now;
                try
                {
                    db.SaveChanges();
                }
                catch { }
                finally
                {
                    db.Dispose();
                }

            }
        }

        private string formatEventName(ActionExecutingContext filterContext,string eventName)
        {
            if (_controllerName == "A4" && _actionName == "TransferToOther")
            {
                var type = filterContext.ActionParameters["type"] as string;
                var debt = filterContext.ActionParameters["debt"] as string;
                return string.Format("{1}{0}({2})", _eventName,
                    type == "All" ? "B01" : type,
                    debt == Debt_Type.M.ToString() ?
                    Debt_Type.M.GetDescription() : Debt_Type.B.GetDescription());
            }
            if (_controllerName == "A5")
            {
                if (_actionName == "TransferA57A58")
                {
                    var date = filterContext.ActionParameters["date"] as string; //report_Date
                    var version = filterContext.ActionParameters["version"] as string; //Version
                    var complement = filterContext.ActionParameters["complement"] as string; //補登信評
                    bool deleteFlag = (bool)filterContext.ActionParameters["deleteFlag"]; //補登信評
                    return string.Format("{0}({1},{2},{3},{4})", _eventName,
                        "報導日:" + date,
                        "版本:" + version,
                        "補登信評:" + complement,
                        "重複執行:" + (deleteFlag ? "是" : "否"));
                }
            }
            if (_controllerName == "A7" && (_actionName == "GetData" || _actionName == "GetExcel"))
            {
                var type = filterContext.ActionParameters["type"] as string;
                return string.Format("{0}({1})", _eventName, type);
            }
            if (_controllerName == "A8" && _actionName == "GetData" )
            {
                var type = filterContext.ActionParameters["type"] as string;
                return string.Format("{0}({1})", _eventName, type);
            }
            if (_controllerName == "C0" && (_actionName == "GetC07Data" || _actionName == "GetExcel"))
            {
                var debtType = filterContext.ActionParameters["debtType"] as string;
                return string.Format("{0}({1})", _eventName,
                    debtType == GroupProductCode.B.ToString() ?
                    Debt_Type.B.GetDescription() : Debt_Type.M.GetDescription());
            }
            if (_controllerName == "D0")
            {
                if (_actionName == "GetD01Data")
                {
                    D01ViewModel dataModel = filterContext.ActionParameters["dataModel"] as D01ViewModel;
                    var debtType = dataModel.DebtType;
                    return string.Format("{0}({1})", _eventName,
                        debtType == GroupProductCode.B.ToString() ?
                        Debt_Type.B.GetDescription() : Debt_Type.M.GetDescription()
                        );
                }
                else if (_actionName == "SaveD01")
                {
                    D01ViewModel dataModel = filterContext.ActionParameters["data"] as D01ViewModel;
                    var debtType = dataModel.DebtType;
                    var actionType = dataModel.ActionType;
                    return string.Format("{1}-{0}({2})", _eventName,
                        actionType == "Add" ? "新增" : "修改",
                        debtType == GroupProductCode.B.ToString() ?
                        Debt_Type.B.GetDescription() : Debt_Type.M.GetDescription()
                        );
                }
                else if (_actionName == "DeleteD01")
                {
                    string prjid = filterContext.ActionParameters["prjid"] as string;
                    string flowid = filterContext.ActionParameters["flowid"] as string;
                    return string.Format("{0}({1})", _eventName,"專案名稱：" + prjid + "，流程名稱：" + flowid);
                }
                else if (_actionName == "GetD05Data" || _actionName == "GetGroupProductData")
                {
                    var debtType = filterContext.ActionParameters["debtType"] as string;
                    return string.Format("{0}({1})", _eventName,
                        debtType == GroupProductCode.B.ToString() ?
                        Debt_Type.B.GetDescription() : Debt_Type.M.GetDescription()
                        );
                }
                else if (_actionName == "SaveD05" || _actionName == "SaveGroupProduct")
                {
                    var debtType = filterContext.ActionParameters["debtType"] as string;
                    var actionType = filterContext.ActionParameters["actionType"] as string;
                    return string.Format("{1}{0}({2})",
                        _eventName,
                        actionType == "Add" ? "新增" : "修改",
                        debtType == GroupProductCode.B.ToString() ?
                        Debt_Type.B.GetDescription() : Debt_Type.M.GetDescription()
                        );
                }
                else if (_actionName == "DeleteD05")
                {
                    string productCode = filterContext.ActionParameters["productCode"] as string;
                    return string.Format("{0}({1})", _eventName, "產品：" + productCode);
                }
                else if (_actionName == "DeleteGroupProduct")
                {
                    string groupProductCode = filterContext.ActionParameters["groupProductCode"] as string;
                    return string.Format("{0}({1})", _eventName, "產品群代碼：" + groupProductCode);
                }
            }
            if (_controllerName == "D6")
            {
                if (_actionName == "SaveD69")
                {
                    D69ViewModel dataModel = filterContext.ActionParameters["data"] as D69ViewModel;
                    var actionType = dataModel.ActionType;
                    return string.Format("{1}-{0}", _eventName,
                        actionType == "Add" ? "新增" : "修改");
                }
                else if (_actionName == "DeleteD69" || _actionName == "SendD69ToAudit")
                {
                    string ruleID = filterContext.ActionParameters["ruleID"] as string;
                    return string.Format("{0}({1})", _eventName, "規則編號：" + ruleID);
                }
                else if (_actionName == "SetReview" || _actionName == "SetReviewVersion")
                {
                    string Reference_Nbr = filterContext.ActionParameters["Reference_Nbr"] as string;
                    string Assessment_Result_Version = filterContext.ActionParameters["Assessment_Result_Version"] as string;
                    return string.Format("{0}({1})", _eventName, "Ref：" + Reference_Nbr+",ARV :"+ Assessment_Result_Version);
                }
            }
            if (_controllerName == "D7")
            {
                if (_actionName == "GetD70Data")
                {
                    var type = filterContext.ActionParameters["type"] as string;
                    return string.Format("{0}({1})", _eventName,type);
                }
                else if (_actionName == "SaveD70")
                {
                    var type = filterContext.ActionParameters["type"] as string;
                    var actionType = filterContext.ActionParameters["actionType"] as string;
                    return string.Format("{1}-{0}({2})", _eventName,actionType == "Add" ? "新增" : "修改",type);
                }
                else if (_actionName == "DeleteD70" || _actionName == "SendD70ToAudit" || _actionName== "D70Audit")
                {
                    var type = filterContext.ActionParameters["type"] as string;
                    string ruleID = filterContext.ActionParameters["ruleID"] as string;
                    return string.Format("{0}-{1}({2})", _eventName,type ,"規則編號：" + ruleID);
                }
            }
            if (_controllerName == "System")
            {
                if (_actionName == "SaveAccount")
                {
                    var action = filterContext.ActionParameters["action"] as string;
                    return string.Format("{0}{1}",
                        action == Action_Type.Add.ToString() ? Action_Type.Add.GetDescription() :
                        action == Action_Type.Edit.ToString() ? Action_Type.Edit.GetDescription() :
                        action == Action_Type.Dele.ToString() ? Action_Type.Dele.GetDescription() :
                         Action_Type.View.ToString(),
                        _eventName
                        );
                }
                if (_actionName == "SearchAccountLog")
                {
                    var type = filterContext.ActionParameters["type"] as string;
                    return string.Format("{0}({1})",
                        _eventName,
                        type == "User" ? "使用者" : type == "Browser" ? "畫面" : "執行動作"
                        );
                }
                if (_actionName == "SearchAssessment")
                {
                    bool assessmentFlag = bool.Parse(filterContext.ActionParameters["assessmentFlag"].ToString());
                    return string.Format("{0}({1})",
                        _eventName,
                        assessmentFlag ? "資料表編號" : "覆核人員/呈送人員"
                        );
                }
                if (_actionName == "AssessmentSetAdd")
                {
                    string type = filterContext.ActionParameters["type"] as string;
                    return string.Format("{0}({1})",
                        _eventName,
                        type == SetAssessmentType.Assessment.ToString() ?
                        "表單" : type == SetAssessmentType.Auditor.ToString() ?
                        SetAssessmentType.Auditor.GetDescription() : SetAssessmentType.Presented.GetDescription()
                        );
                }
                if (_actionName == "AssessmentSetDel")
                {
                    string type = filterContext.ActionParameters["type"] as string;
                    return string.Format("{0}({1})",
                        _eventName,
                        type == SetAssessmentType.Assessment.ToString() ?
                        "表單" : type == SetAssessmentType.Auditor.ToString() ?
                        SetAssessmentType.Auditor.GetDescription() : SetAssessmentType.Presented.GetDescription()
                        );
                }
            }
            if (_controllerName == "Common")
            {
                if (_actionName == "GetReport")
                {
                    var type = filterContext.ActionParameters["data"] as CommonController.reportModel;
                    return string.Format("{0} ({1})",
                        _eventName,
                        type.title );
                }
            }

            #region default
            if (filterContext.ActionParameters.Any())
            {
                System.Collections.Generic.List<string> result = new System.Collections.Generic.List<string>();
                foreach (var item in filterContext.ActionParameters)
                {
                    Type t = item.Value.GetType();
                    if (t.Equals(typeof(string)))
                        result.Add($"{item.Key} : {(string)item.Value}");
                    else if (t.Equals(typeof(int)))
                        result.Add($"{item.Key} : {((int)item.Value).ToString()}");
                    else if (t.Equals(typeof(bool)))
                        result.Add($"{item.Key} : {((bool)item.Value).ToString()}");
                    else if (t.Equals(typeof(double)))
                        result.Add($"{item.Key} : {((double)item.Value).ToString()}");
                    else if (t.Equals(typeof(DateTime)))
                        result.Add($"{item.Key} : {((DateTime)item.Value).ToString("yyyy/MM/dd")}");
                }
                if (result.Any())
                    return $"{eventName}({string.Join(",", result)})";
            }
            #endregion

            return eventName;
        }
    }
}