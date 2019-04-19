using System;
using System.Collections.Generic;
using System.Web.Mvc;
using System.Linq;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class D0Controller : CommonController
    {
        private ID0Repository D0Repository;
        private ID6Repository D6Repository;

        public D0Controller()
        {
            this.D0Repository = new D0Repository();
            this.D6Repository = new D6Repository();
        }

        // GET: D0
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// D03Mortgage(減損階段條件設定)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D03Mortgage()
        {
            ProcessStatusList psl = new ProcessStatusList();
            List<SelectOption> selectOption = psl.statusOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");

            string groupProductCode = Assessment_Type.M.GetDescription();

            if (getAssessment(groupProductCode, "D03", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(groupProductCode, "D03", SetAssessmentType.Auditor)
                        .Select(x => new SelectOption()
                        { Text = x.User_Name, Value = x.User_Account }));
            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            return View();
        }

        /// <summary>
        /// D05(套用產品組合代碼-房貸)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D05Mortgage()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getGroupProductByDebtType(GroupProductCode.M.GetDescription()).Item2)
                .Select(x => new SelectOption()
                { Text = (x.Group_Product_Code + " " + x.Group_Product_Name), Value = x.Group_Product_Code }));

            ViewBag.GroupProduct = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getD05All(GroupProductCode.M.GetDescription()).Item2)
                .Select(x => x.Product_Code).Distinct()
                .Select(x => new SelectOption()
                { Text = (x), Value = x }));

            ViewBag.ProductCode = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        /// <summary>
        /// D05(套用產品組合代碼-債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D05Bond()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getGroupProductByDebtType(GroupProductCode.B.GetDescription()).Item2)
                .Select(x => new SelectOption()
                { Text = (x.Group_Product_Code + " " + x.Group_Product_Name), Value = x.Group_Product_Code }));

            ViewBag.GroupProduct = new SelectList(selectOption, "Value", "Text");


            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getD05All(GroupProductCode.B.GetDescription()).Item2)
                .Select(x => x.Product_Code).Distinct()
                .Select(x => new SelectOption()
                { Text = (x), Value = x }));

            ViewBag.ProductCode = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        /// <summary>
        /// D01Mortgage(套用流程資訊-房貸)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D01Mortgage()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getGroupProductByDebtType(GroupProductCode.M.GetDescription()).Item2)
                .Select(x => new SelectOption()
                { Text = (x.Group_Product_Code + " " + x.Group_Product_Name), Value = x.Group_Product_Code }));

            ViewBag.GroupProduct = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        /// <summary>
        /// D01Bond(套用流程資訊-債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D01Bond()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getGroupProductByDebtType(GroupProductCode.B.GetDescription()).Item2)
                .Select(x => new SelectOption()
                { Text = (x.Group_Product_Code + " " + x.Group_Product_Name), Value = x.Group_Product_Code }));

            ViewBag.GroupProduct = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        /// <summary>
        /// GroupProduct(產品組合-房貸)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult GroupProductMortgage()
        {
            return View();
        }

        /// <summary>
        /// GroupProduct(產品組合-債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult GroupProductBond()
        {
            return View();
        }

        #region GetD03All
        [HttpPost]
        public JsonResult GetD03AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D0Repository.getD03All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D03DbfileData); //清除
                Cache.Set(CacheList.D03DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheD03Data
        [HttpPost]
        public JsonResult GetCacheD03Data(jqGridParam jdata)
        {
            List<D03ViewModel> data = new List<D03ViewModel>();
            if (Cache.IsSet(CacheList.D03DbfileData))
            {
                data = (List<D03ViewModel>)Cache.Get(CacheList.D03DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetD03Data
        [BrowserEvent("查詢D03資料")]
        [HttpPost]
        public JsonResult GetD03Data(D03ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D0Repository.getD03(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D03DbfileData); //清除
                Cache.Set(CacheList.D03DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveD03
        [BrowserEvent("儲存D03資料")]
        [HttpPost]
        public JsonResult SaveD03(string actionType, D03ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = D0Repository.saveD03(actionType, dataModel);

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

        #region DeleteD03
        [BrowserEvent("D03資料設為失效")]
        [HttpPost]
        public JsonResult DeleteD03(string parmID)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = D0Repository.deleteD03(parmID);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = "設為失效成功";

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = "設為失效失敗";
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

        #region SendD03ToAudit
        [BrowserEvent("D03呈送複核")]
        [HttpPost]
        public JsonResult SendD03ToAudit(string parmID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSendToAudit = D0Repository.sendD03ToAudit(parmID, auditor);

                result.RETURN_FLAG = resultSendToAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(resultSendToAudit.DESCRIPTION);
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

        #region D03Audit
        [BrowserEvent("D03 複核確認/退回")]
        [HttpPost]
        public JsonResult D03Audit(string parmID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultAudit = D0Repository.D03Audit(parmID, status);

                result.RETURN_FLAG = resultAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(resultAudit.DESCRIPTION);
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

        #region Get Data

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD05AllData(string debtType)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            if (GroupProductCode.M.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var D05Data = D0Repository.getD05All(debtType);
                result.RETURN_FLAG = D05Data.Item1;
                result.Datas = Json(D05Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢D05資料")]
        [HttpPost]
        public JsonResult GetD05Data(string debtType, string groupProductCode, string productCode, string processingDate1, string processingDate2)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D05");

            if (GroupProductCode.M.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var D05Data = D0Repository.getD05(debtType, groupProductCode, productCode, processingDate1, processingDate2);
                result.RETURN_FLAG = D05Data.Item1;
                result.Datas = Json(D05Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion Get Data

        /// <summary>
        /// 新增、俢改
        /// </summary>
        /// <param name="debtType"></param>
        /// <param name="actionType">動作類型(Add Or Modify)</param>
        /// <param name="groupProductCode">套用產品群代碼</param>
        /// <param name="groupProduct">產品群別說明</param>
        /// <param name="productCode">產品</param>
        /// <returns></returns>
        [BrowserEvent("D05資料")]
        [HttpPost]
        public JsonResult SaveD05(string debtType, string actionType, string groupProductCode, string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D05", result.DESCRIPTION);

            try
            {
                if (GroupProductCode.M.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.M.GetDescription();
                }
                else if (GroupProductCode.B.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.B.GetDescription();
                }

                D05ViewModel dataModel = new D05ViewModel();
                dataModel.Group_Product_Code = groupProductCode;
                dataModel.Product_Code = productCode;
                dataModel.Processing_Date = DateTime.Now.ToString("yyyy/MM/dd");

                MSGReturnModel resultSave = D0Repository.saveD05(debtType, actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription("D05");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription("D05", resultSave.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="productCode">產品</param>
        /// <returns></returns>
        [BrowserEvent("刪除D05資料")]
        [HttpPost]
        public JsonResult DeleteD05(string groupProductCode , string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D05", result.DESCRIPTION);

            try
            {
                MSGReturnModel resultDelete = D0Repository.deleteD05(groupProductCode,productCode);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription("D05");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription("D05", resultDelete.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #region Get Data

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="debtType">1.房貸 4.債券</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetGroupProductByDebtType(string debtType)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            if (GroupProductCode.M.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var returnData = D0Repository.getGroupProductByDebtType(debtType);
                result.RETURN_FLAG = returnData.Item1;
                result.Datas = Json(returnData.Item2);

                if (result.RETURN_FLAG == false)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("套用產品群代碼");
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD01AllData(D01ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            if (GroupProductCode.M.ToString().Equals(dataModel.DebtType))
            {
                dataModel.DebtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(dataModel.DebtType))
            {
                dataModel.DebtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var D01Data = D0Repository.getD01All(dataModel.DebtType);
                result.RETURN_FLAG = D01Data.Item1;
                result.Datas = Json(D01Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [BrowserEvent("查詢D01資料")]
        [HttpPost]
        public JsonResult GetD01Data(D01ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D01");

            if (GroupProductCode.M.ToString().Equals(dataModel.DebtType))
            {
                dataModel.DebtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(dataModel.DebtType))
            {
                dataModel.DebtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var D01Data = D0Repository.getD01(dataModel);
                result.RETURN_FLAG = D01Data.Item1;
                result.Datas = Json(D01Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion Get Data

        #region SaveD01

        /// <summary>
        /// 新增、俢改
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [BrowserEvent("儲存D01資料")]
        [HttpPost]
        public JsonResult SaveD01(D01ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D01", result.DESCRIPTION);

            try
            {
                D01ViewModel dataModel = new D01ViewModel();
                dataModel.DebtType = data.DebtType;
                dataModel.ActionType = data.ActionType;
                dataModel.PRJID = data.PRJID;
                dataModel.FLOWID = data.FLOWID;
                dataModel.Group_Product_Code = data.Group_Product_Code;
                dataModel.Publish_Date = DateTime.Now.ToString("yyyy/MM/dd");
                dataModel.Apply_On_Date = data.Apply_On_Date;
                dataModel.Apply_Off_Date = data.Apply_Off_Date;
                dataModel.Issuer = AccountController.CurrentUserInfo.Name;
                dataModel.Memo = data.Memo;

                MSGReturnModel resultSave = D0Repository.saveD01(dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription("D01");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription("D01", resultSave.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion SaveD01

        #region DeleteD01

        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="prjid">專案名稱</param>
        /// <param name="flowid">流程名稱</param>
        /// <returns></returns>
        [BrowserEvent("刪除D01資料")]
        [HttpPost]
        public JsonResult DeleteD01(string prjid, string flowid)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D01", result.DESCRIPTION);

            try
            {
                MSGReturnModel resultDelete = D0Repository.deleteD01(prjid, flowid);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription("D01");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription("D01", resultDelete.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion DeleteD01

        #region GetGroupProductAllData

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetGroupProductAllData(string debtType)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            if (GroupProductCode.M.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var GroupProductData = D0Repository.getGroupProductAll(debtType);
                result.RETURN_FLAG = GroupProductData.Item1;
                result.Datas = Json(GroupProductData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion GetGroupProductAllData

        #region GetGroupProductData

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢產品群代碼資料")]
        [HttpPost]
        public JsonResult GetGroupProductData(string debtType, string groupProductCode, string groupProductName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("GroupProduct");

            if (GroupProductCode.M.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.M.GetDescription();
            }
            else if (GroupProductCode.B.ToString().Equals(debtType))
            {
                debtType = GroupProductCode.B.GetDescription();
            }

            try
            {
                var GroupProductData = D0Repository.getGroupProduct(debtType, groupProductCode, groupProductName);
                result.RETURN_FLAG = GroupProductData.Item1;
                result.Datas = Json(GroupProductData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion GetGroupProductData

        #region SaveGroupProduct

        /// <summary>
        /// 新增、俢改
        /// </summary>
        /// <param name="debtType"></param>
        /// <param name="actionType">動作類型(Add Or Modify)</param>
        /// <param name="groupProductCode">產品群代碼</param>
        /// <param name="groupProductName">產品群名稱</param>
        /// <returns></returns>
        [BrowserEvent("產品群代碼資料")]
        [HttpPost]
        public JsonResult SaveGroupProduct(string debtType, string actionType, string groupProductCode, string groupProductName)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("GroupProduct", result.DESCRIPTION);

            try
            {
                if (GroupProductCode.M.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.M.GetDescription();
                }
                else if (GroupProductCode.B.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.B.GetDescription();
                }

                GroupProductViewModel dataModel = new GroupProductViewModel();
                dataModel.Group_Product_Code = groupProductCode;
                dataModel.Group_Product_Name = groupProductName;

                MSGReturnModel resultSave = D0Repository.saveGroupProduct(debtType, actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription("GroupProduct");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription("GroupProduct", resultSave.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion SaveGroupProduct

        #region DeleteGroupProduct

        /// <summary>
        /// 刪除
        /// </summary>
        /// <param name="groupProductCode">產品群代碼</param>
        /// <returns></returns>
        [BrowserEvent("刪除產品群代碼資料")]
        [HttpPost]
        public JsonResult DeleteGroupProduct(string groupProductCode)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("GroupProduct", result.DESCRIPTION);

            try
            {
                MSGReturnModel resultDelete = D0Repository.deleteGroupProduct(groupProductCode);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription("GroupProduct");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription("GroupProduct", resultDelete.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion DeleteGroupProduct

        #region GetD02Data
        [BrowserEvent("查詢D02資料")]
        [HttpPost]
        public JsonResult GetD02Data(D02ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = D0Repository.getD02(dataModel);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.D02DbfileData); //清除
                Cache.Set(CacheList.D02DbfileData, queryData.Item2); //把資料存到 Cache
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

        #region GetCacheD02Data
        [HttpPost]
        public JsonResult GetCacheD02Data(jqGridParam jdata)
        {
            List<D02ViewModel> data = new List<D02ViewModel>();
            if (Cache.IsSet(CacheList.D02DbfileData))
            {
                data = (List<D02ViewModel>)Cache.Get(CacheList.D02DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true));
        }
        #endregion

        #region SaveD02
        [BrowserEvent("D02轉檔")]
        [HttpPost]
        public JsonResult SaveD02(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = D0Repository.saveD02(reportDate);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = resultSave.DESCRIPTION;
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