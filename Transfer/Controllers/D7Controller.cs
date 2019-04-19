using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class D7Controller : CommonController
    {
        private IA5Repository A5Repository;
        private ID6Repository D6Repository;
        private ID7Repository D7Repository;
        List<SelectOption> selectOptionYN;
        List<SelectOption> selectOptionRange;
        List<SelectOption> selectOptionStatus;
        private string groupProductCode = Assessment_Type.B.GetDescription();

        public D7Controller()
        {
            this.A5Repository = new A5Repository();
            this.D6Repository = new D6Repository();
            this.D7Repository = new D7Repository();

            selectOptionYN = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "Y：是", Value = "Y" },
                                                  new SelectOption() { Text = "N：否", Value = "N" }
                                              };

            selectOptionRange = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "1：以上", Value = "1" },
                                                  new SelectOption() { Text = "0：以下", Value = "0" }
                                              };

            ProcessStatusList psl = new ProcessStatusList();
            selectOptionStatus = psl.statusOption;
            selectOptionStatus.Insert(0, new SelectOption() { Text = "", Value = "" });
        }

        // GET: D7Controller
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// D70-觀察名單參數檔
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D70()
        {
            List<SelectOption> selectOptionCheckPoint = new List<SelectOption>() {
                                                                  new SelectOption() { Text = "", Value = "" },
                                                                  new SelectOption() { Text = "-2：上上個月報導日", Value = "-2" },
                                                                  new SelectOption() { Text = "0：本月報導日", Value = "0" }
                                                                };
            ViewBag.BasicPassCheckPoint = new SelectList(selectOptionCheckPoint, "Value", "Text");

            List<SelectOption> selectOptionBasicPass = new List<SelectOption>() {
                                                       new SelectOption() { Text = "", Value = "" },
                                                       new SelectOption() { Text = "Y：是", Value = "Y" },
                                                       new SelectOption() { Text = "N：否", Value = "N" },
                                                       new SelectOption() { Text = "-999", Value = "-999" }
                                                     };
            ViewBag.BasicPass = new SelectList(selectOptionBasicPass, "Value", "Text");

            ViewBag.RatingCheckPoint = new SelectList(selectOptionCheckPoint, "Value", "Text");

            List<SelectOption> selectOptionA52 = new List<SelectOption>();
            selectOptionA52.Add(new SelectOption() { Text = "", Value = "" });
            selectOptionA52.AddRange((A5Repository.getA52("SP").Item2)
                           .Select(x => new SelectOption(){ Text = x.Rating, Value = x.Rating }));
            selectOptionA52.Add(new SelectOption() { Text = "-999", Value = "-999" });
            ViewBag.A52 = new SelectList(selectOptionA52, "Value", "Text");

            ViewBag.IncludingInd0 = new SelectList(selectOptionYN, "Value", "Text");
            ViewBag.IncludingInd1 = new SelectList(selectOptionYN, "Value", "Text");
            ViewBag.IncludingInd2 = new SelectList(selectOptionYN, "Value", "Text");

            ViewBag.ApplyRange0 = new SelectList(selectOptionRange, "Value", "Text");
            ViewBag.ApplyRange1 = new SelectList(selectOptionRange, "Value", "Text");
            ViewBag.ApplyRange2 = new SelectList(selectOptionRange, "Value", "Text");

            ViewBag.Status = new SelectList(selectOptionStatus, "Value", "Text");

            if (getAssessment(groupProductCode, "D70", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(groupProductCode, "D70", SetAssessmentType.Auditor)
                                 .Select(x => new SelectOption()
                                 { Text = (x.User_Name), Value = x.User_Account }));

            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");
            ViewBag.A51Year = D7Repository.GetA51Year();
            return View();
        }

        /// <summary>
        /// D71-預警名單參數檔
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D71()
        {
            List<SelectOption> selectOptionCheckPoint = new List<SelectOption>() {
                                                                  new SelectOption() { Text = "", Value = "" },
                                                                  new SelectOption() { Text = "-1：上個月報導日", Value = "-1" },
                                                                  new SelectOption() { Text = "0：本月報導日", Value = "0" }
                                                                };
            ViewBag.BasicPassCheckPoint = new SelectList(selectOptionCheckPoint, "Value", "Text");

            List<SelectOption> selectOptionBasicPass = new List<SelectOption>() {
                                                       new SelectOption() { Text = "", Value = "" },
                                                       new SelectOption() { Text = "Y：是", Value = "Y" },
                                                       new SelectOption() { Text = "N：否", Value = "N" },
                                                       new SelectOption() { Text = "-999", Value = "-999" }
                                                     };
            ViewBag.BasicPass = new SelectList(selectOptionBasicPass, "Value", "Text");

            ViewBag.RatingCheckPoint = new SelectList(selectOptionCheckPoint, "Value", "Text");

            List<SelectOption> selectOptionA52 = new List<SelectOption>();
            selectOptionA52.Add(new SelectOption() { Text = "", Value = "" });
            selectOptionA52.AddRange((A5Repository.getA52("SP").Item2)
                           .Select(x => new SelectOption() { Text = x.Rating, Value = x.Rating }));
            selectOptionA52.Add(new SelectOption() { Text = "-999", Value = "-999" });
            ViewBag.A52 = new SelectList(selectOptionA52, "Value", "Text");

            ViewBag.IncludingInd0 = new SelectList(selectOptionYN, "Value", "Text");
            ViewBag.IncludingInd1 = new SelectList(selectOptionYN, "Value", "Text");
            ViewBag.IncludingInd2 = new SelectList(selectOptionYN, "Value", "Text");

            ViewBag.ApplyRange0 = new SelectList(selectOptionRange, "Value", "Text");
            ViewBag.ApplyRange1 = new SelectList(selectOptionRange, "Value", "Text");
            ViewBag.ApplyRange2 = new SelectList(selectOptionRange, "Value", "Text");

            ViewBag.Status = new SelectList(selectOptionStatus, "Value", "Text");

            if (getAssessment(groupProductCode, "D71", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(groupProductCode, "D71", SetAssessmentType.Auditor)
                                 .Select(x => new SelectOption()
                                 { Text = (x.User_Name), Value = x.User_Account }));

            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");
            ViewBag.A51Year = D7Repository.GetA51Year();
            return View();
        }

        /// <summary>
        /// D72-SMF分群設定檔
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D72Detail()
        {
            ViewBag.action = new SelectList(new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }}, "Value", "Text");
            var jqgridInfo = new D72ViewModel().TojqGridData(new int[] {200,250,150,250});
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// 查詢信評檢核結果 
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D73Detail()
        {
            return View();
        }

        [UserAuth]
        public ActionResult Notice()
        {
            ViewBag.NoticeName = new SelectList(D7Repository.GetNoticeName("D74"), "Value", "Text");
            return View();
        }

        #region GetA51
        [HttpPost]
        public JsonResult GetA51Data(string dataYear, string rating, string pdGrade, string ratingAdjust,
                                     string gradeAdjust, string moodysPD)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            try
            {
                var queryData = A5Repository.getA51(Audit_Type.Enable, dataYear, rating, pdGrade, ratingAdjust,
                                                    gradeAdjust, moodysPD);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetA52
        [HttpPost]
        public JsonResult GetA52Data(string ratingOrg, string pdGrade, string rating)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            try
            {
                if (pdGrade.IsNullOrWhiteSpace())
                    pdGrade = "All";
                var queryData = A5Repository.getA52(ratingOrg, pdGrade, rating);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetD70AllData
        [HttpPost]
        public JsonResult GetD70AllData(string type, D70ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            try
            {
                var queryData = D7Repository.getD70All(type);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region GetD70Data
        [HttpPost]
        [BrowserEvent("查詢資料")]
        public JsonResult GetD70Data(string type, D70ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(type);

            try
            {
                var queryData = D7Repository.getD70(type, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region SaveD70
        [BrowserEvent("儲存資料")]
        [HttpPost]
        public JsonResult SaveD70(string type, string actionType, D70ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(type, result.DESCRIPTION);

            try
            {
                result = D7Repository.saveD70(type, actionType, dataModel);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region DeleteD70
        [BrowserEvent("D70/D71 設為失效")]
        [HttpPost]
        public JsonResult DeleteD70(string type, string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(type, result.DESCRIPTION);

            try
            {
                result = D7Repository.deleteD70(type,ruleID);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region SendD70ToAudit
        [BrowserEvent("呈送複核")]
        [HttpPost]
        public JsonResult SendD70ToAudit(string type, string ruleID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(type, result.DESCRIPTION);

            try
            {
                List<D70ViewModel> dataList = new List<D70ViewModel>();

                string[] arrayRuleID = ruleID.Split(',');
                for (int i = 0; i < arrayRuleID.Length; i++)
                {
                    D70ViewModel dataModel = new D70ViewModel();
                    dataModel.Rule_ID = arrayRuleID[i];
                    dataModel.Auditor = auditor;

                    dataList.Add(dataModel);
                }

                result = D7Repository.sendD70ToAudit(type,dataList);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region D70Audit
        [BrowserEvent("複核確認")]
        [HttpPost]
        public JsonResult D70Audit(string type, string ruleID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(type, result.DESCRIPTION);

            try
            {
                List<D70ViewModel> dataList = new List<D70ViewModel>();

                string[] arrayRuleID = ruleID.Split(',');
                for (int i = 0; i < arrayRuleID.Length; i++)
                {
                    D70ViewModel dataModel = new D70ViewModel();
                    dataModel.Rule_ID = arrayRuleID[i];
                    dataModel.Status = status;

                    dataList.Add(dataModel);
                }

                MSGReturnModel resultAudit = D7Repository.D70Audit(type,dataList);

                result.RETURN_FLAG = resultAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.Audit_Success.GetDescription(type);

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(type, resultAudit.DESCRIPTION);
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

        #region D72
        /// <summary>
        /// 查詢D72資料
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢D72資料")]
        [HttpPost]
        public JsonResult GetD72Data()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var Datas = D7Repository.GetD72();
            if (!Datas.Item1.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = Datas.Item1;
            }
            else
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D72DbfileData); //清除
                Cache.Set(CacheList.D72DbfileData, Datas.Item2); //把資料存到 Cache
            }
            return Json(result);
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("D72 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferD72()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                DateTime startTime = DateTime.Now;

                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.D72ExcelName))
                    fileName = (string)Cache.Get(CacheList.D72ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<D72ViewModel> dataModel = new List<D72ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = D7Repository.GetD72Excel(pathType, stream);
                    dataModel = data.Item2;
                    errorMessage = data.Item1;
                }
                if (!errorMessage.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = errorMessage;
                    return Json(result);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.D72TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save SMF_Group(D72)

                MSGReturnModel resultD72 = D7Repository.SaveD72(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.D72,
                    fileName, SetFile.ProgramName, resultD72.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.D72, resultD72.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save SMF_Group(D72)

                result.RETURN_FLAG = resultD72.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.D72.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.D72.ToString(), resultD72.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發(D72)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("D72上傳檔案")]
        [HttpPost]
        public JsonResult UploadD72(ValidateFiles FileModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 前端無傳送檔案進來

                if (FileModel.File == null)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                    return Json(result);
                }

                #endregion 前端無傳送檔案進來

                #region 前端檔案大小不符或不為Excel檔案(驗證)

                //ModelState
                if (!ModelState.IsValid)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不符或不為Excel檔案(驗證)

                #region 上傳檔案

                string pathType = Path.GetExtension(FileModel.File.FileName)
                                       .Substring(1); //上傳的檔案類型

                var fileName = string.Format("{0}.{1}",
                    Excel_UploadName.D72.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.D72ExcelName); //清除 Cache
                Cache.Set(CacheList.D72ExcelName, fileName); //把資料存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);

                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel.File);
                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料
                var data = D7Repository.GetD72Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.D72ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.D72ExcelfileData, data.Item2); //把資料存到 Cache
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = data.Item1;
                }

                #endregion 讀取Excel資料 

                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }
            return Json(result);
        }
        #endregion

        #region D73

        #region 獲取版本
        [HttpPost]
        public JsonResult GetD73Version(string date)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (date.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt))
            {
                return Json(result);
            }
            List<string> data = new List<string>();
            data = D7Repository.GetD73Veriosn(dt);
            if (data.Any())
            {
                data.Insert(0, " ");
                result.RETURN_FLAG = true;
                result.Datas = Json(data);
            }
            return Json(result);
        }
        #endregion

        #region SearchD73
        /// <summary>
        /// 查詢A56資料
        /// </summary>
        /// <param name="IsActive">是否生效 Y,N</param>
        /// <param name="ReplaceObject">特殊字元</param>
        /// <returns></returns>
        [BrowserEvent("查詢D73資料")]
        [HttpPost]
        public JsonResult SearchD73(
            string date,
            string ver
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var datas = D7Repository.GetD73(TypeTransfer.stringToDateTimeN(date), TypeTransfer.stringToIntN(ver));
            Cache.Invalidate(CacheList.D73DbfileData); //清除
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Set(CacheList.D73DbfileData, datas); //把資料存到 Cache
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.D73.tableNameGetDescription());
            }
            return Json(result);
        }
        #endregion

        #region D73GetTxtFile
        [HttpGet]
        public ActionResult D73GetTxtFile(
          string filePath,
          string fileName
          )
        {
            return File(filePath, "application/octet-stream",fileName);
        }
        #endregion

        #region D73GetTxtStr
        [BrowserEvent("檢視D73評估資料")]
        [HttpPost]
        public JsonResult D73GetTxtStr(
    string Id
    )
        {
            MSGReturnModel result = new MSGReturnModel();
            var msg = D7Repository.GetD73FileLog(TypeTransfer.stringToInt(Id));
            if (msg.IsNullOrWhiteSpace())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.already_Change.GetDescription();
            }
            else
            {
                result.RETURN_FLAG = true;
                result.Datas = Json(msg);
            }
            return Json(result);
        }
        #endregion

        #region D73 Del
        /// <summary>
        /// 刪除D73txt檔案
        /// </summary>
        /// <param name="Id"></param>
        /// <param name="date"></param>
        /// <param name="ver"></param>
        /// <returns></returns>
        [BrowserEvent("刪除D73txt檔案")]
        [HttpPost]
        public JsonResult D73DelTxtFile(
            string Id,
            string date,
            string ver
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result = D7Repository.DelD73(Id);
            if (result.RETURN_FLAG)
            {
                var datas = D7Repository.GetD73(TypeTransfer.stringToDateTimeN(date), TypeTransfer.stringToIntN(ver));
                Cache.Invalidate(CacheList.D73DbfileData); //清除
                if (datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Set(CacheList.D73DbfileData, datas); //把資料存到 Cache
                }
            }
            return Json(result);
        }
        #endregion

        #endregion

        #region 通知設定檔系列

        [BrowserEvent("查詢通知設定檔")]
        [HttpPost]
        public JsonResult GetD74(string noticeID)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            var datas = D7Repository.GetD74(noticeID);
            if (datas.Any())
            {
                Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                Cache.Set(CacheList.D74DbfileData, datas); //把資料存到 Cache
                result.RETURN_FLAG = true;
            }
            return Json(result);
        }

        [BrowserEvent("查詢Mail設定檔")]
        [HttpPost]
        public JsonResult GetD74_1(string noticeID)
        {
            MSGReturnModel result = new MSGReturnModel();

            var datas = D7Repository.GetD74_1(noticeID);
            Cache.Invalidate(CacheList.D74_1DbfileData); //清除 Cache
            Cache.Set(CacheList.D74_1DbfileData, datas); //把資料存到 Cache

            return Json(result);
        }

        /// <summary>
        /// 新增 or 修改 D74
        /// </summary>
        /// <param name="action"></param>
        /// <param name="noticeID"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        [BrowserEvent("異動通知設定檔")]
        [HttpPost]
        public JsonResult SaveD74(string action, string noticeID, D74ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            if (action == Action_Type.Add.ToString())
            {
                result = D7Repository.AddD74(data);
                if (result.RETURN_FLAG)
                {
                    var datas = D7Repository.GetD74(noticeID);
                    Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74DbfileData, datas); //把資料存到 Cache
                }
            }
            else if (action == Action_Type.Edit.ToString())
            {
                result = D7Repository.SaveD74(data);
                if (result.RETURN_FLAG)
                {
                    var datas = D7Repository.GetD74(noticeID);
                    Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74DbfileData, datas); //把資料存到 Cache
                }
            }
            return Json(result);
        }

        [BrowserEvent("異動Mail設定檔")]
        [HttpPost]
        public JsonResult SaveD74_1(string action, string noticeID, D74_1ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            if (action == Action_Type.Add.ToString())
            {
                result = D7Repository.AddD74_1(data);
                if (result.RETURN_FLAG)
                {
                    var data1 = D7Repository.GetD74(noticeID);
                    Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74DbfileData, data1); //把資料存到 Cache
                    var data2 = D7Repository.GetD74_1(data.Notice_ID);
                    Cache.Invalidate(CacheList.D74_1DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74_1DbfileData, data2); //把資料存到 Cache
                }
            }
            else if (action == Action_Type.Edit.ToString())
            {
                result = D7Repository.SaveD74_1(data);
                if (result.RETURN_FLAG)
                {
                    var data1 = D7Repository.GetD74(noticeID);
                    Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74DbfileData, data1); //把資料存到 Cache
                    var data2 = D7Repository.GetD74_1(data.Notice_ID);
                    Cache.Invalidate(CacheList.D74_1DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74_1DbfileData, data2); //把資料存到 Cache
                }
            }
            else if (action == Action_Type.Dele.ToString())
            {
                result = D7Repository.DeleD74_1(data);
                if (result.RETURN_FLAG)
                {
                    var data1 = D7Repository.GetD74(noticeID);
                    Cache.Invalidate(CacheList.D74DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74DbfileData, data1); //把資料存到 Cache
                    var data2 = D7Repository.GetD74_1(data.Notice_ID);
                    Cache.Invalidate(CacheList.D74_1DbfileData); //清除 Cache
                    Cache.Set(CacheList.D74_1DbfileData, data2); //把資料存到 Cache
                }
            }
            return Json(result);
        }

        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="type">D74 通知設定,D74_1 郵件設定</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD7CacheData(jqGridParam jdata, string type)
        {
            switch (type)
            {
                case "D74":
                    if (Cache.IsSet(CacheList.D74DbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<D74ViewModel>)Cache.Get(CacheList.D74DbfileData)));
                    break;
                case "D74_1":
                    if (Cache.IsSet(CacheList.D74_1DbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<D74_1ViewModel>)Cache.Get(CacheList.D74_1DbfileData)));
                    break;
            }
            return null;
        }

        #endregion
    }
}