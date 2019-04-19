using System;
using System.Collections.Generic;
using System.Configuration;
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
    public class A6Controller : CommonController
    {
        private IA6Repository A6Repository;

        public A6Controller()
        {
            this.A6Repository = new A6Repository();
        }

        /// <summary>
        /// 查詢A62 (違約損失資料檔_歷史資料 Moody_LGD_Info)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A62Detail()
        {
            //StatusList statusList = new StatusList();
            //List<SelectOption> selectOption = statusList.statusOption.Where(x=>x.Value != "").ToList();
            //ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            ViewBag.DataYear = new SelectList(A6Repository.GetA62SearchYear()
                                                          .Select(x => new { Text = x, Value = x }), "Value", "Text");

            return View();
        }

        /// <summary>
        /// A6 (上傳檔案)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// A63 (上傳檔案)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A63()
        {
            return View();
        }

        /// <summary>
        /// A62 (違約損失資料率複核)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A62Audit()
        {
            StatusList statusList = new StatusList(){ };
            List<SelectOption> selectOption = statusList.statusOption.Where(x => x.Value != "3").ToList();
            selectOption.Insert(0, new SelectOption() { Text = "全部", Value = "All" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            statusList = new StatusList();
            selectOption = statusList.statusOption.Where(x => x.Value == "1" || x.Value == "2").ToList();
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.AuditStatus = new SelectList(selectOption, "Value", "Text");

            ViewBag.DataYear = new SelectList(A6Repository.GetA62SearchYear().Where(x=> !x.StartsWith("All"))
                                           .Select(x => new { Text = x, Value = x }), "Value", "Text");

            ViewBag.AssessorFlag = getAssessment(Assessment_Type.B.GetDescription(), Table_Type.A62.ToString(),
            SetAssessmentType.Auditor).Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name));
            return View();
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A62資料")]
        [HttpPost]
        public JsonResult GetA62Data(string dataYear)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            try
            {
                var A62 = A6Repository.GetA62(dataYear);
                result.RETURN_FLAG = A62.Item1;
                var jqgridParams = new A62ViewModel().TojqGridData();
                jqgridParams.Datas = A62.Item2;
                result.Datas = Json(jqgridParams);
                if (!result.RETURN_FLAG)
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #region GetA62Excel
        [BrowserEvent("A62匯出Excel")]
        [HttpPost]
        public ActionResult GetA62Excel(string dataYear)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                var data = A6Repository.GetA62(dataYear);

                path = "A62".GetExelName();
                result = A6Repository.DownLoadA62Excel(ExcelLocation(path), data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail.GetDescription(ex.Message);
            }

            return Json(result);
        }
        #endregion

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("Exhibit 7Excel檔存到DB")]
        [HttpPost]
        public JsonResult Transfer()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                DateTime startTime = DateTime.Now;
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A62ExcelName))
                    fileName = (string)Cache.Get(CacheList.A62ExcelName);  //從Cache 抓資料
                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                List<Exhibit7Model> dataModel = new List<Exhibit7Model>();
                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    string pathType = path.Split('.')[1]; //抓副檔名
                    dataModel = A6Repository.getExcel(pathType, stream); //Excel轉成 Exhibit7Model
                }

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A62TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA6Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save 資料

                #region save Tm_Adjust_YYYY(A62)

                MSGReturnModel resultA62 = A6Repository.saveA62(dataModel); //save to DB
                bool A62Log = CommonFunction.saveLog(Table_Type.A62, fileName, SetFile.ProgramName,
                    resultA62.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,0); //寫sql Log
                TxtLog.txtLog(Table_Type.A62, resultA62.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Tm_Adjust_YYYY(A62)

                result = resultA62;

                #endregion save 資料
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.save_Fail
                    .GetDescription(null, ex.Message);
            }
            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("上傳Exhibit 7Excel檔案")]
        [HttpPost]
        public JsonResult Upload(ValidateFiles FileModel)
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

                #region 前端檔案大小不服或不為Excel檔案(驗證)

                if (!ModelState.IsValid)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不服或不為Excel檔案(驗證)

                #region 上傳檔案

                string pathType = Path.GetExtension(FileModel.File.FileName)
                                      .Substring(1); //上傳的檔案類型

                var fileName = string.Format("{0}.{1}",
                    Excel_UploadName.A62.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A62ExcelName); //清除 Cache
                Cache.Set(CacheList.A62ExcelName, fileName); //把資料存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);
                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel.File);
                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料 使用ExcelDataReader 並且組成 json

                var stream = FileModel.File.InputStream;
                List<Exhibit7Model> dataModel = A6Repository.getExcel(pathType, stream);
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    result.Datas = Json(dataModel); //給JqGrid 顯示
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.data_Not_Compare.GetDescription();
                }

                #endregion 讀取Excel資料 使用ExcelDataReader 並且組成 json

                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.upload_Fail
                    .GetDescription(FileModel.File.FileName, ex.Message);
            }
            return Json(result);
        }

        #region 上傳A63檔案
        [BrowserEvent("上傳A63檔案")]
        [HttpPost]
        public JsonResult UploadA63(ValidateFiles FileModel)
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

                #region 前端檔案大小不服或不為Excel檔案(驗證)

                if (!ModelState.IsValid)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不服或不為Excel檔案(驗證)

                #region 上傳檔案

                string pathType = Path.GetExtension(FileModel.File.FileName)
                                      .Substring(1); //上傳的檔案類型

                var fileName = string.Format("{0}.{1}",Excel_UploadName.A63.GetDescription(),pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A63ExcelName); //清除 Cache
                Cache.Set(CacheList.A63ExcelName, fileName); //把資料存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);
                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel.File);

                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料 使用ExcelDataReader 並且組成 json

                var stream = FileModel.File.InputStream;
                List<A63ViewModel> dataModel = A6Repository.getA63Excel(pathType, stream);
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    result.Datas = Json(dataModel); //給JqGrid 顯示
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.data_Not_Compare.GetDescription();
                }

                #endregion 讀取Excel資料 使用ExcelDataReader 並且組成 json

                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.upload_Fail.GetDescription(FileModel.File.FileName, ex.Message);
            }
            return Json(result);
        }
        #endregion

        #region TransferA63
        [BrowserEvent("Exhibit 8Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA63()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                DateTime startTime = DateTime.Now;
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;

                if (Cache.IsSet(CacheList.A63ExcelName))
                    fileName = (string)Cache.Get(CacheList.A63ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                List<A63ViewModel> dataModel = new List<A63ViewModel>();
                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    string pathType = path.Split('.')[1]; //抓副檔名
                    dataModel = A6Repository.getA63Excel(pathType, stream);
                }

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A63TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA63Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save
                MSGReturnModel resultA63 = A6Repository.saveA63(dataModel); //save to DB
                bool A63Log = CommonFunction.saveLog(Table_Type.A63, fileName, SetFile.ProgramName,
                    resultA63.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, 0); //寫sql Log
                TxtLog.txtLog(Table_Type.A63, resultA63.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log
                #endregion

                result = resultA63;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.save_Fail.GetDescription(null, ex.Message);
            }

            return Json(result);
        }
        #endregion

        #region A62Audit
        [BrowserEvent("A62複核")]
        [HttpPost]
        public JsonResult A62Audit(List<A62ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultAudit = new MSGReturnModel();
                resultAudit = A6Repository.A62Audit(dataModel);

                result.RETURN_FLAG = resultAudit.RETURN_FLAG;
                result.DESCRIPTION = resultAudit.DESCRIPTION;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        [HttpPost]
        public JsonResult A62YearByStatus(string Status)
        {
            if (Status == "")
                Status = null;
            return Json(A6Repository.GetA62SearchYear(Status)
                .Select(x => new SelectOption() { Text = x, Value = x }).Skip(1));
        }
    }
}