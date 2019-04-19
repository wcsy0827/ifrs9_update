using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using System.Linq;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class A9Controller : CommonController
    {
        private IA9Repository A9Repository;
        private List<SelectOption> actions = null;
        DateTime startTime = DateTime.MinValue;
        public A9Controller()
        {
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }};
            this.A9Repository = new A9Repository();
            startTime = DateTime.Now;
        }

        /// <summary>
        /// A94(主權債測試指標_Ticker)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A94()
        {
            return View();
        }

        /// <summary>
        /// A95_1(主權債測試指標_Ticker)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A95_1Detail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A95_1).TojqGridData(new int[] {100,150,230,190,165,100 },null,true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        public ActionResult A96Detail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A96).TojqGridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        public ActionResult A96DateDetail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A96_Trade).TojqGridData(null, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        #region GetA94All
        [HttpPost]
        public JsonResult GetA94AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A9Repository.getA94All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A94DbfileData); //清除
                Cache.Set(CacheList.A94DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheA94Data
        [HttpPost]
        public JsonResult GetCacheA94Data(jqGridParam jdata)
        {
            List<A94ViewModel> data = new List<A94ViewModel>();
            if (Cache.IsSet(CacheList.A94DbfileData))
            {
                data = (List<A94ViewModel>)Cache.Get(CacheList.A94DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetA94Data
        [BrowserEvent("查詢A94資料")]
        [HttpPost]
        public JsonResult GetA94Data(A94ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A9Repository.getA94(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A94DbfileData); //清除
                Cache.Set(CacheList.A94DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveA94
        [BrowserEvent("儲存A94資料")]
        [HttpPost]
        public JsonResult SaveA94(string actionType, A94ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = A9Repository.saveA94(actionType, dataModel);

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

        #region DeleteA94
        [BrowserEvent("A94刪除資料")]
        [HttpPost]
        public JsonResult DeleteA94(string country)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = A9Repository.deleteA94(country);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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

        #region A95_1

        /// <summary>
        /// 查詢 A95_1(產業別對應Ticker補登檔) 資料
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢 A95_1(產業別對應Ticker補登檔) 資料")]
        [HttpPost]
        public JsonResult GetA95_1Data(string bondNumber)
        {
            MSGReturnModel result = new MSGReturnModel();
            var datas = A9Repository.getA95_1(bondNumber);
            if (datas.Any())
            {
                Cache.Invalidate(CacheList.A95_1DbfileData); //清除
                Cache.Set(CacheList.A95_1DbfileData, datas); //把資料存到 Cache
                result.RETURN_FLAG = true;                
            }
            else
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A95_1.tableNameGetDescription());
            }
            return Json(result);
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A95_1(產業別對應Ticker補登檔)Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA95_1()
        {
            MSGReturnModel resultA95_1 = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A95_1ExcelName))
                    fileName = (string)Cache.Get(CacheList.A95_1ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    resultA95_1.RETURN_FLAG = false;
                    resultA95_1.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(resultA95_1);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A95_1ViewModel> dataModel = new List<A95_1ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A9Repository.getA95_1Excel(pathType, stream);
                    dataModel = data.Item2;
                    errorMessage = data.Item1;
                }
                if (!errorMessage.IsNullOrWhiteSpace())
                {
                    resultA95_1.RETURN_FLAG = false;
                    resultA95_1.DESCRIPTION = errorMessage;
                    return Json(resultA95_1);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A95_1TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save Assessment_Sub_Kind_Ticker(A95_1)

                resultA95_1 = A9Repository.insertA95_1(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A95_1,
                    fileName, SetFile.ProgramName, resultA95_1.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A95_1, resultA95_1.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log
                #endregion save Assessment_Sub_Kind_Ticker(A95_1)
            }
            catch (Exception ex)
            {
                resultA95_1.RETURN_FLAG = false;
                resultA95_1.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(resultA95_1);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發(A95)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("A95_1(產業別對應Ticker補登檔)上傳檔案")]
        [HttpPost]
        public JsonResult UploadA95_1(ValidateFiles FileModel)
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
                    Excel_UploadName.A95_1.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A95_1ExcelName); //清除 Cache
                Cache.Set(CacheList.A95_1ExcelName, fileName); //把資料存到 Cache

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
                var data = A9Repository.getA95_1Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A95_1ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A95_1ExcelfileData, data.Item2); //把資料存到 Cache
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

        /// <summary>
        /// 修改 A95_1 資料
        /// </summary>
        /// <param name="data">該筆資料</param>
        /// <param name="action">修改狀態</param>
        /// <param name="bondNumber">查詢bondNumber</param>
        /// <returns></returns>
        [BrowserEvent("修改A95_1(產業別對應Ticker補登檔)檔案")]
        [HttpPost]
        public JsonResult UpdateA95_1(A95_1ViewModel data, string action,string bondNumber)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            bool breakFlag = false;
            EnumUtil.GetValues<Action_Type>().ToList()
                .ForEach(x =>
                {
                    if (!breakFlag && action == x.ToString())
                    {
                        result = A9Repository.saveA95_1(data, x);
                        breakFlag = true;
                    }
                });
            if (result.RETURN_FLAG)
            {
                var datas = A9Repository.getA95_1(bondNumber);
                Cache.Invalidate(CacheList.A95_1DbfileData); //清除 Cache
                Cache.Set(CacheList.A95_1DbfileData, datas); //把資料存到 Cache            
            }
            return Json(result);
        }


        /// <summary>
        /// 下載A95_1(產業別對應Ticker補登檔)Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A95_1(產業別對應Ticker補登檔)Excel檔案")]
        [HttpPost]
        public JsonResult GetA95_1Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.A95_1DbfileData))
            {
                var A95_1 = Excel_DownloadName.A95_1.ToString();
                var A95_1Data = (List<A95_1ViewModel>)Cache.Get(CacheList.A95_1DbfileData);  //從Cache 抓資料
                result = A9Repository.DownLoadExcel(Excel_DownloadName.A95_1, ExcelLocation(A95_1.GetExelName()), A95_1Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }
        #endregion

        #region A96 系列

        #region A96
        [BrowserEvent("查詢A96資料")]
        [HttpPost]
        public JsonResult GetA96(string reportDate,int version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            DateTime dt = DateTime.MinValue;
            if (reportDate.IsNullOrWhiteSpace() || !DateTime.TryParse(reportDate, out dt))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                var datas = A9Repository.getA96(dt,version);
                if (datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A96DbfileData); //清除
                    Cache.Set(CacheList.A96DbfileData, datas); //把資料存到 Cache
                }
            }
            
            return Json(result);    
        }

        [HttpPost]
        public JsonResult GetA96Version(string date, string tableName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (!DateTime.TryParse(date, out dt))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                List<string> data = new List<string>();
                data = A9Repository.getA96Version(dt);
                if (data.Any())
                {
                    data.Insert(0, " ");
                    result.RETURN_FLAG = true;
                    result.Datas = Json(data);
                }
            }
            return Json(result);
        }


        /// <summary>
        /// 選擇檔案後點選資料上傳觸發(A96)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("A96上傳檔案")]
        [HttpPost]
        public JsonResult UploadA96(ValidateFiles FileModel)
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
                    Excel_UploadName.A96.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A96ExcelName); //清除 Cache
                Cache.Set(CacheList.A96ExcelName, fileName); //把資料存到 Cache

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
                var data = A9Repository.getA96Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A96ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A96ExcelfileData, data.Item2); //把資料存到 Cache
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

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A96 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA96()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A96ExcelName))
                    fileName = (string)Cache.Get(CacheList.A96ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A96ViewModel> dataModel = new List<A96ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A9Repository.getA96Excel(pathType, stream);
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

                string txtpath = SetFile.A96TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save Bond_Spread_Info(A96)

                MSGReturnModel resultA96 = A9Repository.saveA96(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A96,
                    fileName, SetFile.ProgramName, resultA96.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A96, resultA96.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Bond_Spread_Info(A96)

                result.RETURN_FLAG = resultA96.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A96.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A96.ToString(), resultA96.DESCRIPTION);
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
        ///下載A96Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A96Excel檔案")]
        [HttpPost]
        public JsonResult GetA96Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (Cache.IsSet(CacheList.A96DbfileData))
            {
                var A96 = Excel_DownloadName.A96.ToString();
                var A96Data = (List<A96ViewModel>)Cache.Get(CacheList.A96DbfileData);  //從Cache 抓資料
                result = A9Repository.DownLoadExcel(Excel_DownloadName.A96, ExcelLocation(A96.GetExelName()), A96Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }

        #endregion

        #region A96 最後交易日
        [BrowserEvent("查詢A96交易日")]
        [HttpPost]
        public JsonResult GetA96Trade(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            DateTime? dt = null;
            DateTime _dt = DateTime.MinValue;
            if (DateTime.TryParse(reportDate, out _dt))
            {
                dt = _dt;
            }
            var datas = A9Repository.getA96Trade(dt);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.A96TradeDbfileData); //清除
                Cache.Set(CacheList.A96TradeDbfileData, datas); //把資料存到 Cache
            }           
            return Json(result);
        }
        #endregion

        [HttpPost]
        public JsonResult saveA96Trade(A96TradeViewModel data, string action, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            Action_Type _action = EnumUtil.GetValues<Action_Type>()
                       .FirstOrDefault(x => x.ToString() == action);
            DateTime? dt = null;
            DateTime _dt = DateTime.MinValue;
            if (DateTime.TryParse(reportDate, out _dt))
            {
                dt = _dt;
            }
            result = A9Repository.saveA96Trade(data, _action);
            if (result.RETURN_FLAG)
            {
                var datas = A9Repository.getA96Trade(dt);
                Cache.Invalidate(CacheList.A96TradeDbfileData); //清除
                Cache.Set(CacheList.A96TradeDbfileData, datas); //把資料存到 Cache              
            }
            return Json(result);
        }

        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata,string type)
        {
            switch (type)
            {
                case "A96DB":
                    return Json(jdata.modelToJqgridResult((List<A96ViewModel>)Cache.Get(CacheList.A96DbfileData)));
                case "A96Excel":
                    return Json(jdata.modelToJqgridResult((List<A96ViewModel>)Cache.Get(CacheList.A96ExcelfileData))); 
                case "A96Trade":
                    return Json(jdata.modelToJqgridResult((List<A96TradeViewModel>)Cache.Get(CacheList.A96TradeDbfileData)));
            }
            return null;
        }

        #endregion

    }
}