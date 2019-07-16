using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class C0Controller : CommonController
    {
        private IKriskRepository KriskRepository;
        private IC0Repository C0Repository;
        private List<SelectOption> actions = null;
        DateTime startTime = DateTime.MinValue;

        protected Common common
        {
            get;
            private set;
        }

        public C0Controller()
        {
            this.KriskRepository = new KriskRepository();
            this.C0Repository = new C0Repository();
            startTime = DateTime.Now;
            this.common = new Common();
        }

        // GET: C0
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// C01(減損計算資料上傳-債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C01Bond()
        {
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            List<SelectOption> listProductCode = new List<SelectOption>() {
                new SelectOption() {Text="香港",Value="HK" },
                new SelectOption() {Text="越南",Value="VN" }};
            ViewBag.ProductCode = new SelectList(listProductCode, "Value", "Text");

            List<string> listVersion = C0Repository.getC01Version("", "HK");
            listVersion.Insert(0, "");
            ViewBag.Version = new SelectList(listVersion.Select(x => new { Text = x, Value = x }), "Value", "Text");

            int[] widths = new int[] { 150, 120, 150, 150, 100, 150, 150, 150, 150, 150,
                                       150, 150, 100, 100, 100, 150, 150, 150, 180 };
            string[] aligns = new string[] { "left", "left", "left", "left", "left", "right", "right", "right", "right", "right",
                                             "right", "right", "right", "right", "right", "right", "right", "right", "right"};
            var jqgridInfo = new C01ViewModel().TojqGridData(widths, aligns);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;

            return View();
        }

        /// <summary>
        /// C07(減損計算輸出資料-債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C07Bond()
        {
            ViewBag.ProductCode = new SelectList(C0Repository.getProduct(GroupProductCode.B), "Value", "Text");

            List<string> listVersion = C0Repository.getC07Version(Transfer.Enum.Ref.GroupProductCode.B.GetDescription(), "", "");
            listVersion.Insert(0, "");

            ViewBag.Version = new SelectList(listVersion.Select(x => new { Text = x, Value = x }), "Value", "Text");

            return View();
        }

        /// <summary>
        /// C07(減損計算輸出資料-房貸)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C07Mortgage()
        {
            ViewBag.ProductCode = new SelectList(C0Repository.getProduct(GroupProductCode.M), "Value", "Text");

            List<string> listVersion = C0Repository.getC07Version(Transfer.Enum.Ref.GroupProductCode.M.GetDescription(), "", "");
            listVersion.Insert(0, "");

            ViewBag.Version = new SelectList(listVersion.Select(x => new { Text = x, Value = x }), "Value", "Text");

            return View();
        }

        /// <summary>
        /// C07Advanced(債券減損計算結果進階查詢)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C07Advanced()
        {
            ViewBag.ProductCode = new SelectList(C0Repository.getProduct(GroupProductCode.B), "Value", "Text");

            List<string> listVersion = C0Repository.getC07Version(Transfer.Enum.Ref.GroupProductCode.B.GetDescription(), "", "");
            listVersion.Insert(0, "");
            ViewBag.Version = new SelectList(listVersion.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<string> listAssessment_Sub_Kind = C0Repository.getA41AdvancedAssessment_Sub_Kind();
            listAssessment_Sub_Kind.Insert(0, "");
            ViewBag.Assessment_Sub_Kind = new SelectList(listAssessment_Sub_Kind.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<SelectOption> listWatch_IND = new List<SelectOption>() {
                new SelectOption() {Text="", Value="" },
                new SelectOption() {Text="Y：是",Value="Y" },
                new SelectOption() {Text="N：否",Value="N" },
                new SelectOption() {Text="FVPL:不適用",Value = "不適用"}
            };
            ViewBag.Watch_IND = new SelectList(listWatch_IND, "Value", "Text");

            return View();
        }

        /// <summary>
        /// C04
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C04()
        {
            var _DateRange = C0Repository.GetC04YearQuartly();
            ViewBag.DateRange = _DateRange;
            var _DateRangeSelect = new List<string>() { " " };
            _DateRange.ForEach(x =>
            {
                _DateRangeSelect.Add(x);
            });
            ViewBag.SelectDateRange = new SelectList(_DateRangeSelect
                .Select(x => new { Text = x, Value = x }), "Value", "Text");
            var jqgridInfo = new C04ViewModel().TojqGridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        public ActionResult C10()
        {
            var jqgridInfo = C10GridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        #region C10Detail
        [UserAuth]
        public ActionResult C10Detail()
        {
            var jqgridInfo = C10DetailGridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }
        #endregion


        [UserAuth]
        public ActionResult C04_1()
        {
            return View();
        }

        #region 轉檔把Excel 資料存到 DB
        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("C10 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferC10(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.C10ExcelName))
                    fileName = (string)Cache.Get(CacheList.C10ExcelName);  //從Cache 抓資料Name

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<C10ViewModel> dataModel = new List<C10ViewModel>();

                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(reportDate, out dt);

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = C0Repository.getExcel(pathType, stream, dt);
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

                
                string txtpath = SetFile.C10TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogC0Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代
                

                #endregion txtlog 檔案名稱

                #region save C10



                MSGReturnModel resultC10 = C0Repository.saveC10(dataModel, reportDate); //save to DB

                int v = 0; ///上傳檔案，版本皆為0
                
                bool C10Log = CommonFunction.saveLog(Table_Type.C10,
                    fileName, SetFile.ProgramName, resultC10.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, v, dt); //寫sql Log
                TxtLog.txtLog(Table_Type.C10, resultC10.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log
                
                #endregion save C10

                result.RETURN_FLAG = resultC10.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.C10.ToString());

                if (!result.RETURN_FLAG)
                {
                    
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.C10.ToString(), resultC10.DESCRIPTION);
                        
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

        #region GetC01Version
        [HttpPost]
        public JsonResult GetC01Version(string reportDate, string productCode)
        {
            List<string> versionData = C0Repository.getC01Version(reportDate, productCode);
            return Json(string.Join(",", versionData));
        }
        #endregion

        #region GetC01LogData
        [HttpPost]
        public JsonResult GetC01LogData(string debt)
        {
            string tableType = "C01_HK,C01_VN";
            List<string> logDatas = C0Repository.GetC01LogData(tableType, debt);
            return Json(string.Join(",", logDatas));
        }
        #endregion

        #region GetC01Data
        [BrowserEvent("查詢C01資料")]
        [HttpPost]
        public JsonResult GetC01Data(string reportDate, string productCode, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = C0Repository.getC01(reportDate, productCode, version);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.C01DbfileData); //清除
                Cache.Set(CacheList.C01DbfileData, queryData.Item2); //把資料存到 Cache
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

        #region GetCacheC01Data
        [HttpPost]
        public JsonResult GetCacheC01Data(jqGridParam jdata, string type)
        {
            List<C01ViewModel> data = new List<C01ViewModel>();

            switch (type)
            {
                case "Excel":
                    if (Cache.IsSet(CacheList.C01ExcelfileData))
                        data = (List<C01ViewModel>)Cache.Get(CacheList.C01ExcelfileData);  //從Cache 抓資料
                    break;

                case "Db":
                    if (Cache.IsSet(CacheList.C01DbfileData))
                        data = (List<C01ViewModel>)Cache.Get(CacheList.C01DbfileData);
                    break;
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr" }));
        }
        #endregion

        #region UploadC01
        [BrowserEvent("C01上傳檔案")]
        [HttpPost]
        public JsonResult UploadC01()
        {
            MSGReturnModel result = new MSGReturnModel();

            //## 如果有任何檔案類型才做
            if (Request.Files.AllKeys.Any())
            {
                var FileModel = Request.Files["UploadedFile"];
                string version = Request.Form["Version"];

                try
                {
                    #region 前端無傳送檔案進來

                    if (FileModel == null)
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
                    else
                    {
                        string ExtensionName = Path.GetExtension(FileModel.FileName).ToLower();
                        if (ExtensionName != ".xls" && ExtensionName != ".xlsx")
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                            return Json(result);
                        }
                    }

                    #endregion 前端檔案大小不符或不為Excel檔案(驗證)

                    #region 上傳檔案

                    string pathType = Path.GetExtension(FileModel.FileName)
                                      .Substring(1); //上傳的檔案類型

                    var fileName = string.Format("{0}.{1}",
                                                 Excel_UploadName.C01.GetDescription(),
                                                 pathType); //固定轉成此名稱

                    Cache.Invalidate(CacheList.C01ExcelName); //清除 Cache
                    Cache.Set(CacheList.C01ExcelName, fileName); //把資料存到 Cache

                    #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                    string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                    string path = Path.Combine(projectFile, fileName);

                    FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                    //呼叫上傳檔案 function
                    result = FileRelated.FileUpLoadinPath(path, FileModel);
                    if (!result.RETURN_FLAG)
                    {
                        return Json(result);
                    }

                    #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                    #region 讀取Excel資料 使用ExcelDataReader 並且組成 json

                    List<C01ViewModel> dataModel = C0Repository.getC01Excel(pathType, path, version);
                    if (dataModel.Count > 0)
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.C01ExcelfileData); //清除 Cache
                        Cache.Set(CacheList.C01ExcelfileData, dataModel); //把資料存到 Cache
                        new BondsCheckRepository<C01ViewModel>(dataModel, Check_Table_Type.Bonds_C01_HK_VN_UpLoad);
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
                    result.DESCRIPTION = ex.Message;
                }
            }
            else
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
            }

            return Json(result);
        }
        #endregion

        #region TransferC01
        [BrowserEvent("C01 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferC01(string country, string version)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.C01ExcelName))
                {
                    fileName = (string)Cache.Get(CacheList.C01ExcelName);  //從Cache 抓資料
                }

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();

                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                string pathType = path.Split('.')[1]; //抓副檔名
                List<C01ViewModel> dataModel = C0Repository.getC01Excel(pathType, path, version);

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.C01TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogC01Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                {
                    txtpath = configTxtName; //有設定webConfig且不為空就取代
                }

                #endregion txtlog 檔案名稱

                #region save
                MSGReturnModel resultSave = C0Repository.saveC01(country,version,dataModel); //save to DB
                #endregion save

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success
                                     .GetDescription(Table_Type.C01.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                                         .GetDescription(Table_Type.C01.ToString(), resultSave.DESCRIPTION);
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

        #region GetC07Version
        [HttpPost]
        public JsonResult GetC07Version(string productCode, string reportDate)
        {
            List<string> versionData = C0Repository.getC07Version("", productCode, reportDate);
            return Json(string.Join(",", versionData));
        }
        #endregion

        #region Get C07
        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢C07系列資料")]
        [HttpPost]
        public JsonResult GetC07Data(string debtType, string groupProductCode, string productCode, string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

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
                var queryData = C0Repository.getC07(debtType, groupProductCode, productCode, reportDate, version);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.C07DbfileData); //清除
                Cache.Set(CacheList.C07DbfileData, queryData.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = "查無資料，請檢查：" + Environment.NewLine + "1. 基準日是否正確" + Environment.NewLine + "2. C07 的 專案名稱 和 流程名稱 是否存在 D01";
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

        #region GetCacheC07Data
        [HttpPost]
        public JsonResult GetCacheC07Data(jqGridParam jdata)
        {
            List<C07ViewModel> data = new List<C07ViewModel>();
            if (Cache.IsSet(CacheList.C07DbfileData))
            {
                data = (List<C07ViewModel>)Cache.Get(CacheList.C07DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr" }));
        }
        #endregion

        #region 下載 Excel

        /// <summary>
        /// 下載 Excel (C07Mortgage.C07Bond)
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載C07Excel檔案")]
        [HttpPost]
        public ActionResult GetC07Excel(string debtType, string groupProductCode, string productCode, string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;
            string type = "";

            try
            {
                if (GroupProductCode.M.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.M.GetDescription();
                    type = Excel_DownloadName.C07Mortgage.ToString();
                }
                else if (GroupProductCode.B.ToString().Equals(debtType))
                {
                    debtType = GroupProductCode.B.GetDescription();
                    type = Excel_DownloadName.C07Bond.ToString();
                }

                if (type != "")
                {
                    var C07Data = C0Repository.getC07(debtType, groupProductCode, productCode, reportDate, version);

                    path = type.GetExelName();
                    result = C0Repository.DownLoadExcel(type, ExcelLocation(path), C07Data.Item2);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(type, ex.Message);
            }

            return Json(result);
        }

        #endregion 下載 Excel

        #region GetC07AavancedData
        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("債券減損計算結果進階查詢")]
        [HttpPost]
        public JsonResult GetC07AdvancedData(string groupProductCode, string productCode, string reportDate, string version, string assessmentSubKind, string watchIND)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = C0Repository.getC07Advanced(groupProductCode, productCode, reportDate, version, assessmentSubKind, watchIND);
                result.RETURN_FLAG = (queryData.Item1 == "" ? true:false);
                Cache.Invalidate(CacheList.C07AdvancedDbfileData); //清除
                Cache.Set(CacheList.C07AdvancedDbfileData, queryData.Item2); //把資料存到 Cache

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = queryData.Item1;
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

        #region GetCacheC07AavancedData
        [HttpPost]
        public JsonResult GetCacheC07AavancedData(jqGridParam jdata)
        {
            List<C07AdvancedViewModel> data = new List<C07AdvancedViewModel>();
            if (Cache.IsSet(CacheList.C07AdvancedDbfileData))
            {
                data = (List<C07AdvancedViewModel>)Cache.Get(CacheList.C07AdvancedDbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr" }));
        }
        #endregion

        #region GetC07AdvancedExcel
        [BrowserEvent("下載C07AdvancedExcel檔案")]
        [HttpPost]
        public ActionResult GetC07AdvancedExcel(string groupProductCode, string productCode, string reportDate, string version, string assessmentSubKind, string watchIND)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;
            string type = "";

            try
            {
                type = Excel_DownloadName.C07Advanced.ToString();

                if (type != "")
                {
                    var C07AdvancedData = C0Repository.getC07Advanced(groupProductCode, productCode, reportDate, version, assessmentSubKind, watchIND);

                    path = type.GetExelName();
                    result = C0Repository.DownloadC07AdvancedExcel(type, ExcelLocation(path), C07AdvancedData.Item2);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(type, ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region GetC07AavancedSumData
        [BrowserEvent("債券減損計算結果進階查詢金額總計")]
        [HttpPost]
        public JsonResult GetC07AavancedSumData(string reportDate, string version, string groupProductCode, string groupProductName, string referenceNbr,string assessmentSubKind, string watchIND,string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = C0Repository.getC07AdvancedSum(reportDate, version, groupProductCode, groupProductName, referenceNbr, assessmentSubKind, watchIND, productCode);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.C07AdvancedSumDbfileData); //清除
                Cache.Set(CacheList.C07AdvancedSumDbfileData, queryData.Item2); //把資料存到 Cache

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

        #region GetCacheC07AavancedSumData
        [HttpPost]
        public JsonResult GetCacheC07AavancedSumData(jqGridParam jdata)
        {
            List<C07AdvancedSumViewModel> data = new List<C07AdvancedSumViewModel>();
            if (Cache.IsSet(CacheList.C07AdvancedSumDbfileData))
            {
                data = (List<C07AdvancedSumViewModel>)Cache.Get(CacheList.C07AdvancedSumDbfileData);
            }
            return Json(jdata.modelToJqgridResult(data, true));
        }
        #endregion

        #region GetC07AdvancedSumExcel
        [BrowserEvent("下載C07AdvancedSumExcel檔案")]
        [HttpPost]
        public ActionResult GetC07AdvancedSumExcel(string reportDate, string version, string groupProductCode, string groupProductName, string referenceNbr,string assessmentSubKind, string watchIND,string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;
            string type = "";

            try
            {
                type = Excel_DownloadName.C07AdvancedSum.ToString();

                if (type != "")
                {
                    var queryData = C0Repository.getC07AdvancedSum(reportDate, version, groupProductCode, groupProductName, referenceNbr, assessmentSubKind, watchIND, productCode);

                    path = type.GetExelName();
                    result = C0Repository.DownloadC07AdvancedSumExcel(type, ExcelLocation(path), queryData.Item2);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(type, ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region UploadC10
        [BrowserEvent("Upload C10")]
        [HttpPost]
        public JsonResult UploadC10()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 前端無傳送檔案進來

                if (!Request.Files.AllKeys.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                    return Json(result);
                }

                #endregion 前端無傳送檔案進來

                #region 前端檔案大小不符或不為Excel檔案(驗證)

                var FileModel = Request.Files["UploadedFile"];
                string reportDate = Request.Form["reportDate"];
                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(reportDate, out dt);

                string pathType = Path.GetExtension(FileModel.FileName)
                       .Substring(1); //上傳的檔案類型

                List<string> pathTypes = new List<string>()
                {
                    "xlsx","xls"
                };
                if (!pathTypes.Contains(pathType.ToLower()))
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不符或不為Excel檔案(驗證)

                #region 上傳檔案

                var fileName = string.Format("{0}.{1}",
                    Excel_UploadName.C10.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.C10ExcelName); //清除 Cache
                Cache.Set(CacheList.C10ExcelName, fileName); //把Excel_name存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);

                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel);
                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料 使用ExcelDataReader 並且組成 json

                var stream = FileModel.InputStream;
                var data = C0Repository.getExcel(pathType, stream, dt);
                if (!data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = data.Item1;
                    return Json(result);
                }
                List<C10ViewModel> dataModel = data.Item2;
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.C10ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.C10ExcelfileData, dataModel); //把資料存到 Cache
                    new BondsCheckRepository<C10ViewModel>(dataModel, Check_Table_Type.Bonds_C10_UpLoad_Check);
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
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }
        #endregion

        #region C10 jqGrid欄位大小配置
        private jqGridData<C10ViewModel> C10GridData()
        {
            return new C10ViewModel().TojqGridData(new int[] {
               120, 50, 165, 140, 160, 80, 170, 130, 125, 160, 140, 130, 120, 135, 140, 90, 80, 90, 80,
                135, 150, 150, 150, 150, 215, 150, 150, 150, 130, 180, 70, 125, 150, 110, 100, 90, 160
            });
        }
        #endregion

        #region C10Detail jqGrid欄位大小配置
        private jqGridData<C10DetailViewModel> C10DetailGridData()
        {
            return new C10DetailViewModel().TojqGridData(new int[] {
               120, 50, 165, 220, 60, 100, 150, 130, 125, 160, 140, 130, 120, 135, 140, 90, 80, 90, 120
            });
        }
        #endregion

        #region C04前瞻性資料
        [BrowserEvent("查詢C04檔案")]
        [HttpPost]
        public JsonResult SearchC04(
            string dateStart,
            string dateEnd,
            bool lastFlag
        )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var datas = C0Repository.GetC04(dateStart, dateEnd, lastFlag);
            Cache.Invalidate(CacheList.C04_1DbfileData); //清除
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Set(CacheList.C04_1DbfileData, datas); //把資料存到 Cache
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.C04.tableNameGetDescription());
            }
            return Json(result);
        }

        /// <summary>
        ///下載Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載Excel檔案")]
        [HttpPost]
        public JsonResult GetC0Excel(string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            switch (type)
            {
                case "C04":
                    if (Cache.IsSet(CacheList.C04_1DbfileData))
                    {
                        var C04_1 = Excel_DownloadName.C04_1.ToString();
                        var C04_1Data = (List<C04ViewModel>)Cache.Get(CacheList.C04_1DbfileData);  //從Cache 抓資料
                        result = C0Repository.DownLoadExcel(C04_1, ExcelLocation(C04_1.GetExelName()), C04_1Data);
                        return Json(result);
                    }
                    break;
                case "C04Trnasfer":
                    if (Cache.IsSet(CacheList.C04TransferData))
                    {
                        var C04_Transfer = Excel_DownloadName.C04_Transfer.ToString();
                        var C04_TransferData = (List<System.Dynamic.ExpandoObject>)Cache.Get(CacheList.C04TransferData);  //從Cache 抓資料
                        result = C0Repository.DownLoadExcel(C04_Transfer, ExcelLocation(C04_Transfer.GetExelName()), C04_TransferData);
                        return Json(result);
                    }
                    break;
            }
            result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            return Json(result);
        }

        [HttpPost]
        public JsonResult SearchC04Pros()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = true;
            result.Datas = Json(C0Repository.GetC04Pro());
            return Json(result);
        }

        [BrowserEvent("轉檔落後期數")]
        [HttpPost]
        public JsonResult TrnasferC04(string from, string to, List<C04ProViewModel> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var _data = C0Repository.C04Transfer(from, to, data);
            Cache.Invalidate(CacheList.C04TransferData); //清除
            if (_data.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Set(CacheList.C04TransferData, _data); //把資料存到 Cache
                var obj = _data.First();
                var jqgridInfo = obj.TojqGridData();
                result.Datas = Json(jqgridInfo);
            }
            else
            {
                result.DESCRIPTION = "";
            }
            return Json(result);
        }
        

        [HttpPost]
        public ActionResult GetCacheData(jqGridParam jdata, string type)
        {
            switch (type)
            {
                case "C04_Transfer":
                    if (Cache.IsSet(CacheList.C04TransferData))
                        return Content((jdata.dynToJqgridResult(
                         (List<System.Dynamic.ExpandoObject>)Cache.Get(CacheList.C04TransferData))));
                    break;
            }
            return null;
        }
        #endregion

        #region GetCacheC10Data
        [HttpPost]
        public ActionResult GetCacheC10Data(jqGridParam jdata, string type)
        {
            switch (type)
            {
                case "C10_Excel":
                    if (Cache.IsSet(CacheList.C10ExcelfileData))
                        //從Cache 抓資料
                        return Json(jdata.modelToJqgridResult((List<C10ViewModel>)Cache.Get(CacheList.C10ExcelfileData), true));
                    break;

                case "C10_Db":
                    if (Cache.IsSet(CacheList.C10DbfileData))
                        return Json(jdata.modelToJqgridResult((List<C10DetailViewModel>)Cache.Get(CacheList.C10DbfileData), true));
                    break;
            }
            return null;
        }
        #endregion

        #region 查詢C10資料
        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="ReportDate">報導日</param>
        /// <param name="Version">版本</param>
        /// <param name="BondNumber">債券編號</param>
        /// <returns></returns>
        [BrowserEvent("查詢C10資料")]
        [HttpPost]
        public JsonResult GetData(string ReportDate, string Version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A41.tableNameGetDescription());

            DateTime _ReportDate = new DateTime();
            int _Version = 0;
            DateTime _OriginationDate = DateTime.MinValue;
            if (!DateTime.TryParse(ReportDate, out _ReportDate) ||
                !Int32.TryParse(Version, out _Version))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            try
            {
                var resultData = C0Repository.GetC10(_ReportDate, _Version);
                

                if (resultData.Any())
                {
                    Cache.Invalidate(CacheList.C10DbfileData); //清除
                    Cache.Set(CacheList.C10DbfileData, resultData); //把資料存到 Cache
                    result.RETURN_FLAG = true;
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.C10.tableNameGetDescription());
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }
        #endregion

        #region 下載C10 Excel
        /// <summary>
        /// 下載C10 Excel
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        [BrowserEvent("下載C10Excel檔案")]
        [HttpPost]
        public JsonResult GetC10Excel(string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.C10DbfileData))
            {
                var C10 = Excel_DownloadName.C10.ToString();
                var C10Data = (List<C10DetailViewModel>)Cache.Get(CacheList.C10DbfileData);  //從Cache 抓資料
                result = C0Repository.DownLoadExcelC10(type, ExcelLocation(type.GetExelName()), C10Data);


            }
            return Json(result);
        }
        #endregion


        #region 轉檔前檢核
        [BrowserEvent("回傳資料表C10有無資料")]
        [HttpPost]
        public JsonResult TransferConfirmC10(string reportDate) {

            #region 回傳C10有無相同reportdate的資料
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            int version = 0;
            DateTime dt = TypeTransfer.stringToDateTime(reportDate);
            string filename = Table_Type.C10.ToString();

            
            var queryData = C0Repository.GetC10(dt, version);
            

            if (common.checkTransferCheck(filename,"C10", dt, version) && queryData != null && queryData.Count> 0)
            {
                result.RETURN_FLAG = true;
                result.DESCRIPTION = Message_Type.Uplaod_data_Overwrite_File.GetDescription();
            }
            else
            {
                result.RETURN_FLAG = false;
            }
            return Json(result);
            #endregion 回傳C10有無相同reportdate的資料
        }
        #endregion
    }
}