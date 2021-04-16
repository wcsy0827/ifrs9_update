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
    public class A4Controller : CommonController
    {
        private IA4Repository A4Repository;
        private string[] selects = { "All", "B01", "C01" };
        private string[] selectsMortgage = { "All", "B01", "C01", "C02" };
        private List<SelectOption> actions = null;
        DateTime startTime = DateTime.MinValue;
        protected Common common
        {
            get;
            private set;
        }

        public A4Controller()
        {
            this.A4Repository = new A4Repository();
            this.common = new Common();
            startTime = DateTime.Now;
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }};
        }


        /// <summary>
        /// A4(上傳檔案)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult Index()
        {
            var jqgridInfo = A41GridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// A41(債券明細檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A41Detail()
        {
            var jqgridInfo = A41GridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// A42(國庫券月結資料檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A42()
        {     
            actions =  new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            return View();
        }

        /// <summary>
        /// A45(產業別資訊檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A45()
        {
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");
            int[] widths = new int[] { 100, 100, 160, 160, 160, 160, 160, 280, 120, 120};
            string[] aligns = new string[] { "left", "left", "left", "left", "left", "left", "left", "left", "left", "center" };
            var jqgridInfo = new A45ViewModel().TojqGridData(widths, aligns);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;

            return View();
        }

        /// <summary>
        /// 執行減損計算 (債券)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A98Detail()
        {
            ViewBag.selectOption = new SelectList(
                selects.Select(x => new { Text = x, Value = x }), "Value", "Text");
            return View();
        }

        /// <summary>
        /// 執行減損計算 (房貸)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A99Detail()
        {
            ViewBag.selectOption = new SelectList(
                selectsMortgage.Select(x => new { Text = x, Value = x }), "Value", "Text");
            return View();
        }

        /// <summary>
        /// A46(固收提供的CEIC資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A46Detail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A46).TojqGridData(new int[2]{200,250});
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// A47(固收提供的Moody資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A47Detail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A47).TojqGridData(new int[3] { 150,250,250});
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// A48(固收提供的阿布達比資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A48Detail()
        {
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.A48).TojqGridData(new int[] {55,110,85,125,180,210,200});
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// A49(財報揭露會計帳值資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A49()
        {
            actions = new List<SelectOption>() {
                      new SelectOption() {Text="查詢",Value="search" },
                      new SelectOption() {Text="上傳&存檔",Value="upload" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            return View();
        }

        /// <summary>
        /// 查詢&補登A95(債券種類與減損階段評估)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A95Detail()
        {
            var jqgridInfo = new A95ViewModel().TojqGridData();
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢&下載",Value="downLoad" },
                new SelectOption() {Text="上傳&存檔",Value="upLoad" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var _All = new SelectOption() { Text = "All", Value = "All" };     
            var _ASK = EnumUtil.GetValues<AssessmentSubKind_Type>().Select(x => new SelectOption() {
                Text = x.GetDescription(),
                Value = x.GetDescription()
            }).ToList();
            _ASK.Insert(0, _All);
            var _bondType = new List<SelectOption>();
            _bondType.Add(_All);
            _bondType.AddRange(new List<SelectOption>() {
                new SelectOption() { Text = "主權及國營事業債",Value = "主權及國營事業債"},
                new SelectOption() { Text = "其他債券",Value = "其他債券"}
            });
            ViewBag.bondType = new SelectList(_bondType, "Value", "Text");
            ViewBag.ASK = new SelectList(_ASK, "Value", "Text");
            return View();
        }

        /// <summary>
        /// A44(換券資訊檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A44()
        {
            return View();
        }

        #region 190628 John.投會換券應收未收金額修正
        /// <summary>
        /// A44_2Upload(換券應收未收修正檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A44_2()
        {
            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;//20200926 alibaba 系統使用者 202008210166-00
            return View();
        }

        /// <summary>
        /// A44_2Detail(換券應收未收修正檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A44_2Detail()
        {
            return View();
        }

        #region UploadA44_2
        [BrowserEvent("上傳A44_2換券應收未收金額修正檔")]
        [HttpPost]
        public JsonResult UploadA44_2()
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
                    Excel_UploadName.A44_2Upload.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A44_2ExcelName); //清除 Cache
                Cache.Set(CacheList.A44_2ExcelName, fileName); //把Excel_name存到 Cache

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

                var data = A4Repository.getA44_2Excel(pathType, stream, dt);
                if (!data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = data.Item1;
                    return Json(result);
                }
                List<A44_2ViewModel> dataModel = data.Item2;
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A44_2ExcelfileDate); //清除 Cache
                    Cache.Set(CacheList.A44_2ExcelfileDate, dataModel); //把資料存到 Cache
                    //20200930 alibaba 1筆舊券換多筆新券提示訊息 202008210166-00
                    if (dataModel.Exists(y => y.Multi_NewBonds == "Y"))
                    { result.DESCRIPTION = "請注意! 本次上傳應收息修正檔有一筆舊券換多筆新券情形。"; }
                    //end 20200930 alibaba
                    //new BondsCheckRepository<A44_2ViewModel>(dataModel, Check_Table_Type.Bonds_C10_UpLoad_Check);  //0615 檢核部分先不寫，之後再看要做哪些檢核
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

        #region Excel檔存進DB
        [BrowserEvent("A44_2換券應收未收金額修正檔存入DB")]
        [HttpPost]
        public JsonResult TransferA44_2(string reportDate, List<A44_2ViewModel> dataModel)//20200924 alibaba 新增傳入jqgrid的obj 202008210166-00
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A44_2ExcelName))
                    fileName = (string)Cache.Get(CacheList.A44_2ExcelName);  //從Cache 抓資料Name

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);               

                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(reportDate, out dt);

                string errorMessage = string.Empty;

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱


                string txtpath = SetFile.A44_2TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA44_2Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代


                #endregion txtlog 檔案名稱
                #region save A44_2
                MSGReturnModel resultA44_2 = A4Repository.SaveA44_2(dataModel, reportDate); //save to DB

                int ver = 0; ///上傳檔案，版本皆為0

                bool A44_2Log = CommonFunction.saveLog(Table_Type.A44_2,
                    fileName, SetFile.ProgramName, resultA44_2.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, ver, dt); //寫sql Log
                TxtLog.txtLog(Table_Type.A44_2, resultA44_2.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save A44_2

                result.RETURN_FLAG = resultA44_2.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A44_2.ToString());

                if (!result.RETURN_FLAG)
                {

                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A44_2.ToString(), resultA44_2.DESCRIPTION);

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

        #region CheckA44_2MaxVersion
        [BrowserEvent("檢查目前報導日A44_2最大版本")]
        [HttpPost]
        public JsonResult CheckA44_2MaxVersion(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = TypeTransfer.stringToDateTime(reportDate);
            try
            {
                var CheckFlag = A4Repository.CheckMaxVerion(dt);
                if (CheckFlag)
                {
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "目前最大版本A41已經執行過B01、C01，請投資風控重新上傳A41!";
                }
                return Json(result);
            }
            catch(Exception ex)
            {

                result.RETURN_FLAG = true;
                result.DESCRIPTION = $"取得最大版本錯誤，錯誤訊息:{ex}";
                return Json(result);
            }

        }
        #endregion

        #region CheckA44_2Data
        [BrowserEvent("檢查A44_2上傳內容")]
        [HttpPost]
        public JsonResult CheckA44_2Data(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = TypeTransfer.stringToDateTime(reportDate);
            string filename = Table_Type.A44_2.ToString();
            int version = 0;

            var queryData = A4Repository.GetA44_2(dt, version);
            if (common.checkTransferCheck(filename, "A44_2", dt, version) && queryData != null && queryData.Count > 0)
            {
                result.RETURN_FLAG = true;

                result.DESCRIPTION = Message_Type.Uplaod_data_Overwrite_File.GetDescription();
            }
            else
            {
                result.RETURN_FLAG = false;
            }

            return Json(result);
        }
        #endregion

        #region Get A44_2Data
        [BrowserEvent("查詢A44_2資料")]
        [HttpPost]
        public JsonResult GetA44_2Data(string ReportDate, string Version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            //result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A44_2.tableNameGetDescription());

            DateTime _ReportDate = new DateTime();
            int _Version = 0;
            if (!DateTime.TryParse(ReportDate, out _ReportDate) ||
                !int.TryParse(Version, out _Version))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            try
            {
                var resultData = A4Repository.GetA44_2Detail(_ReportDate, _Version);
                if (resultData.Any())
                {
                    Cache.Invalidate(CacheList.A44_2DbfileData); //清除
                    Cache.Set(CacheList.A44_2DbfileData, resultData); //把資料存到 Cache
                    result.RETURN_FLAG = true;
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A44_2.tableNameGetDescription());
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
        #region GetCacheA44_2Data
        [HttpPost]
        public ActionResult GetCacheA44_2Data(jqGridParam jdata, string type)
        {
            switch (type)
            {
                case "A44_2Upload":
                    if (Cache.IsSet(CacheList.A44_2ExcelfileDate))
                        //從Cache 抓資料
                        return Json(jdata.modelToJqgridResult((List<A44_2ViewModel>)Cache.Get(CacheList.A44_2ExcelfileDate), true));
                    break;

                case "A44_2_Db":
                    if (Cache.IsSet(CacheList.A44_2DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A44_2DetailViewModel>)Cache.Get(CacheList.A44_2DbfileData), true));
                    break;
            }
            return null;
        }
        #endregion




        #endregion

        #region GetCacheData
        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata, string type)
        {
            List<A41ViewModel> data = new List<A41ViewModel>();
            switch (type)
            {
                case "Excel":
                    if (Cache.IsSet(CacheList.A41ExcelfileData))
                        data = (List<A41ViewModel>)Cache.Get(CacheList.A41ExcelfileData);  //從Cache 抓資料
                    break;

                case "Db":
                    if (Cache.IsSet(CacheList.A41DbfileData))
                        data = (List<A41ViewModel>)Cache.Get(CacheList.A41DbfileData);
                    break;
            }
            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr", "Lots", "Version" }));
        }
        #endregion

        [HttpPost]
        public JsonResult GetCacheA95Data(jqGridParam jdata, string type)
        {
            List<A95ViewModel> data = new List<A95ViewModel>();
            switch (type)
            {
                case "Excel":
                    if (Cache.IsSet(CacheList.A95ExcelfileData))
                        data = (List<A95ViewModel>)Cache.Get(CacheList.A95ExcelfileData);  //從Cache 抓資料
                    break;

                case "Db":
                    if (Cache.IsSet(CacheList.A95DbfileData))
                        data = (List<A95ViewModel>)Cache.Get(CacheList.A95DbfileData);
                    break;
            }
            return Json(jdata.modelToJqgridResult(data));
        }

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="ReportDate">報導日</param>
        /// <param name="OriginationDate">購入日</param>
        /// <param name="Version">版本</param>
        /// <param name="BondNumber">債券編號</param>
        /// <returns></returns>
        [BrowserEvent("查詢A41資料")]
        [HttpPost]
        public JsonResult GetData(string ReportDate, string OriginationDate, string Version, string BondNumber)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A41.tableNameGetDescription());

            DateTime _ReportDate = new DateTime();
            int _Version = 0;
            DateTime _OriginationDate = DateTime.MinValue;
            if (!DateTime.TryParse(ReportDate, out _ReportDate) ||
                !Int32.TryParse(Version, out _Version) ||
                (!OriginationDate.IsNullOrWhiteSpace() &&
                !DateTime.TryParse(OriginationDate, out _OriginationDate)))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            try
            {
                var resultData = A4Repository.GetA41(_ReportDate, _OriginationDate, _Version, BondNumber);
                if (resultData.Any())
                {
                    Cache.Invalidate(CacheList.A41DbfileData); //清除
                    Cache.Set(CacheList.A41DbfileData, resultData); //把資料存到 Cache
                    result.RETURN_FLAG = true;
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A41.tableNameGetDescription());
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
        /// /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="reportDate">評估基準日/報導日</param>
        /// <returns></returns>
        [BrowserEvent("查詢A42資料")]
        [HttpPost]
        public JsonResult GetA42Data(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("A42");

            try
            {
                var queryData = A4Repository.getA42(reportDate);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.A42DbfileData); //清除
                Cache.Set(CacheList.A42DbfileData, queryData.Item2); //把資料存到 Cache

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

        #region GetCacheA42Data
        [HttpPost]
        public JsonResult GetCacheA42Data(jqGridParam jdata, string type)
        {
            List<A42ViewModel> data = new List<A42ViewModel>();

            switch (type)
            {
                case "Excel":
                    if (Cache.IsSet(CacheList.A42ExcelfileData))
                    {
                        data = (List<A42ViewModel>)Cache.Get(CacheList.A42ExcelfileData); //從Cache 抓資料
                    }  
                    break;

                case "Db":
                    if (Cache.IsSet(CacheList.A42DbfileData))
                    { data = (List<A42ViewModel>)Cache.Get(CacheList.A42DbfileData); }
                    break;
            }

            return Json(jdata.modelToJqgridResult(data, true));
        }
        #endregion

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("A42上傳檔案")]
        [HttpPost]
        public JsonResult UploadA42()
        {
            MSGReturnModel result = new MSGReturnModel();

            //## 如果有任何檔案類型才做
            if (Request.Files.AllKeys.Any())
            {
                var FileModel = Request.Files["UploadedFile"];
                string processingDate = Request.Form["processingDate"];
                string reportDate = Request.Form["reportDate"];

                //result = A4Repository.checkA42Duplicate(reportDate);
                //if (result.RETURN_FLAG == false)
                //{
                //    return Json(result);
                //}

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
                                                 Excel_UploadName.A42.GetDescription(),
                                                 pathType); //固定轉成此名稱

                    Cache.Invalidate(CacheList.A42ExcelName); //清除 Cache
                    Cache.Set(CacheList.A42ExcelName, fileName); //把資料存到 Cache

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

                    var stream = FileModel.InputStream;
                    List<A42ViewModel> dataModel = A4Repository.getA42Excel(pathType, path, processingDate, reportDate);
                    if (dataModel.Count > 0)
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.A42ExcelfileData); //清除 Cache
                        Cache.Set(CacheList.A42ExcelfileData, dataModel); //把資料存到 Cache
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
                return Json(result);
            }

            return Json(result);
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A42 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA42(string processingDate, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime dt = DateTime.MinValue;
            DateTime.TryParse(reportDate, out dt);
            var data = CommonFunction.getVersion(dt, "A41");//抓A41報導日最大版本
            var version = data.Last();
            try
            {
                //result = A4Repository.checkA42Duplicate(reportDate);
                //if (result.RETURN_FLAG == false)
                //{
                //    return Json(result);
                //}

                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A42ExcelName))
                {
                    fileName = (string)Cache.Get(CacheList.A42ExcelName);  //從Cache 抓資料
                }

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();

                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                string pathType = path.Split('.')[1]; //抓副檔名
                List<A42ViewModel> dataModel = A4Repository.getA42Excel(pathType, path, processingDate, reportDate);

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A42TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA42Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                {
                    txtpath = configTxtName; //有設定webConfig且不為空就取代
                }

                #endregion txtlog 檔案名稱

                #region save
                MSGReturnModel resultSave = A4Repository.saveA42(dataModel, version); //save to DB
                #endregion save
                #region 20200514 John.A42回寫A41邏輯調整.增加手動上傳A42時的Log
                bool A42Log = CommonFunction.saveLog(Table_Type.A42, fileName, SetFile.ProgramName, resultSave.RETURN_FLAG, Debt_Type.B.ToString(),
                    startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,int.Parse(version), dt); //寫sql Log
                #endregion
                               
                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A42.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(Table_Type.A42.ToString(), resultSave.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #region GetA42Excel
        [BrowserEvent("A42匯出Excel")]
        [HttpPost]
        public ActionResult GetA42Excel(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                var data = A4Repository.getA42(reportDate);

                path = "A42".GetExelName();
                result = A4Repository.DownLoadA42Excel(ExcelLocation(path), data.Item2);
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
        /// /// 前端抓資料時呼叫
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("查詢A45資料")]
        [HttpPost]
        public JsonResult GetA45Data(string bloombergTicker, string processingDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("A45");

            try
            {
                var queryData = A4Repository.getA45(bloombergTicker,processingDate);
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

        /// <summary>
        /// 查詢A46資料
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A46資料")]
        [HttpPost]
        public JsonResult GetA46Data(bool searchAll)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            try
            {
                var Datas = A4Repository.GetA46(searchAll);
                if (Datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A46DbfileData); //清除
                    Cache.Set(CacheList.A46DbfileData, Datas); //把資料存到 Cache
                }
                else
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A46.tableNameGetDescription());
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢A47資料
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A47資料")]
        [HttpPost]
        public JsonResult GetA47Data(bool searchAll)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            try
            {
                var Datas = A4Repository.GetA47(searchAll);
                if (Datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A47DbfileData); //清除
                    Cache.Set(CacheList.A47DbfileData, Datas); //把資料存到 Cache
                }
                else
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A47.tableNameGetDescription());
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢A48資料
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A48資料")]
        [HttpPost]
        public JsonResult GetA48Data(bool searchAll)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            try
            {
                var Datas = A4Repository.GetA48(searchAll);
                if (Datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A48DbfileData); //清除
                    Cache.Set(CacheList.A48DbfileData, Datas); //把資料存到 Cache
                }
                else
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A48.tableNameGetDescription());
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢A95資料
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A95資料")]
        [HttpPost]
        public JsonResult GetA95Data(string reportDate,string version,string bondType,string ASK)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime d = new DateTime();
            int v = 0;
            if (!DateTime.TryParse(reportDate, out d) || !Int32.TryParse(version, out v))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            try
            {
                var data = A4Repository.GetA95(d, v, bondType, ASK);
                if (data.Item1)
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A95DbfileData); //清除
                    Cache.Set(CacheList.A95DbfileData, data.Item2); //把資料存到 Cache
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 下載A95Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A95Excel檔案")]
        [HttpPost]
        public JsonResult GetA95Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.A95DbfileData))
            {
                var A95 = Excel_DownloadName.A95.ToString();
                var A95Data = (List<A95ViewModel>)Cache.Get(CacheList.A95DbfileData);  //從Cache 抓資料
                result = A4Repository.DownLoadExcel(A95, ExcelLocation(A95.GetExelName()), A95Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            return Json(result);
        }

        /// <summary>
        /// 抓取資料庫最後一天日期log Data
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetLogData(string debt)
        {
            if (Debt_Type.M.ToString().Equals(debt))
            {
                selects = selectsMortgage;
            }
            List<string> logDatas = A4Repository.GetLogData(selects.ToList(), debt);
            return Json(string.Join(",", logDatas));
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("Data Requirements Excel檔存到DB")]
        [HttpPost]
        public JsonResult Transfer(string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
               
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A41ExcelName))
                    fileName = (string)Cache.Get(CacheList.A41ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A41ViewModel> dataModel = new List<A41ViewModel>();

                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(reportDate , out dt);

                int v = 0;
                Int32.TryParse(version, out v);

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A4Repository.getExcel(pathType, stream, dt, v);
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

                string txtpath = SetFile.A41TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA4Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save Bond_Account_Info(A41)



                MSGReturnModel resultA41 = A4Repository.saveA41(dataModel, reportDate, version); //save to DB
                bool A41Log = CommonFunction.saveLog(Table_Type.A41,
                    fileName, SetFile.ProgramName, resultA41.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, v, dt); //寫sql Log
                TxtLog.txtLog(Table_Type.A41, resultA41.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Bond_Account_Info(A41)

                result.RETURN_FLAG = resultA41.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A41.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A41.ToString(), resultA41.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }
            return Json(result);
        }

        #region TransferA45
        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A45 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA45()
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A45ExcelName))
                {
                    fileName = (string)Cache.Get(CacheList.A45ExcelName);  //從Cache 抓資料
                }

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();

                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                List<A45ViewModel> dataModel = new List<A45ViewModel>();
                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    string pathType = path.Split('.')[1]; //抓副檔名
                    dataModel = A4Repository.getA45Excel(pathType, stream);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A45TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA45Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                {
                    txtpath = configTxtName; //有設定webConfig且不為空就取代
                }

                #endregion txtlog 檔案名稱

                #region save

                MSGReturnModel resultA45 = A4Repository.saveA45(dataModel); //save to DB

                bool A45Log = CommonFunction.saveLog(Table_Type.A45,
                                                     fileName, SetFile.ProgramName, resultA45.RETURN_FLAG,
                                                     Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A45, resultA45.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save

                result.RETURN_FLAG = resultA45.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription() + " A95：Bond_Ticker_Info、A41：Bond_Account_Info 也一併更新了";

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                                         .GetDescription(Table_Type.A45.ToString(), resultA45.DESCRIPTION);
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

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A46 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA46()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A46ExcelName))
                    fileName = (string)Cache.Get(CacheList.A46ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A46ViewModel> dataModel = new List<A46ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A4Repository.getA46Excel(pathType, stream);
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

                string txtpath = SetFile.A46TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA4Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save Fixed_Income_CEIC_Info(A46)

                MSGReturnModel resultA46 = A4Repository.saveA46(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A46,
                    fileName, SetFile.ProgramName, resultA46.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A46, resultA46.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Fixed_Income_CEIC_Info(A46)

                result.RETURN_FLAG = resultA46.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A46.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A46.ToString(), resultA46.DESCRIPTION);
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
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A47 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA47()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A47ExcelName))
                    fileName = (string)Cache.Get(CacheList.A47ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A47ViewModel> dataModel = new List<A47ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A4Repository.getA47Excel(pathType, stream);
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

                string txtpath = SetFile.A47TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA4Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save Fixed_Income_Moody_Info(A47)

                MSGReturnModel resultA47 = A4Repository.saveA47(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A47,
                    fileName, SetFile.ProgramName, resultA47.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A47, resultA47.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Fixed_Income_Moody_Info(A47)

                result.RETURN_FLAG = resultA47.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A47.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A47.ToString(), resultA47.DESCRIPTION);
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
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A48 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA48()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A48ExcelName))
                    fileName = (string)Cache.Get(CacheList.A48ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<A48ViewModel> dataModel = new List<A48ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = A4Repository.getA48Excel(pathType, stream);
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

                string txtpath = SetFile.A48TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA4Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                    txtpath = configTxtName; //有設定webConfig且不為空就取代

                #endregion txtlog 檔案名稱

                #region save Fixed_Income_Moody_Info(A48)

                MSGReturnModel resultA48 = A4Repository.saveA48(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A48,
                    fileName, SetFile.ProgramName, resultA48.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A48, resultA48.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Fixed_Income_Moody_Info(A48)

                result.RETURN_FLAG = resultA48.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A48.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A48.ToString(), resultA48.DESCRIPTION);
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
        /// 轉檔把 A95Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A95 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA95()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A95ExcelName))
                    fileName = (string)Cache.Get(CacheList.A95ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                string pathType = path.Split('.')[1]; //抓副檔名
                List<A95ViewModel> dataModel = A4Repository.getA95Excel(pathType, path); //Excel轉成 Exhibit10Model

                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A95TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save Bond_Ticker_Info(A95)

                int v = 0;
                if (dataModel.Any())
                   Int32.TryParse(dataModel.First().Version, out v);

                MSGReturnModel resultA95 = A4Repository.saveA95(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.A95,
                    fileName, SetFile.ProgramName, resultA95.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,v); //寫sql Log
                TxtLog.txtLog(Table_Type.A95, resultA95.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save Bond_Ticker_Info(A95)

                result.RETURN_FLAG = resultA95.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.A95.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.A95.ToString(), resultA95.DESCRIPTION);
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
        /// 前端轉檔一系列動作
        /// </summary>
        /// <param name="type">目前要轉檔的表名</param>
        /// <param name="date">日期</param>
        /// <param name="version">版本(房貸沒有)</param>
        /// <param name="next">是否要執行下一項</param>
        /// <param name="debt">M:房貸 B:債券</param>
        /// <returns></returns>
        [BrowserEvent("轉檔")]
        [HttpPost]
        public JsonResult TransferToOther(string type, string date, string version, bool next, string debt)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            DateTime dat = DateTime.MinValue; //基準日
            int ver = 0; //版本

            if ((Debt_Type.M.ToString().Equals(debt) ? false : version.IsNullOrWhiteSpace() ||
                !Int32.TryParse(version, out ver) || ver == 0) //房貸沒有版本 , 債券version不等於0
                || !DateTime.TryParse(date, out dat))
                return Json(result);

            string tableName = string.Empty;
            string fileName = Excel_UploadName.A41.GetDescription().GetExelName(); //預設
            if (Debt_Type.B.ToString().Equals(debt))
            {
                string configFileName = ConfigurationManager.AppSettings["fileA4Name"];
                if (!string.IsNullOrWhiteSpace(configFileName))
                    fileName = configFileName; //config 設定就取代
            }
            if (Debt_Type.M.ToString().Equals(debt))
                fileName = "A01-IAS39,A02";

           
            switch (type)
            {
                case "All": //All 也是重B01開始 B01 => C01
                case "B01":
                    result = A4Repository.saveB01(ver, dat, debt);
                    bool B01Log = CommonFunction.saveLog(Table_Type.B01, fileName, SetFile.ProgramName,
                        result.RETURN_FLAG, debt, startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, ver == 0 ? 1 : ver,dat); //寫sql Log
                    result.Datas = Json(transferMessage(next, Transfer_Table_Type.C01.ToString())); //回傳要不要做下一個transfer
                    break;

                case "C01":
                    result = A4Repository.saveC01(ver, dat, debt);
                    bool C01Log = CommonFunction.saveLog(Table_Type.C01, fileName, SetFile.ProgramName,
                        result.RETURN_FLAG, debt, startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, ver == 0 ? 1 : ver,dat); //寫sql Log
                    //債券最多到C01
                    result.Datas = Json(transferMessage((debt.Equals("B") ? false : next), Transfer_Table_Type.C02.ToString()));
                    break;

                case "C02":
                    result = A4Repository.saveC02(ver, dat, debt);
                    bool C02Log = CommonFunction.saveLog(Table_Type.C02, fileName, SetFile.ProgramName,
                        result.RETURN_FLAG, debt, startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,1,dat); //寫sql Log
                    result.Datas = Json(transferMessage(false, string.Empty)); //目前到C02 而已
                    break;
            }

            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("上傳Data Requirements Excel檔案")]
        [HttpPost]
        public JsonResult Upload()
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
                string version = Request.Form["version"];
                DateTime dt = DateTime.MinValue;
                int ver = 0;
                DateTime.TryParse(reportDate, out dt);
                Int32.TryParse(version, out ver);
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
                    Excel_UploadName.A41.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A41ExcelName); //清除 Cache
                Cache.Set(CacheList.A41ExcelName, fileName); //把資料存到 Cache

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
                var data = A4Repository.getExcel(pathType, stream, dt, ver);
                if (!data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = data.Item1;
                    return Json(result);
                }
                List<A41ViewModel> dataModel = data.Item2;
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A41ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A41ExcelfileData, dataModel); //把資料存到 Cache
                    new BondsCheckRepository<A41ViewModel>(dataModel, Check_Table_Type.Bonds_A41_UpLoad);
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

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("A45上傳檔案")]
        [HttpPost]
        public JsonResult UploadA45()
        {
            MSGReturnModel result = new MSGReturnModel();

            //## 如果有任何檔案類型才做
            if (Request.Files.AllKeys.Any())
            {
                var FileModel = Request.Files["UploadedFile"];

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
                                                 Excel_UploadName.A45.GetDescription(),
                                                 pathType); //固定轉成此名稱

                    Cache.Invalidate(CacheList.A45ExcelName); //清除 Cache
                    Cache.Set(CacheList.A45ExcelName, fileName); //把資料存到 Cache

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

                    var stream = FileModel.InputStream;
                    List<A45ViewModel> dataModel = A4Repository.getA45Excel(pathType, stream);
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
                    result.DESCRIPTION = ex.Message;
                }
            }
            else
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                return Json(result);
            }

            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發(A46)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("A46上傳檔案")]
        [HttpPost]
        public JsonResult UploadA46(ValidateFiles FileModel)
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
                    Excel_UploadName.A46.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A46ExcelName); //清除 Cache
                Cache.Set(CacheList.A46ExcelName, fileName); //把資料存到 Cache

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
                var data = A4Repository.getA46Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A46ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A46ExcelfileData, data.Item2); //把資料存到 Cache
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
        /// 選擇檔案後點選資料上傳觸發(A47)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("A47上傳檔案")]
        [HttpPost]
        public JsonResult UploadA47(ValidateFiles FileModel)
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
                    Excel_UploadName.A47.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A47ExcelName); //清除 Cache
                Cache.Set(CacheList.A47ExcelName, fileName); //把資料存到 Cache

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
                var data = A4Repository.getA47Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A47ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A47ExcelfileData, data.Item2); //把資料存到 Cache
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
        /// 選擇檔案後點選資料上傳觸發(A48)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("A48上傳檔案")]
        [HttpPost]
        public JsonResult UploadA48(ValidateFiles FileModel)
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
                    Excel_UploadName.A48.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A48ExcelName); //清除 Cache
                Cache.Set(CacheList.A48ExcelName, fileName); //把資料存到 Cache

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
                var data = A4Repository.getA48Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A48ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A48ExcelfileData, data.Item2); //把資料存到 Cache
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
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("A95上傳檔案")]
        [HttpPost]
        public JsonResult UploadA95(ValidateFiles FileModel)
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
                    Excel_UploadName.A95.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A95ExcelName); //清除 Cache
                Cache.Set(CacheList.A95ExcelName, fileName); //把資料存到 Cache

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

                List<A95ViewModel> dataModel = A4Repository.getA95Excel(pathType, path);
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.A95ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A95ExcelfileData, dataModel); //把資料存到 Cache
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
            return Json(result);
        }

        #region private Function
        /// <summary>
        /// 判斷轉檔有沒有後續
        /// </summary>
        /// <param name="next"></param>
        /// <param name="nextType"></param>
        /// <returns></returns>
        private string transferMessage(bool next, string nextType)
        {
            return next ? "true," + nextType : "false";
        }

        private jqGridData<A41ViewModel> A41GridData()
        {
            return  new A41ViewModel().TojqGridData(new int[] {
                100, 130, 50, 165, 130, 170, 150, 130, 125, 160, 140, 130, 120, 135, 130, 90, 80, 90, 80,
                135, 150, 150, 150, 150, 215, 150, 150, 150, 130, 170, 70, 125, 150, 110, 65, 90, 215, 85,
                70, 115, 120, 95, 70, 75, 150, 80, 160, 170, 150, 150, 90, 150, 150, 80, 150, 75, 150, 100
            });
        }

        #endregion

        #region GetA44All
        [HttpPost]
        public JsonResult GetA44AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A4Repository.getA44All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A44DbfileData); //清除
                Cache.Set(CacheList.A44DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheA44_2Data (190628 John.投會換券應收未收金額修正)
        public JsonResult GetCacheA44_2Data(jqGridParam jdata)
        {
            List<A44_2ViewModel> data = new List<A44_2ViewModel>();
            if (Cache.IsSet(CacheList.A44_2ExcelfileDate))
            {
                data = (List<A44_2ViewModel>)Cache.Get(CacheList.A44_2ExcelfileDate);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetCacheA44Data
        [HttpPost]
        public JsonResult GetCacheA44Data(jqGridParam jdata)
        {
            List<A44ViewModel> data = new List<A44ViewModel>();
            if (Cache.IsSet(CacheList.A44DbfileData))
            {
                data = (List<A44ViewModel>)Cache.Get(CacheList.A44DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetA44Data
        [BrowserEvent("查詢A44資料")]
        [HttpPost]
        public JsonResult GetA44Data(A44ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A4Repository.getA44(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A44DbfileData); //清除
                Cache.Set(CacheList.A44DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveA44
        [BrowserEvent("儲存A44資料")]
        [HttpPost]
        public JsonResult SaveA44(string actionType, A44ViewModel dataModel, bool isoldbondsave = false)//20200925 alibaba isoldbondsave：舊券重複確認是否存檔 202008210166-00
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = A4Repository.saveA44(actionType, dataModel, isoldbondsave);//20200925 alibaba isoldbondsave：舊券重複確認是否存檔 202008210166-00

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;

               

               

                if (!result.RETURN_FLAG)
                { //20200930 alibaba 回傳舊券是否重複 202008210166-00
                    if (resultSave.REASON_CODE == "1")
                    {
                        result.REASON_CODE = resultSave.REASON_CODE;
                        result.DESCRIPTION = resultSave.DESCRIPTION;
                    } 
                    //end 20200930 alibaba
                    else
                    {
                        result.DESCRIPTION = Message_Type.save_Fail.GetDescription() + " " + resultSave.DESCRIPTION;
                    }
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

        #region DeleteA44
        [BrowserEvent("A44刪除資料")]
        [HttpPost]
        public JsonResult DeleteA44(string bondNumberNew, string lotsNew, string portfolioNameNew)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = A4Repository.deleteA44(bondNumberNew, lotsNew, portfolioNameNew);

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

        #region GetA49Data
        [BrowserEvent("查詢A49資料")]
        [HttpPost]
        public JsonResult GetA49Data(A49ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A4Repository.getA49(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A49DbfileData); //清除
                Cache.Set(CacheList.A49DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheA49Data
        [HttpPost]
        public JsonResult GetCacheA49Data(jqGridParam jdata, string type)
        {
            List<A49ViewModel> data = new List<A49ViewModel>();

            switch (type)
            {
                case "Excel":
                    if (Cache.IsSet(CacheList.A49ExcelfileData))
                    {
                        data = (List<A49ViewModel>)Cache.Get(CacheList.A49ExcelfileData); //從Cache 抓資料
                    }
                    break;
                case "Db":
                    if (Cache.IsSet(CacheList.A49DbfileData))
                    { data = (List<A49ViewModel>)Cache.Get(CacheList.A49DbfileData); }
                    break;
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr", "Lots" }));

        }
        #endregion

        #region UploadA49
        [BrowserEvent("A49上傳檔案")]
        [HttpPost]
        public JsonResult UploadA49()
        {
            MSGReturnModel result = new MSGReturnModel();

            //## 如果有任何檔案類型才做
            if (Request.Files.AllKeys.Any())
            {
                var FileModel = Request.Files["UploadedFile"];

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
                                                 Excel_UploadName.A49.GetDescription(),
                                                 pathType); //固定轉成此名稱

                    Cache.Invalidate(CacheList.A49ExcelName); //清除 Cache
                    Cache.Set(CacheList.A49ExcelName, fileName); //把資料存到 Cache

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

                    var stream = FileModel.InputStream;
                    List<A49ViewModel> dataModel = A4Repository.getA49Excel(pathType, stream);
                    if (dataModel.Count > 0)
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.A49ExcelfileData); //清除 Cache
                        Cache.Set(CacheList.A49ExcelfileData, dataModel); //把資料存到 Cache
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
                return Json(result);
            }

            return Json(result);
        }
        #endregion

        #region TransferA49
        [BrowserEvent("A49 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferA49(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);
                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.A49ExcelName))
                {
                    fileName = (string)Cache.Get(CacheList.A49ExcelName);  //從Cache 抓資料
                }

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();

                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);
                List<A49ViewModel> dataModel = new List<A49ViewModel>();
                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    string pathType = path.Split('.')[1]; //抓副檔名
                    dataModel = A4Repository.getA49Excel(pathType, stream);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A49TransferTxtLog; //預設txt名稱
                string configTxtName = ConfigurationManager.AppSettings["txtLogA49Name"];
                if (!string.IsNullOrWhiteSpace(configTxtName))
                {
                    txtpath = configTxtName; //有設定webConfig且不為空就取代
                }

                #endregion txtlog 檔案名稱

                #region save

                MSGReturnModel resultA49 = A4Repository.saveA49(dataModel,reportDate); //save to DB

                bool A49Log = CommonFunction.saveLog(Table_Type.A49,
                                                     fileName, SetFile.ProgramName, resultA49.RETURN_FLAG,
                                                     Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.A49, resultA49.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save

                result.RETURN_FLAG = resultA49.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                                         .GetDescription(resultA49.DESCRIPTION);
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

        #region 下載 A49範例 Excel
        [HttpGet]
        public ActionResult DownloadA49TempExcel(string type)
        {
            try
            {
                string path = string.Empty;
                if (EnumUtil.GetValues<Excel_DownloadName>()
                    .Any(x => x.ToString().Equals(type)))
                {
                    path = type.GetExelName();
                    return File(ExcelLocation(path), "application/octet-stream", path);
                }
            }
            catch
            {
            }
            return new HttpStatusCodeResult(System.Net.HttpStatusCode.Forbidden);
        }
        #endregion
    }
}