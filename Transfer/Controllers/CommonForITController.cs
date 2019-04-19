using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using static Transfer.Enum.Ref;
using Microsoft.Reporting.WebForms;
using System.Data;
using System.Reflection;
using Transfer.ViewModels;

namespace Transfer.Controllers
{
    public class CommonForITController : Controller
    {
        internal ICommonForIT CommonFunctionForIT;
        internal ICacheProviderForIT CacheForIT { get; set; }
        public CommonForITController()
        {
            this.CommonFunctionForIT = new CommonForIT();
            this.CacheForIT = new CacheProviderForIT();
        }

        #region 下載 Excel
        [HttpGet]
        [DeleteFile]
        public ActionResult DownloadExcel(string type)
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

        #region 下載檢核訊息
        [HttpGet]
        public ActionResult DownloadMessage(string key)
        {
            try
            {
                if (CacheForIT.IsSet(getCheckCacheKey(key)))
                {
                    var _cacheData = (Tuple<Check_Table_Type, string>)CacheForIT.Get(getCheckCacheKey(key));
                    if (_cacheData.Item1.ToString() == key)
                    {
                        return File(System.Text.Encoding.UTF8.GetBytes(_cacheData.Item2),
                            "application/octet-stream",
                            $@"{_cacheData.Item1.GetDescription()}.txt");
                    }
                }
            }
            catch
            {

            }
            return new HttpStatusCodeResult(System.Net.HttpStatusCode.Forbidden);
        }
        #endregion

        #region txtlog 設定位置
        protected string txtLocation(string path)
        {
            try
            {
                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //預設txt位置
                string configTxtLocation = ConfigurationManager.AppSettings["txtLogLocation"];
                if (!string.IsNullOrWhiteSpace(configTxtLocation))
                    projectFile = configTxtLocation; //有設定webConfig且不為空就取代
                FileRelated.createFile(projectFile);
                string folderPath = Path.Combine(projectFile, path); //合併路徑&檔名
                return folderPath;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion txtlog 設定位置

        #region 獲取版本
        [HttpPost]
        public JsonResult GetVersion(string date, string tableName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (tableName.IsNullOrWhiteSpace() ||
                !EnumUtil.GetValues<Transfer_Table_Type>()
                .Any(x => x.ToString() == tableName) ||
                date.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt))
            {
                return Json(result);
            }
            List<string> data = new List<string>();
            data = CommonFunctionForIT.getVersion(dt, tableName);
            if (data.Any())
            {
                data.Insert(0, " ");
                result.RETURN_FLAG = true;
                result.Datas = Json(data);
            }
            return Json(result);
        }
        #endregion

        #region 獲取類別權限 (SelectOption)
        protected List<SelectOption> GetDebtSelectOption()
        {
            return CommonFunctionForIT.getDebtSelectOption("");
        }
        #endregion

        #region 獲取類別權限
        protected string GetDebt()
        {
            return CommonFunctionForIT.getUserDebt("");
        }
        #endregion

        #region Excel 設定下載位置

        protected string ExcelLocation(string path)
        {
            try
            {
                string projectFile = Server.MapPath("~/" + SetFile.FileDownloads); //預設Excel下載位置
                string configExcelLocation = ConfigurationManager.AppSettings["ExcelDlLocation"];
                if (!string.IsNullOrWhiteSpace(configExcelLocation))
                    projectFile = configExcelLocation; //有設定webConfig且不為空就取代
                FileRelated.createFile(projectFile);
                string folderPath = Path.Combine(projectFile, path); //合併路徑&檔名
                return folderPath;
            }
            catch
            {
                return string.Empty;
            }
        }

        #endregion Excel 設定下載位置

        #region  get 複核者 or 評估者
        public List<IFRS9_User> getAssessment(string productCode, string tableId, SetAssessmentType type)
        {
            return CommonFunctionForIT.getAssessmentInfo(productCode, tableId, type);
        }
        #endregion get 複核者 or 評估者

        #region getCheckData 檢核訊息
        [HttpPost]
        public JsonResult GetCommonCheckData(string Key)
        {
            checkMessageModel result = new checkMessageModel();
            if (CacheForIT.IsSet(getCheckCacheKey(Key)))
            {
                var _cacheData = (Tuple<Check_Table_Type, string>)CacheForIT.Get(getCheckCacheKey(Key));
                if (_cacheData.Item1.ToString() == Key)
                {
                    result.message = _cacheData.Item2;
                    result.title = _cacheData.Item1.GetDescription();
                }
            }
            return Json(result);
        }
        #endregion

        #region GetCacheData 
        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetDefaultCacheData(jqGridParam jdata, string type)
        {
            switch (type)
            {
                case "A46Excel":
                    if (CacheForIT.IsSet(CacheList.A46ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A46ViewModel>)CacheForIT.Get(CacheList.A46ExcelfileData)));
                    break;
                case "A46Db":
                    if (CacheForIT.IsSet(CacheList.A46DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A46ViewModel>)CacheForIT.Get(CacheList.A46DbfileData)));
                    break;
                case "A47Excel":
                    if (CacheForIT.IsSet(CacheList.A47ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A47ViewModel>)CacheForIT.Get(CacheList.A47ExcelfileData)));
                    break;
                case "A47Db":
                    if (CacheForIT.IsSet(CacheList.A47DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A47ViewModel>)CacheForIT.Get(CacheList.A47DbfileData)));
                    break;
                case "A48Excel":
                    if (CacheForIT.IsSet(CacheList.A48ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A48ViewModel>)CacheForIT.Get(CacheList.A48ExcelfileData)));
                    break;
                case "A48Db":
                    if (CacheForIT.IsSet(CacheList.A48DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A48ViewModel>)CacheForIT.Get(CacheList.A48DbfileData)));
                    break;
                case "A56Db":
                    if (CacheForIT.IsSet(CacheList.A56DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A56ViewModel>)CacheForIT.Get(CacheList.A56DbfileData)));
                    break;
                case "A95_1Excel":
                    if (CacheForIT.IsSet(CacheList.A95_1ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A95_1ViewModel>)CacheForIT.Get(CacheList.A95_1ExcelfileData)));
                    break;
                case "A95_1Db":
                    if (CacheForIT.IsSet(CacheList.A95_1DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A95_1ViewModel>)CacheForIT.Get(CacheList.A95_1DbfileData)));
                    break;
                case "D53Excel":
                    if (CacheForIT.IsSet(CacheList.D53ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<D53ViewModel>)CacheForIT.Get(CacheList.D53ExcelfileData)));
                    break;
                case "D53Db":
                    if (CacheForIT.IsSet(CacheList.D53DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D53ViewModel>)CacheForIT.Get(CacheList.D53DbfileData)));
                    break;
                case "D72Excel":
                    if (CacheForIT.IsSet(CacheList.D72ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<D72ViewModel>)CacheForIT.Get(CacheList.D72ExcelfileData)));
                    break;
                case "D72Db":
                    if (CacheForIT.IsSet(CacheList.D72DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D72ViewModel>)CacheForIT.Get(CacheList.D72DbfileData)));
                    break;
                case "D73Db":
                    if (CacheForIT.IsSet(CacheList.D73DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D73ViewModel>)CacheForIT.Get(CacheList.D73DbfileData)));
                    break;
            }
            return Json(null);
        }
        #endregion

        #region 執行報表
        [BrowserEvent("執行報表")]
        [HttpPost]
        public JsonResult GetReport(
            reportModel data,
            List<reportParm> parms,
            List<reportParm> extensionParms = null)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            try
            {
                string title = "(報表名稱)";
                if (data.className.IsNullOrWhiteSpace())
                {
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(null, "無呼叫的className");
                    return Json(result);
                }
                if (!data.title.IsNullOrWhiteSpace())
                    title = data.title;
                //object obj = AppDomain.CurrentDomain.CreateInstanceAndUnwrap("Transfer", $"Transfer.Report.Data.{data.className}");
                object obj = Activator.CreateInstance(Assembly.Load("Transfer").GetType($"Transfer.Report.Data.{data.className}"));
                MethodInfo[] methods = obj.GetType().GetMethods();
                MethodInfo mi = methods.FirstOrDefault(x => x.Name == "GetData");
                if (mi == null)
                {
                    //請檢查是否實作資料獲取
                    result.DESCRIPTION = "報表錯誤請聯絡IT人員!";
                    return Json(result);
                }
                DataSet ds = (DataSet)mi.Invoke(obj, new object[] { parms });
                List<reportParm> eparm = (List<reportParm>)(obj.GetType().GetProperty("extensionParms").GetValue(obj));
                // Set report info
                ReportWrapper rw = new ReportWrapper();

                rw.ReportPath = Server.MapPath($"~/Report/Rdlc/{data.className}.rdlc");
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    rw.ReportDataSources.Add(new ReportDataSource("DataSet" + (i + 1).ToString(), ds.Tables[i]));
                }
                rw.ReportParameters.Add(new ReportParameter("Title", title));
                rw.ReportParameters.Add(new ReportParameter("ReportTitle", "富邦人壽"));
                rw.ReportParameters.Add(new ReportParameter("Emp", $@"{AccountController.CurrentUserInfo.Name}({AccountController.CurrentUserName})"));
                rw.ReportParameters.Add(new ReportParameter("Name", data.className));
                if (extensionParms != null)
                    rw.ReportParameters.AddRange(extensionParms.Select(x => new ReportParameter(x.key, x.value)));
                if (eparm.Any())
                    rw.ReportParameters.AddRange(eparm.Select(x => new ReportParameter(x.key, x.value)));
                rw.IsDownloadDirectly = false;
                // Pass report info via session
                Session["ReportWrapper"] = rw;
                result.RETURN_FLAG = true;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }
        #endregion

        public class reportModel
        {
            public string title { get; set; }
            public string className { get; set; }
        }

        public class checkMessageModel
        {
            public checkMessageModel()
            {
                title = string.Empty;
                message = string.Empty;
            }
            public string title { get; set; }
            public string message { get; set; }
        }

        private string getCheckCacheKey(string key)
        {
            if (key.IsNullOrWhiteSpace())
                return string.Empty;
            return $@"{CacheList.CheckData}{key}";
        }
    }
}