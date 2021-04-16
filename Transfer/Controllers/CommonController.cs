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
using System.Xml.Linq;

namespace Transfer.Controllers
{
    [LogAuth]
    public class CommonController : Controller
    {
        internal ICommon CommonFunction;
        internal ICacheProvider Cache { get; set; }
        public CommonController()
        {
            this.CommonFunction = new Common();
            this.Cache = new DefaultCacheProvider();
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
                    //return File(ExcelLocation(path), "application/vnd.ms-excel", path);
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
                if (Cache.IsSet(getCheckCacheKey(key)))
                {
                    var _cacheData = (Tuple<Check_Table_Type, string>)Cache.Get(getCheckCacheKey(key));
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

        protected string mailLogLocation(string path)
        {
            try
            {
                string projectFile = Server.MapPath("~/mailTxt"); //預設txt位置
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
            data = CommonFunction.getVersion(dt, tableName);
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
        /// <summary>
        /// 獲取登入者可執行類別權限 轉成SelectOption (A,B,M) 
        /// </summary>
        /// <returns></returns>
        protected List<SelectOption> GetDebtSelectOption()
        {
            return CommonFunction.getDebtSelectOption(AccountController.CurrentUserInfo.Name);
        }
        #endregion

        #region 獲取類別權限
        /// <summary>
        /// 獲取登入者類別權限 (A,B,M) 
        /// </summary>
        /// <returns></returns>
        protected string GetDebt()
        {
            return CommonFunction.getUserDebt(AccountController.CurrentUserInfo.Name);
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
            return CommonFunction.getAssessmentInfo(productCode, tableId, type);
        }
        #endregion get 複核者 or 評估者

        #region getCheckData 檢核訊息
        [HttpPost]
        public JsonResult GetCommonCheckData(string Key)
        {
            checkMessageModel result = new checkMessageModel();
            if (Cache.IsSet(getCheckCacheKey(Key)))
            {
                var _cacheData = (Tuple<Check_Table_Type, string>)Cache.Get(getCheckCacheKey(Key));
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
                    if (Cache.IsSet(CacheList.A46ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A46ViewModel>)Cache.Get(CacheList.A46ExcelfileData)));
                    break;
                case "A46Db":
                    if (Cache.IsSet(CacheList.A46DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A46ViewModel>)Cache.Get(CacheList.A46DbfileData)));
                    break;
                case "A47Excel":
                    if (Cache.IsSet(CacheList.A47ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A47ViewModel>)Cache.Get(CacheList.A47ExcelfileData)));
                    break;
                case "A47Db":
                    if (Cache.IsSet(CacheList.A47DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A47ViewModel>)Cache.Get(CacheList.A47DbfileData)));
                    break;
                case "A48Excel":
                    if (Cache.IsSet(CacheList.A48ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A48ViewModel>)Cache.Get(CacheList.A48ExcelfileData)));
                    break;
                case "A48Db":
                    if (Cache.IsSet(CacheList.A48DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A48ViewModel>)Cache.Get(CacheList.A48DbfileData)));
                    break;
                case "A56Db":
                    if (Cache.IsSet(CacheList.A56DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A56ViewModel>)Cache.Get(CacheList.A56DbfileData)));
                    break;
                case "A95_1Excel":
                    if (Cache.IsSet(CacheList.A95_1ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<A95_1ViewModel>)Cache.Get(CacheList.A95_1ExcelfileData)));
                    break;
                case "A95_1Db":
                    if (Cache.IsSet(CacheList.A95_1DbfileData))
                        return Json(jdata.modelToJqgridResult((List<A95_1ViewModel>)Cache.Get(CacheList.A95_1DbfileData)));
                    break;
                case "C04_1":
                    if (Cache.IsSet(CacheList.C04_1DbfileData))
                        return Json(jdata.modelToJqgridResult((List<C04ViewModel>)Cache.Get(CacheList.C04_1DbfileData)));
                    break;
                case "D53Excel":
                    if (Cache.IsSet(CacheList.D53ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<D53ViewModel>)Cache.Get(CacheList.D53ExcelfileData)));
                    break;
                case "D53Db":
                    if (Cache.IsSet(CacheList.D53DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D53ViewModel>)Cache.Get(CacheList.D53DbfileData)));
                    break;
                case "D72Excel":
                    if (Cache.IsSet(CacheList.D72ExcelfileData))
                        return Json(jdata.modelToJqgridResult((List<D72ViewModel>)Cache.Get(CacheList.D72ExcelfileData)));
                    break;
                case "D72Db":
                    if (Cache.IsSet(CacheList.D72DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D72ViewModel>)Cache.Get(CacheList.D72DbfileData)));
                    break;
                case "D73Db":
                    if (Cache.IsSet(CacheList.D73DbfileData))
                        return Json(jdata.modelToJqgridResult((List<D73ViewModel>)Cache.Get(CacheList.D73DbfileData)));
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
                rw.ReportDataSources.Clear();
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    rw.ReportDataSources.Add(new ReportDataSource("DataSet" + (i + 1).ToString(), ds.Tables[i]));
                }
                
                //參數檢查
                XDocument rdlcXml = XDocument.Load(rw.ReportPath);
                XNamespace xmlns = rdlcXml.Root.FirstAttribute.Value;

                foreach (XElement element in rdlcXml.Descendants(xmlns + "ReportParameter"))
                {
                    string parameter = element.FirstAttribute.Value;
                    //rw.ReportDataSources.Add(new ReportDataSource("DataSet1", ds.Tables[0]));
                    if (parameter == "Title")
                    {
                        rw.ReportParameters.Add(new ReportParameter("Title", title));
                    }
                    if (parameter == "ReportTitle")
                    {
                        rw.ReportParameters.Add(new ReportParameter("ReportTitle", "富邦人壽"));
                    }
                    if (parameter == "Emp")
                    {
                        rw.ReportParameters.Add(new ReportParameter("Emp", $@"{AccountController.CurrentUserInfo.Name}({AccountController.CurrentUserName})"));
                    }
                    if (parameter == "Name")
                    {
                        rw.ReportParameters.Add(new ReportParameter("Name", data.className));
                    }
                }

                if (extensionParms != null)
                    rw.ReportParameters.AddRange(extensionParms.Select(x => new ReportParameter(x.key, x.value)));
                if (eparm.Any())
                    rw.ReportParameters.AddRange(eparm.Select(x => new ReportParameter(x.key, x.value)));
                
                rw.IsDownloadDirectly = false;
                //20.03.04 John.下載多張報表修正
                var g = Guid.NewGuid().ToString();
                Session[g] = rw;
                result.RETURN_FLAG = true;
                result.DESCRIPTION = g;
                // Pass report info via session
                //Session["ReportWrapper"] = rw;
                //result.RETURN_FLAG = true;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }
        #endregion




        #region 直接下載報表
        [BrowserEvent("執行下載報表")]
        [HttpPost]
        public JsonResult GetReportFile(
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
                LocalReport rw = new LocalReport();

                rw.ReportPath = Server.MapPath($"~/Report/Rdlc/{data.className}.rdlc");
                rw.DataSources.Clear();
                List<ReportParameter> _parm = new List<ReportParameter>();
                _parm.Add(new ReportParameter("Title", title));

                if (extensionParms != null)
                    _parm.AddRange(extensionParms.Select(x => new ReportParameter(x.key, x.value)));
                if (eparm.Any())
                    _parm.AddRange(eparm.Select(x => new ReportParameter(x.key, x.value)));

                ////參數檢查
                XDocument rdlcXml = XDocument.Load(rw.ReportPath);
                XNamespace xmlns = rdlcXml.Root.FirstAttribute.Value;

                foreach (XElement element in rdlcXml.Descendants(xmlns + "ReportParameter"))
                {
                    string parameter = element.FirstAttribute.Value;
                    //rw.ReportDataSources.Add(new ReportDataSource("DataSet1", ds.Tables[0]));
                    if (parameter == "Title")
                    {
                        _parm.Add(new ReportParameter("Title", title));
                    }
                    if (parameter == "ReportTitle")
                    {
                        _parm.Add(new ReportParameter("ReportTitle", "富邦人壽"));
                    }
                    if (parameter == "Emp")
                    {
                        _parm.Add(new ReportParameter("Emp", $@"{AccountController.CurrentUserInfo.Name}({AccountController.CurrentUserName})"));
                    }
                    if (parameter == "Name")
                    {
                        _parm.Add(new ReportParameter("Name", data.className));
                    }
                }



                if (_parm.Any())
                    rw.SetParameters(_parm);
                for (int i = 0; i < ds.Tables.Count; i++)
                {
                    rw.DataSources.Add(new ReportDataSource("DataSet" + (i + 1).ToString(), ds.Tables[i]));
                }



                // Set DownLoad Name
                string _DisplayName = title;
                var Name = "";
                if (_DisplayName != null)
                {
                    _DisplayName = _DisplayName.Replace("(", "-").Replace(")", "");
                    //var _name = _parm.FirstOrDefault(x => x.Name == "Title");
                    //if (_name != null)
                    //    _DisplayName = $"{_DisplayName}_{_name.Values[0]}";
                    Name = _DisplayName;
                }
                rw.DisplayName = _DisplayName;

                Warning[] warnings;
                string[] streams;
                string mimeType = string.Empty;
                string encoding = string.Empty;
                string extension = string.Empty;
                    byte[] renderedBytes = rw.Render
                        (
                            "PDF",
                            null,
                            out mimeType,
                            out encoding,
                            out extension,
                            out streams,
                            out warnings
                        );

                string fullpath = Path.Combine(Server.MapPath("~/temp"), Name+"."+extension);
                using (var exportdata=new MemoryStream())
                {
                    
                    FileStream file = new FileStream(fullpath, FileMode.Create, FileAccess.Write);
                    exportdata.WriteTo(file);
                    file.Close();
                }



                    Response.Buffer = true;
                Response.Clear();
                Response.ContentType = mimeType;
                Response.AddHeader("content-disposition", "attachment; filename=" + Name + "." + extension);
                Response.BinaryWrite(renderedBytes);
                Response.Flush();
                Response.End();

                //20.03.04 John.下載多張報表修正
                var g = Guid.NewGuid().ToString();
                Session[g] = rw;
                result.RETURN_FLAG = true;
                result.DESCRIPTION = "Success";
                // Pass report info via session
                //Session["ReportWrapper"] = rw;
                //result.RETURN_FLAG = true;
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