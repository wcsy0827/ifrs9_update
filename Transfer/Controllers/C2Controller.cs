using System;
using System.Collections.Generic;
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
    public class C2Controller : CommonController
    {
        private IC2Repository C2Repository;

        public C2Controller()
        {
            this.C2Repository = new C2Repository();
        }

        /// <summary>
        /// C23Mortgage(前瞻性篩選變數落後期數資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C23Mortgage()
        {
            List<SelectOption> actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="執行轉換",Value="transfer" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            //List<SelectOption> tableName = new List<SelectOption>() {
            //    new SelectOption() {Text="",Value="" },
            //    new SelectOption() {Text="C13-前瞻性篩選變數資料",Value="C13"},
            //    new SelectOption() {Text="C03-前瞻性模型資料",Value="C03"},
            //    new SelectOption() {Text="C04-前瞻性國外總經資料",Value="C04"}
            //};

            List<SelectOption> tableName = new List<SelectOption>() {
                new SelectOption() {Text="C13-前瞻性篩選變數資料",Value="C13"}
            };
            ViewBag.tableName = new SelectList(tableName, "Value", "Text");

            return View();
        }

        #region GetC23Data
        [BrowserEvent("查詢C23資料")]
        [HttpPost]
        public JsonResult GetC23Data(string tableName, string transferYear)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                string createTableName = $"Econ_D_Var_Lag_{tableName}_{transferYear}";

                switch (tableName)
                {
                    case "C13":
                        var queryC13Result = C2Repository.getC13LagData(createTableName);
                        result.RETURN_FLAG = queryC13Result.Item1;
                        result.colModel = Json(queryC13Result.Item3);
                        result.colNames = Json(queryC13Result.Item4);
                        Cache.Invalidate(CacheList.C13DbfileData);
                        Cache.Set(CacheList.C13DbfileData, queryC13Result.Item2);
                        break;
                    case "C03":
                        var queryC03Result = C2Repository.getC03LagData(createTableName);
                        result.RETURN_FLAG = queryC03Result.Item1;
                        result.colModel = Json(queryC03Result.Item3);
                        result.colNames = Json(queryC03Result.Item4);
                        Cache.Invalidate(CacheList.C03DbfileData);
                        Cache.Set(CacheList.C03DbfileData, queryC03Result.Item2);
                        break;
                    case "C04":
                        var queryC04Result = C2Repository.getC04LagData(createTableName);
                        result.RETURN_FLAG = queryC04Result.Item1;
                        result.colModel = Json(queryC04Result.Item3);
                        result.colNames = Json(queryC04Result.Item4);
                        Cache.Invalidate(CacheList.C04DbfileData);
                        Cache.Set(CacheList.C04DbfileData, queryC04Result.Item2);
                        break;
                    default:
                        break;
                }

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

        #region GetCacheC13Data
        [HttpPost]
        public JsonResult GetCacheC13Data(jqGridParam jdata)
        {
            List<C13ViewModel> data = new List<C13ViewModel>();
            if (Cache.IsSet(CacheList.C13DbfileData))
            {
                data = (List<C13ViewModel>)Cache.Get(CacheList.C13DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetCacheC03Data
        [HttpPost]
        public JsonResult GetCacheC03Data(jqGridParam jdata)
        {
            List<C03ViewModel> data = new List<C03ViewModel>();
            if (Cache.IsSet(CacheList.C03DbfileData))
            {
                data = (List<C03ViewModel>)Cache.Get(CacheList.C03DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetCacheC04Data
        [HttpPost]
        public JsonResult GetCacheC04Data(jqGridParam jdata)
        {
            List<C04ViewModel> data = new List<C04ViewModel>();
            if (Cache.IsSet(CacheList.C04DbfileData))
            {
                data = (List<C04ViewModel>)Cache.Get(CacheList.C04DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetC23Excel
        [BrowserEvent("C23匯出")]
        [HttpPost]
        public ActionResult GetC23Excel(string tableName, string transferYear)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                string createtableName = $"Econ_D_Var_Lag_{tableName}_{transferYear}";

                switch (tableName)
                {
                    case "C13":
                        var dataC13 = C2Repository.getC13LagData(createtableName);
                        path = "C13".GetExelName();
                        result = C2Repository.DownLoadC13Excel(ExcelLocation(path), dataC13.Item2, dataC13.Item3, dataC13.Item4);
                        break;
                    case "C03":
                        var dataC03 = C2Repository.getC03LagData(createtableName);
                        path = "C03".GetExelName();
                        result = C2Repository.DownLoadC03Excel(ExcelLocation(path), dataC03.Item2, dataC03.Item3, dataC03.Item4);
                        break;
                    case "C04":
                        var dataC04 = C2Repository.getC04LagData(createtableName);
                        path = "C04".GetExelName();
                        result = C2Repository.DownLoadC04Excel(ExcelLocation(path), dataC04.Item2, dataC04.Item3, dataC04.Item4);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region GetC23Column
        [BrowserEvent("取得C23資料表欄位")]
        [HttpPost]
        public JsonResult GetC23Column(string tableName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var data = new Tuple<bool, List<C23ViewModel>>(false, new List<C23ViewModel>());

                switch (tableName)
                {
                    case "C13":
                        data = C2Repository.getC13Column();
                        break;
                    case "C03":
                        data = C2Repository.getC03Column();
                        break;
                    case "C04":
                        data = C2Repository.getC04Column();
                        break;
                    default:
                        break;
                }

                result.RETURN_FLAG = data.Item1;
                result.Datas = Json(data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveC23
        [BrowserEvent("C23執行轉換")]
        [HttpPost]
        public JsonResult SaveC23(string lagNumber, string tableName, List<C23ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = new MSGReturnModel();

                string createTableName = $"Econ_D_Var_Lag_{tableName}_{DateTime.Now.ToString("yyyy")}";

                switch (tableName)
                {
                    case "C13":
                        resultSave = C2Repository.transferC13(lagNumber, createTableName, dataModel);
                        break;
                    case "C03":
                        resultSave = C2Repository.transferC03(lagNumber, createTableName, dataModel);
                        break;
                    case "C04":
                        resultSave = C2Repository.transferC04(lagNumber, createTableName, dataModel);
                        break;
                    default:
                        break;
                }

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;
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