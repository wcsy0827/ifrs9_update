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
    public class A5Controller : CommonController
    {
        private IA5Repository A5Repository;
        private List<SelectOption> searchOption = null;
        private List<SelectOption> sType = null;
        private List<SelectOption> actions = null;

        public A5Controller()
        {
            this.A5Repository = new A5Repository();
            searchOption = new List<SelectOption>();
            searchOption.AddRange(EnumUtil.GetValues<A59_SelectType>()
                        .Select(x => new SelectOption()
                        {
                            Text = x.GetDescription(),
                            Value = x.ToString()
                        }));
            sType = new List<SelectOption>() {
                new SelectOption() {Text=Rating_Type.A.GetDescription(),Value=Rating_Type.A.GetDescription() },
                new SelectOption() {Text=Rating_Type.B.GetDescription(),Value=Rating_Type.B.GetDescription() }};
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢&下載",Value="downLoad" },
                new SelectOption() {Text="上傳&存檔",Value="upLoad" }};
        }

        /// <summary>
        /// A52-信評主標尺對應檔_其他
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A52()
        {
            var _Authority = new Authority(Table_Type.A52.ToString());
            ViewBag.Auditor = new SelectList(_Authority.Auditors
                .Select(x => new SelectOption()
                {
                    Text = $"{x.User_Account}({x.User_Name})",
                    Value = x.User_Account
                }).ToList(), "Value", "Text");
            ViewBag.UserAccount = _Authority.userAccount;
            ViewBag.Authority = _Authority.AuthorityType.GetDescription();
            ViewBag.IsActive = new SelectList(new List<SelectOption>() {
                new SelectOption() { Text = "全部(All)", Value="All"},
                new SelectOption() { Text = "生效", Value = "Y"},
                new SelectOption() { Text = "失效", Value = "N"}
            }, "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult A52Audit()
        {
            var _Authority = new Authority(Table_Type.A52.ToString());
            ViewBag.UserAccount = _Authority.userAccount;
            ViewBag.Authority = _Authority.AuthorityType.GetDescription();
            ViewBag.Auditdate = new SelectList(A52AuditDate(null), "Value", "Text");
            ViewBag.Status = new SelectList(new List<SelectOption>() {
                new SelectOption() { Text = "接受(1)", Value = "1"},
                new SelectOption() { Text = "拒絕(0)", Value = "0"},
            }, "Value", "Text");
            return View();
        }

        /// <summary>
        /// A56 信評設殊值補正設定檔
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A56Detail()
        {
            ViewBag.IsActive = new SelectList(
                new List<SelectOption>()
                {
                    new SelectOption() { Text = "",Value = "" },
                    new SelectOption() { Text = "生效",Value = "Y" },
                    new SelectOption() { Text = "失效",Value = "N" },
                }, "Value", "Text");

            ViewBag.Update_Method =
                EnumUtil.GetValues<Rating_Update_Method>().Select(x =>
                new SelectOption()
                {
                    Text = $"{Rating_Update_MethodText(x)}:{x.GetDescription()}",
                    Value = Rating_Update_MethodText(x)
                }).ToList();
            return View();
        }

        /// <summary>
        /// A57 查詢
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A57()
        {
            ViewBag.sType = new SelectList(
                new List<SelectOption>() {
                    new SelectOption() { Text = "All",Value = "All" },
                    new SelectOption() {Text=Rating_Type.A.GetDescription(),Value=Rating_Type.A.GetDescription() },
                    new SelectOption() {Text=Rating_Type.B.GetDescription(),Value=Rating_Type.B.GetDescription() }}
                , "Value", "Text");
            ViewBag.statusOption = new SelectList(
                EnumUtil.GetValues<Rating_Status>()
                    .Select(x => new SelectOption()
                    {
                        Text = x.GetDescription(),
                        Value = x.ToString()
                    }), "Value", "Text");
            var jqgridInfo = new A57ViewModel().TojqGridData(
                new int[] { 95, 115, 50, 200, 200, 150, 150, 150, 110, 110,105,80,85,
                  80,90,110,105,130,130,150,170,220,150,125,235,150,150,80,225,150,150,150,150});
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// 執行信評轉檔
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A57Detail()
        {
            GetCheckDataToCache();
            ViewBag.complement = new SelectList(
                new List<SelectOption>() {
                    new SelectOption() {Text="補前一版本信評",Value="Y" },
                    new SelectOption() {Text="不抓前一版信評",Value="N" }}
                , "Value", "Text");
            var jqgridInfo = new CheckTableViewModel().TojqGridData(
                new int[] { 120, 80, 180, 90, 100, 100, 100, 100, 300 });
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// 債券信評補登(整理檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A58Detail()
        {
            //ViewBag.searchOption = new SelectList(searchOption, "Value", "Text");
            ViewBag.portfolio = new SelectList(new List<SelectOption>()
            {
                new SelectOption() { Text = "全部", Value = "All" },
                new SelectOption() { Text = "國外固收", Value = "國外固收" },
                new SelectOption() { Text = "國內固收", Value = "國內固收" }
            }, "Value", "Text");
            ViewBag.sType = new SelectList(sType, "Value", "Text");
            ViewBag.action = new SelectList(actions, "Value", "Text");
            return View();
        }

        /// <summary>
        /// 補登畫面
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A59Detail()
        {
            ViewBag.complement = new SelectList(
                new List<SelectOption>() {
                    new SelectOption() {Text="補前一版本信評",Value="Y" },
                    new SelectOption() {Text="不抓前一版信評",Value="N" }}
                , "Value", "Text");
            //190222 顯示jqGrid
            GetCheckDataToCache();
            var jqgridInfo = new CheckTableViewModel().TojqGridData(
                new int[] { 120, 80, 180, 90, 100, 100, 100, 100, 300 });
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        #region A52 相關

        #region GetA52All
        [HttpPost]
        public JsonResult GetA52AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A5Repository.getA52All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A52DbfileData); //清除
                Cache.Set(CacheList.A52DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheA52Data
        [HttpPost]
        public JsonResult GetCacheA52Data(jqGridParam jdata)
        {
            List<A52ViewModel> data = new List<A52ViewModel>();
            if (Cache.IsSet(CacheList.A52DbfileData))
            {
                data = (List<A52ViewModel>)Cache.Get(CacheList.A52DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region A52動態條件查詢
        /// <summary>
        /// A52動態條件查詢
        /// </summary>
        /// <param name="ratingOrg"></param>
        /// <param name="IsActive"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetA52SearchData(string ratingOrg, string IsActive, string pdGrade, string rating)
        {
            var data = A5Repository.GetA52SearchData(ratingOrg, IsActive, pdGrade,rating);
            return Json(data);
        }
        #endregion

        #region 查詢A52 複核時間
        [HttpPost]
        public JsonResult GetA52Auditdate()
        {
            return Json(A5Repository.GetA52Auditdate(null));
        }
        #endregion

        #region GetA52RatingOrg
        [HttpPost]
        public JsonResult GetA52RatingOrg(string ratingOrg)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            try
            {
                var queryResult = A5Repository.getA52(ratingOrg, "All", "All", "All", "All");
                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A52AuditDetailDbfileData); //清除
                Cache.Set(CacheList.A52AuditDetailDbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetA52Data
        [BrowserEvent("查詢A52資料")]
        [HttpPost]
        public JsonResult GetA52Data(string ratingOrg, string pdGrade, string rating, string IsActive)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = A5Repository.getA52(ratingOrg, pdGrade, rating, IsActive, "All");

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.A52DbfileData); //清除
                Cache.Set(CacheList.A52DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetA52AuditData
        [HttpPost]
        public JsonResult GetA52AuditData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            try
            {
                var queryResult = A5Repository.getA52("All", "All", "All", "All", null);
                result.RETURN_FLAG = queryResult.Item1;
                Cache.Invalidate(CacheList.A52AuditDbfileData); //清除
                Cache.Set(CacheList.A52AuditDbfileData, queryResult.Item2); //把資料存到 Cache
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

        #region SaveA52
        [BrowserEvent("儲存A52資料")]
        [HttpPost]
        public JsonResult SaveA52(string actionType, A52ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            try
            {
                result = A5Repository.saveA52(actionType, dataModel);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }

            return Json(result);
        }
        #endregion

        #region A52Audit
        /// <summary>
        /// A52複核資料
        /// </summary>
        /// <param name="ID">A52.ID</param>
        /// <param name="Status">複核結果</param>
        /// <param name="Auditor_Reply">複核者意見</param>
        /// <returns></returns>
        [BrowserEvent("A52複核資料")]
        [HttpPost]
        public JsonResult A52Aduit(string ID, string Status, string Auditor_Reply)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = A5Repository.A52Audit(TypeTransfer.stringToIntN(ID), Status, Auditor_Reply);
            if (result.RETURN_FLAG)
            {
                var queryResult = A5Repository.getA52("All", "All", "All", "All", null);
                Cache.Invalidate(CacheList.A52AuditDbfileData); //清除
                Cache.Set(CacheList.A52AuditDbfileData, queryResult.Item2); //把資料存到 Cache
            }
            return Json(result);
        }
        #endregion

        #region DeleteA52
        [BrowserEvent("A52刪除資料")]
        [HttpPost]
        public JsonResult DeleteA52(A52ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = A5Repository.deleteA52(dataModel);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #endregion

        [HttpPost]
        public JsonResult SearchA56ReplaceObject()
        {
            List<string> result = A5Repository.GetA56_Replace_Object();
            result.Insert(0, " ");
            return Json(result);
        }

        /// <summary>
        /// 查詢A56資料
        /// </summary>
        /// <param name="IsActive">是否生效 Y,N</param>
        /// <param name="ReplaceObject">特殊字元</param>
        /// <returns></returns>
        [BrowserEvent("查詢A56資料")]
        [HttpPost]
        public JsonResult SearchA56(
            string IsActive,
            string ReplaceObject
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var datas = A5Repository.GetA56(IsActive, ReplaceObject);
            Cache.Invalidate(CacheList.A56DbfileData); //清除
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Set(CacheList.A56DbfileData, datas); //把資料存到 Cache
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.A56.tableNameGetDescription());
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢A57資料
        /// </summary>
        /// <param name="datepicker"></param>
        /// <param name="version"></param>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="SMF"></param>
        /// <param name="stype"></param>
        /// <param name="bondNumber"></param>
        /// <param name="issuer"></param>
        /// <param name="status"></param>
        /// <returns></returns>
        [BrowserEvent("查詢A57資料")]
        [HttpPost]
        public JsonResult SearchA57(
            string datepicker,
            string version,
            string from,
            string to,
            string SMF,
            string stype,
            string bondNumber,
            string issuer,
            string status)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dp = DateTime.MinValue;
            DateTime? df = TypeTransfer.stringToDateTimeN(from);
            DateTime? dt = TypeTransfer.stringToDateTimeN(to);
            int v = 0;
            bool flag = false;
            Rating_Status rs = Rating_Status.All;
            EnumUtil.GetValues<Rating_Status>().ToList().ForEach(x =>
            {
                if (x.ToString() == status)
                {
                    rs = x;
                    flag = true;
                }
            });
            if (!flag || !DateTime.TryParse(datepicker, out dp) || !Int32.TryParse(version, out v))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                Cache.Invalidate(CacheList.A57DbfileData); //清除
                var datas = A5Repository.GetA57(dp,v,df,dt,SMF,stype,bondNumber, issuer,rs);
                if (datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Set(CacheList.A57DbfileData, datas); //把資料存到 Cache
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            return Json(result);
        }


        [HttpPost]
        public JsonResult GetSMF(string date, string version)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            int v = 0;
            if (version.IsNullOrWhiteSpace() ||
                !Int32.TryParse(version ,out v) ||
                date.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt))
            {
                return Json(result);
            }
            List<string> data = new List<string>();
            data = A5Repository.getSMF(dt, v);
            if (data.Any())
            {
                data.Insert(0, " ");
                result.RETURN_FLAG = true;
                result.Datas = Json(data.Select(x => new SelectOption()
                {
                    Text = x,
                    Value = x
                }).ToList());
            }
            return Json(result);
        }

        /// <summary>
        /// 查詢A58資料
        /// </summary>
        /// <param name="datepicker">報導日</param>
        /// <param name="sType">評等種類(1:原始投資信評 ,2:評估日最近信評)</param>
        /// <param name="from">購入日期(起始)</param>
        /// <param name="to">購入日期(結束)</param>
        /// <param name="bondNumber">債券編號</param>
        /// <param name="version">資料版本</param>
        /// <param name="search">All:全查, Miss:查缺Grade_Adjust</param>
        /// <param name="portfolio">All,國外固收,國內固收</param>
        /// <returns></returns>
        [BrowserEvent("查詢A58資料")]
        [HttpPost]
        public JsonResult SearchA58(
            string datepicker,
            string sType,
            string from,
            string to,
            string bondNumber,
            string version,
            string search,
            string portfolio)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            try
            {
                if (search.IsNullOrWhiteSpace())
                {
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    return Json(result);
                }
                Cache.Invalidate(CacheList.A58DbMissfileData); //清除
                Cache.Invalidate(CacheList.A58DbfileData); //清除
                var A58Data = A5Repository.GetA58(datepicker, sType, from, to, bondNumber, version, search, portfolio);
                result.RETURN_FLAG = A58Data.Item1;
                if (A58Data.Item1)
                {
                    if (search.IndexOf("Miss") > -1)
                    {                     
                        Cache.Set(CacheList.A58DbMissfileData, A58Data.Item2); //把資料存到 Cache
                    }
                    else
                    {                   
                        Cache.Set(CacheList.A58DbfileData, A58Data.Item2); //把資料存到 Cache
                    }
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.Message;
            }
            return Json(result);
        }

        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="action">downLoad or upLoad</param>
        /// <param name="type">downLoad(All:A58(全查) or Miss:A58(查缺Grade_Adjust) or A59 or CheckTable)</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata, string action, string type)
        {
            List<A58ViewModel> A58Data = new List<A58ViewModel>();
            List<A59ViewModel> A59Data = new List<A59ViewModel>();
            List<A57ViewModel> A57Data = new List<A57ViewModel>();
            switch (type)
            {
                case "Miss":
                    if (Cache.IsSet(CacheList.A58DbMissfileData))
                        A58Data = (List<A58ViewModel>)Cache.Get(CacheList.A58DbMissfileData);  //從Cache 抓資料
                    break;

                case "All":
                    if (Cache.IsSet(CacheList.A58DbfileData))
                        A58Data = (List<A58ViewModel>)Cache.Get(CacheList.A58DbfileData);
                    break;

                case "A59":
                    if (Cache.IsSet(CacheList.A59ExcelfileData))
                        A59Data = (List<A59ViewModel>)Cache.Get(CacheList.A59ExcelfileData);
                    break;
                case "CheckTable":
                    if (Cache.IsSet(CacheList.CheckTableDbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<CheckTableViewModel>)Cache.Get(CacheList.CheckTableDbfileData)));
                    break;
                case "CheckTable2":
                    if (Cache.IsSet(CacheList.CheckTableDbfileData2))
                        return Json(jdata.modelToJqgridResult(
                            (List<CheckTableViewModel>)Cache.Get(CacheList.CheckTableDbfileData2)));
                    break;
                case "A57":
                    if (Cache.IsSet(CacheList.A57DbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<A57ViewModel>)Cache.Get(CacheList.A57DbfileData)));
                    break;
                case "A52Aduit":
                    if (Cache.IsSet(CacheList.A52AuditDbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<A52ViewModel>)Cache.Get(CacheList.A52AuditDbfileData)));
                    break;
                case "A52AduitDetail":
                    if (Cache.IsSet(CacheList.A52AuditDetailDbfileData))
                        return Json(jdata.modelToJqgridResult(
                            (List<A52ViewModel>)Cache.Get(CacheList.A52AuditDetailDbfileData)));
                    break;
            }
            if (action.Equals("downLoad"))
                return Json(jdata.modelToJqgridResult(A58Data)); //下載查詢資料
            return Json(jdata.modelToJqgridResult(A59Data)); //上傳查詢資料
        }

        /// <summary>
        /// 下載A56Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A56Excel檔案")]
        [HttpPost]
        public JsonResult GetA56Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (Cache.IsSet(CacheList.A56DbfileData))
            {
                var A56 = Excel_DownloadName.A56.ToString();
                var A56Data = (List<A56ViewModel>)Cache.Get(CacheList.A56DbfileData);  //從Cache 抓資料
                result = A5Repository.DownLoadExcel(A56, ExcelLocation(A56.GetExelName()), A56Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            return Json(result);
        }

        /// <summary>
        ///下載A52Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A52Excel檔案")]
        [HttpPost]
        public JsonResult GetA52Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (Cache.IsSet(CacheList.A52DbfileData))
            {
                var A52 = Excel_DownloadName.A52.ToString();
                var A52Data = (List<A52ViewModel>)Cache.Get(CacheList.A52DbfileData);  //從Cache 抓資料
                result = A5Repository.DownLoadExcel(A52, ExcelLocation(A52.GetExelName()), A52Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }

        /// <summary>
        ///下載A52AuditExcel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A52AuditExcel檔案")]
        [HttpPost]
        public JsonResult GetA52AuditExcel(string Audit_date)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var A52Audit = Excel_DownloadName.A52Audit.ToString();
            var A52AuditData = A5Repository.GetA52byAuditdate(Audit_date);  //從Cache 抓資料
            result = A5Repository.DownLoadExcel(A52Audit, ExcelLocation(A52Audit.GetExelName()), A52AuditData);
            return Json(result);
        }

        /// <summary>
        ///下載A57Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載A57Excel檔案")]
        [HttpPost]
        public JsonResult GetA57Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.A57DbfileData))
            {
                var A57 = Excel_DownloadName.A57.ToString();
                var A57Data = (List<A57ViewModel>)Cache.Get(CacheList.A57DbfileData);  //從Cache 抓資料
                result = A5Repository.DownLoadExcel(A57, ExcelLocation(A57.GetExelName()), A57Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }

        /// <summary>
        /// 下載缺漏A59Excel檔案
        /// </summary>
        /// <param name="type">A59(原A59),A59b(熟高、熟低)</param>
        /// <returns></returns>
        [BrowserEvent("下載A59Excel檔案")]
        [HttpPost]
        public JsonResult GetA59Excel(string type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.A58DbMissfileData))
            {
                var A58Data = (List<A58ViewModel>)Cache.Get(CacheList.A58DbMissfileData);  //從Cache 抓資料
                if (type == Excel_DownloadName.A59.ToString())
                {
                    result = A5Repository.DownLoadExcel(type, ExcelLocation(type.GetExelName()), A58Data);
                }
                if (type == Excel_DownloadName.A59b.ToString())
                {
                    result = A5Repository.DownLoadExcel(type, ExcelLocation(type.GetExelName()), A58Data);
                }
            }
            else if (Cache.IsSet(CacheList.A58DbfileData))
            {
                var A58Data = (List<A58ViewModel>)Cache.Get(CacheList.A58DbfileData);  //從Cache 抓資料
                if (type == Excel_DownloadName.A59.ToString())
                {
                    result = A5Repository.DownLoadExcel(type, ExcelLocation(type.GetExelName()), A58Data, true);
                }
                if (type == Excel_DownloadName.A59b.ToString())
                {
                    result = A5Repository.DownLoadExcel(type, ExcelLocation(type.GetExelName()), A58Data, true);
                }
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns>MSGReturnModel</returns>
        [BrowserEvent("上傳A59Excel檔案")]
        [HttpPost]
        public JsonResult UploadA59(ValidateFiles FileModel)
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
                    Excel_UploadName.A59.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.A59ExcelName); //清除 Cache
                Cache.Set(CacheList.A59ExcelName, fileName); //把資料存到 Cache

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
                List<A59ViewModel> dataModel = new List<A59ViewModel>();                
                var A59result =  A5Repository.getA59Excel(pathType, path, "upload");
                if (A59result.Item1.IsNullOrWhiteSpace())
                    dataModel = A59result.Item2;
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = A59result.Item1;
                    return Json(result);
                }
                if (dataModel.Count > 0)
                {
                    result.RETURN_FLAG = true;
                    var A59 = new A59ViewModel();
                    var jqgridParams = A59.TojqGridData();
                    result.Datas = Json(jqgridParams);
                    Cache.Invalidate(CacheList.A59ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.A59ExcelfileData, dataModel); //把資料存到 Cache
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

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("A59缺漏資料存入A57&A58")]
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
                if (Cache.IsSet(CacheList.A59ExcelName))
                    fileName = (string)Cache.Get(CacheList.A59ExcelName);  //從Cache 抓資料
                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                string pathType = path.Split('.')[1]; //抓副檔名
                List<A59ViewModel> dataModel = new List<A59ViewModel>();
                var A59result = A5Repository.getA59Excel(pathType, path, "Transfer");
                if (A59result.Item1.IsNullOrWhiteSpace())
                    dataModel = A59result.Item2;
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = A59result.Item1;
                    return Json(result);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.A59TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save 資料

                #region save A59(A59=>A57=>A58)

                MSGReturnModel resultA59 = A5Repository.saveA59(dataModel); //save to DB
                bool A59Log = CommonFunction.saveLog(Table_Type.A59, fileName, SetFile.ProgramName,
                    resultA59.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name,0); //寫sql Log
                TxtLog.txtLog(Table_Type.A59, resultA59.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save A59(A59=>A57=>A58)

                result = resultA59;

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
        /// 手動轉換 A57 & A58
        /// </summary>
        /// <param name="date">reportDate</param>
        /// <param name="version">version</param>
        /// <param name="complement">補登信評</param>
        /// <param name="deleteFlag">已經執行過信評轉檔,重新執行</param>
        /// <returns></returns>
        [BrowserEvent("手動轉檔A57&A58")]
        [HttpPost]
        public JsonResult TransferA57A58(string date, string version,string complement,bool deleteFlag,bool A59Flag)
        {
            MSGReturnModel result = new MSGReturnModel();

            DateTime dt = DateTime.MinValue;
            int ver = 0;
            if (date.IsNullOrWhiteSpace() || version.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt) || !Int32.TryParse(version, out ver))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                result = A5Repository.saveA57A58(dt, ver, complement, deleteFlag, A59Flag);
                GetCheckDataToCache();
            }
            return Json(result);
        }

        /// <summary>
        /// 變更A56信評設殊值補正設定檔
        /// </summary>
        /// <param name="model">A56ViewModel</param>
        /// <param name="IsActive">true = 新增 false = 刪除</param>
        /// <param name="IsActiveStr">查詢 IsActive 參數</param>
        /// <param name="ReplaceObject">查詢 ReplaceObject 參數</param>
        /// <returns></returns>
        [BrowserEvent("變更A56信評設殊值補正設定檔")]
        [HttpPost]
        public JsonResult SaveA56(
            A56ViewModel model, 
            bool IsActive, 
            string IsActiveStr,
            string ReplaceObject)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = A5Repository.saveA56(model, IsActive);
            if (result.RETURN_FLAG)
            {
                Cache.Invalidate(CacheList.A56DbfileData); //清除
                var datas = A5Repository.GetA56(IsActiveStr, ReplaceObject);
                if (datas.Any())
                {
                    Cache.Set(CacheList.A56DbfileData, datas); //把資料存到 Cache
                }
            }
            return Json(result);
        }

        /// <summary>
        /// 判斷是否有A53 有就不能執行 A59直接補檔
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetA53 (string date)
        {
            MSGReturnModel result = new MSGReturnModel();

            DateTime dt = DateTime.MinValue;
            if (date.IsNullOrWhiteSpace()||!DateTime.TryParse(date, out dt) )
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                result.RETURN_FLAG = A5Repository.getA53(dt);
            }
            return Json(result);
        }

        #region  190222 自動補入A57、A58原始信評資料
        /// <summary>
        /// 自動補入A57、A58原始信評資料
        /// </summary>
        /// <param name="date"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        /// 
        [BrowserEvent("自動補入A57、A58原始信評資料")]
        [HttpPost]
        public JsonResult autoTransfer(string datepicker, string version)
        {
            string sType = "原始投資信評";
            string portfolio = "All";
            string search = "Miss";
            string from = "";
            string to = "";
            string bondNumber = "";
            MSGReturnModel result_searchA58 = new MSGReturnModel();
            MSGReturnModel result_getA59 = new MSGReturnModel();
            MSGReturnModel result_saveA59Excel = new MSGReturnModel();
            MSGReturnModel result_autoTransfer = new MSGReturnModel();
            result_searchA58.RETURN_FLAG = false;
            result_getA59.RETURN_FLAG = false;
            result_autoTransfer.RETURN_FLAG = false;
            #region 找A58缺漏
            var A58Data = A5Repository.GetA58(datepicker, sType, from, to, bondNumber, version, search, portfolio);
            result_searchA58.RETURN_FLAG = A58Data.Item1;
            #endregion
            #region 轉入寶碩信評資料
            if (!A58Data.Item1)
            {
                result_searchA58.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                return Json(result_searchA58);
            }
            var A59Data = A5Repository.getA59(A58Data.Item2, datepicker);
            if (!A59Data.Item1)
            {
                result_getA59.DESCRIPTION = Message_Type.not_Find_CounterPartyCreditRating.GetDescription();
                GetCheckDataToCache();
                return Json(result_getA59);
            }

            #region 產excel出來
            var A59Filled = Excel_DownloadName.A59Filled.ToString();
            var fileName = string.Format("{0}.{1}", Excel_DownloadName.A59Filled.GetDescription(), "xlsx"); //固定轉成此名稱

            //檢查資料夾是否存在
            string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
            string path = Path.Combine(projectFile, fileName);
            FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

            result_saveA59Excel = A5Repository.SaveA59Excel(A59Filled, path, A59Data.Item2);
            //儲存excel失敗不用中止，只要記錄就好         
            DateTime startTime = DateTime.Now;
            DateTime reportdate = DateTime.MinValue;
            DateTime.TryParse(datepicker, out reportdate);
            bool A59Log = CommonFunction.saveLog(Table_Type.A59, fileName, SetFile.ProgramName,
            result_saveA59Excel.RETURN_FLAG, Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name, Convert.ToInt32(version), reportdate); //寫sql Log

            #endregion 測試產excel出來看

            #endregion

            #region 更新A57、A58
            #region txtlog 檔案名稱

            string txtpath = SetFile.A59TransferTxtLog; //預設txt名稱

            #endregion txtlog 檔案名稱

            #region save A59(A59=>A57=>A58)
            DateTime st_saveA59 = DateTime.Now;
            result_autoTransfer = A5Repository.saveA59(A59Data.Item2);
            A5Repository.SaveA59TransLog(result_autoTransfer, datepicker, st_saveA59, Convert.ToInt32(version));
            TxtLog.txtLog(Table_Type.A59Trans, result_autoTransfer.RETURN_FLAG, st_saveA59, txtLocation(txtpath)); //寫txt Log

            #endregion save A59(A59=>A57=>A58)
            #endregion 更新A57、A58
            #region 執行檢核
            A5Repository.GetA58TransferCheck(reportdate, Convert.ToInt32(version));
            GetCheckDataToCache();
            #endregion 執行檢核                                       

            return Json(result_autoTransfer);

        }

        #endregion


        [HttpPost]
        public JsonResult CheckVerison(string date, string version)
        {
            MSGReturnModel result = new MSGReturnModel();

            DateTime dt = DateTime.MinValue;
            int ver = 0;
            if (date.IsNullOrWhiteSpace() || version.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt) || !Int32.TryParse(version, out ver))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                result = A5Repository.checkVersion(dt, ver);
            }
            return Json(result);
        }

        private string Rating_Update_MethodText(Rating_Update_Method parm)
        {
            return ((int)parm).ToString().PadLeft(2, '0');
        }

        //190222 cache增加顯示A59
        private void GetCheckDataToCache()
        {
            var dataModel = A5Repository.getCheckTable(new List<string>() { Table_Type.A41.ToString() });
            Cache.Invalidate(CacheList.CheckTableDbfileData); //清除 Cache
            Cache.Set(CacheList.CheckTableDbfileData, dataModel); //把資料存到 Cache
            var dataModel2 = A5Repository.getCheckTable(new List<string>()
            {
                Table_Type.A53.ToString(),
                Table_Type.A57.ToString(),
                Table_Type.A58.ToString(),
                Table_Type.A59.ToString(),
                Table_Type.A59Apex.ToString(),
                Table_Type.A59Trans.ToString()
            });
            Cache.Invalidate(CacheList.CheckTableDbfileData2); //清除 Cache
            Cache.Set(CacheList.CheckTableDbfileData2, dataModel2); //把資料存到 Cache
        }

        private List<SelectOption> A52AuditDate(string Audit_Date)
        {
            return A5Repository.GetA52Auditdate(Audit_Date);
        }
    }
}