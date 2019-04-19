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
    public class A0Controller : CommonController
    {
        private A0Repository A0Repository;
        public ICacheProvider Cache { get; set; }

        public A0Controller()
        {
            this.A0Repository = new A0Repository();
            this.Cache = new DefaultCacheProvider();
        }

        /// <summary>
        /// A08(房貸其他揭露資訊)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult A08()
        {
            string[] aligns = new string[] {"left", "left", "left", "left", "right"};
            var jqgridInfo = new A08ViewModel().TojqGridData(null,aligns);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;

            return View();
        }

        #region QueryA08
        [BrowserEvent("查詢A08資料")]
        [HttpPost]
        public JsonResult QueryA08(string referenceNbr, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = A0Repository.queryA08(referenceNbr, reportDate);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.A08DbfileData); //清除
                Cache.Set(CacheList.A08DbfileData, queryData.Item2); //把資料存到 Cache

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

        #region GetCacheA08Data
        [HttpPost]
        public JsonResult GetCacheA08Data(jqGridParam jdata)
        {
            List<A08ViewModel> data = new List<A08ViewModel>();
            if (Cache.IsSet(CacheList.A08DbfileData))
            {
                data = (List<A08ViewModel>)Cache.Get(CacheList.A08DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr" }));
        }
        #endregion
    }
}