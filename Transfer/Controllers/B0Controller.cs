using System;
using System.Collections.Generic;
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
    public class B0Controller : CommonController
    {
        private IB0Repository B0Repository;

        public B0Controller()
        {
            this.B0Repository = new B0Repository();
        }

        /// <summary>
        /// B06前膽性參數-風控
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult B06Bond()
        {
            List<string> PRJID = B0Repository.getB06PRJID("Bond");
            PRJID.Insert(0, "");
            ViewBag.PRJID = new SelectList(PRJID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<string> FLOWID = new List<string>();
            FLOWID.Insert(0, "");
            ViewBag.FLOWID = new SelectList(FLOWID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<string> CompID = new List<string>();
            CompID.Insert(0, "");
            ViewBag.CompID = new SelectList(CompID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            return View();
        }

        /// <summary>
        /// B06前膽性參數-房貸
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult B06Loan()
        {
            List<string> PRJID = B0Repository.getB06PRJID("loan");
            PRJID.Insert(0, "");
            ViewBag.PRJID = new SelectList(PRJID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<string> FLOWID = new List<string>();
            FLOWID.Insert(0, "");
            ViewBag.FLOWID = new SelectList(FLOWID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            List<string> CompID = new List<string>();
            CompID.Insert(0, "");
            ViewBag.CompID = new SelectList(CompID.Select(x => new { Text = x, Value = x }), "Value", "Text");

            return View();
        }

        #region GetB06PRJID
        [HttpPost]
        public JsonResult GetB06PRJID(string productCode)
        {
            List<string> Datas = B0Repository.getB06PRJID(productCode);
            return Json(string.Join(",", Datas));
        }
        #endregion

        #region GetB06FLOWID
        [HttpPost]
        public JsonResult GetB06FLOWID(string productCode, string prjid)
        {
            List<string> Datas = B0Repository.getB06FLOWID(productCode,prjid);
            return Json(string.Join(",", Datas));
        }
        #endregion

        #region GetB06CompID
        [HttpPost]
        public JsonResult GetB06CompID(string productCode, string prjid, string flowid)
        {
            List<string> Datas = B0Repository.getB06CompID(productCode, prjid, flowid);
            return Json(string.Join(",", Datas));
        }
        #endregion

        #region GetB06All
        [HttpPost]
        public JsonResult GetB06AllData(B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = B0Repository.getB06All(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.B06DbfileData); //清除
                Cache.Set(CacheList.B06DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheB06Data
        [HttpPost]
        public JsonResult GetCacheB06Data(jqGridParam jdata)
        {
            List<B06ViewModel> data = new List<B06ViewModel>();
            if (Cache.IsSet(CacheList.B06DbfileData))
            {
                data = (List<B06ViewModel>)Cache.Get(CacheList.B06DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetB06Data
        [BrowserEvent("查詢B06資料")]
        [HttpPost]
        public JsonResult GetB06Data(B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = B0Repository.getB06(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.B06DbfileData); //清除
                Cache.Set(CacheList.B06DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveB06
        [BrowserEvent("儲存B06資料")]
        [HttpPost]
        public JsonResult SaveB06(string actionType, B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = B0Repository.saveB06(actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = resultSave.DESCRIPTION;
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

        #region DeleteB06
        [BrowserEvent("B06刪除資料")]
        [HttpPost]
        public JsonResult DeleteB06(B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = B0Repository.deleteB06(dataModel);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = resultDelete.DESCRIPTION;
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
    }
}