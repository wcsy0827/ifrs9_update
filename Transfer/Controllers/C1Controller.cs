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
    public class C1Controller : CommonController
    {
        private IC1Repository C1Repository;

        public C1Controller()
        {
            this.C1Repository = new C1Repository();
        }

        /// <summary>
        /// C13Mortgage(前瞻性篩選變數資料)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult C13Mortgage()
        {
            List<SelectOption> selectOption = new List<SelectOption>()
                                              {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "Q1", Value = "Q1" },
                                                  new SelectOption() { Text = "Q2", Value = "Q2" },
                                                  new SelectOption() { Text = "Q3", Value = "Q3" },
                                                  new SelectOption() { Text = "Q4", Value = "Q4" }
                                              };
            ViewBag.Quartly = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>()
                           {
                               new SelectOption() { Text = "", Value = "" },
                               new SelectOption() { Text = "Y", Value = "Y" },
                               new SelectOption() { Text = "N", Value = "N" }
                           };
            ViewBag.YN = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        #region GetC13All
        [HttpPost]
        public JsonResult GetC13AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = C1Repository.getC13All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.C13DbfileData); //清除
                Cache.Set(CacheList.C13DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetC13Data
        [BrowserEvent("查詢C13資料")]
        [HttpPost]
        public JsonResult GetC13Data(C13ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = C1Repository.getC13(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.C13DbfileData); //清除
                Cache.Set(CacheList.C13DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveC13
        [BrowserEvent("儲存C13資料")]
        [HttpPost]
        public JsonResult SaveC13(string actionType, C13ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = C1Repository.saveC13(actionType, dataModel);

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

        #region DeleteC13
        [BrowserEvent("刪除C13資料")]
        [HttpPost]
        public JsonResult DeleteC13(string yearQuartly)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = C1Repository.deleteC13(yearQuartly);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription();
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