using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;
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
    public class ScheduleController : CommonController
    {
        private ISystemForITRepository SystemForITRepository;

        public ScheduleController()
        {
            this.SystemForITRepository = new SystemForITRepository();
        }

        /// <summary>
        /// 設定主頁 (預留)
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// 排程重啟
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult TaskSchedule()
        {
            return View();
        }

        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata)
        {
            List<AccountViewModel> datas = new List<AccountViewModel>();
            if (Cache.IsSet(CacheList.userDbData))
                datas = (List<AccountViewModel>)Cache.Get(CacheList.userDbData);  //從Cache 抓資料
            return Json(jdata.modelToJqgridResult(datas)); //查詢資料
        }        

        #region GetTaskSchedule
        public JsonResult GetTaskSchedule()
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = SystemForITRepository.getTaskSchedule(null,false);
            result.Datas = Json(data);
            return Json(result);
        }
        #endregion

        #region RestartTaskSchedule
        [HttpPost]
        public JsonResult RestartTaskSchedule(string triggerTask, string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultRestart = SystemForITRepository.restartTaskSchedule(triggerTask, reportDate);

                result.RETURN_FLAG = resultRestart.RETURN_FLAG;
                result.DESCRIPTION = "重啟完成";

                if (result.RETURN_FLAG == false)
                {
                    result.DESCRIPTION = "重啟失敗：" + resultRestart.DESCRIPTION;
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