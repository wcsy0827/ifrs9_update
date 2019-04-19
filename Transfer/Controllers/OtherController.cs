using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;

namespace Transfer.Controllers
{
    public class OtherController : Controller
    {
        private IOtherRepository OtherRepository;
        public OtherController()
        {
            OtherRepository = new OtherRepository();
        }
        // GET: Other
        public ActionResult RatingSchedule()
        {       
            return View();
        }

        [HttpPost]
        public JsonResult GetData()
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = OtherRepository.getRatingSchedule(null);
            result.Datas = Json(data);
            return Json(result);
        }
    }
}