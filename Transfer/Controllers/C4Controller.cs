using Microsoft.Reporting.WebForms;
using System.Collections.Generic;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using System.Linq;

namespace Transfer.Controllers
{
    [LogAuth]
    public class C4Controller : CommonController
    {
        private IC4Repository C4Repository;

        public C4Controller()
        {
            this.C4Repository = new C4Repository();
        }

        [UserAuth]
        public ActionResult Index()
        {
            return View();
        }
    }
}