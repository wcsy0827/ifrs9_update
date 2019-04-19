using Microsoft.Reporting.WebForms;
using System.Collections.Generic;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using System.Linq;
using static Transfer.Enum.Ref;

namespace Transfer.Controllers
{
    [LogAuth]
    public class ReportController : CommonController
    {
        private IKriskRepository KriskRepository;

        public ReportController()
        {
            this.KriskRepository = new KriskRepository();
        }

        [UserAuth]
        public ActionResult Impairment01()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult Impairment02()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            ViewBag.AssetSeg = new SelectList(new List<SelectOption>() {
                new SelectOption() { Text = "排除外幣&OIU",Value = "Default"},
                new SelectOption() { Text = "外幣",Value = "外幣"},
                new SelectOption() { Text = "OIU",Value = "OIU"}
            }, "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult Impairment03()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult Impairment04()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult Impairment05()
        {
            return View();
        }

        [UserAuth]
        public ActionResult Impairment06()
        {
            return View();
        }

        [UserAuth]
        public ActionResult D02_1()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_2()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_3()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_4()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_4_1()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_4_2()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult D02_5()
        {
            return View();
        }

        [UserAuth]
        public ActionResult RiskControl01()
        {
            return View();
        }

        [UserAuth]
        public ActionResult RiskControl02()
        {
            return View();
        }

        [UserAuth]
        public ActionResult WarningIND()
        {
            return View();
        }

        [UserAuth]
        public ActionResult WatchIND()
        {
            return View();
        }

        [UserAuth]
        public ActionResult Bond_Accounting_EL()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            return View();
        }
    }
}