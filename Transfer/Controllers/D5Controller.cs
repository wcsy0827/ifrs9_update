using System;
using System.Collections.Generic;
using System.Configuration;
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
    public class D5Controller : CommonController
    {
        private ID5Repository D5Repository;
        DateTime startTime = DateTime.MinValue;
        private List<SelectOption> actions = null;

        public D5Controller()
        {
            this.D5Repository = new D5Repository();
            startTime = DateTime.Now;
            //actions = new List<SelectOption>() {
            //    new SelectOption() {Text="查詢",Value="search" },
            //    new SelectOption() {Text="上傳&存檔",Value="upload" }};
        }

    }
}