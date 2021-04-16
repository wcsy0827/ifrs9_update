using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using System.Linq;
using static Transfer.Enum.Ref;
using Newtonsoft.Json;
using Transfer.Models;
using System.Configuration;
using System.Threading;

namespace Transfer.Controllers
{
    [LogAuth]
    public class KRiskController : CommonController
    {
        private IKriskRepository KriskRepository;
        private ID0Repository D0Repository;
        private List<SelectOption> products = new List<SelectOption>();
        private string kriskUrl = string.Empty;
        private TimeSpan ts = new TimeSpan(0,0,300); //設定連線 KRisk 連線時間

        //public ICacheProvider Cache { get; set; }
        public KRiskController()
        {
            this.KriskRepository = new KriskRepository();
            this.D0Repository = new D0Repository();
            kriskUrl = Properties.Settings.Default["kriskUrl"].ToString();
        }

        [UserAuth]
        public ActionResult KRisk01Detail()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.B), "Value", "Text");
            return View();
        }

        [UserAuth]
        public ActionResult KRisk02Detail()
        {
            ViewBag.product = new SelectList(KriskRepository.getProduct(GroupProductCode.M), "Value", "Text");
            return View();
        }

        public delegate void SaveD04Action(Flow_Apply_Status D04);

        public delegate void SaveD02Action(string reportDate, string result);

        /// <summary>
        /// 執行減損計算(債券)
        /// </summary>
        /// <param name="date">reportDate</param>
        /// <param name="version">version</param>
        /// <param name="product">ex:Bond_A,Bond_B,Bond_C</param> 
        /// <param name="productCode">ex:401,402</param> 
        /// <returns></returns>
        [BrowserEvent("執行減損計算(債券)")]
        [HttpPost]
        public async Task<JsonResult> KriskBondsComplete(string date, string version, string product,string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime dt = DateTime.MinValue;
            int ver = 0;
            if (date.IsNullOrWhiteSpace() || version.IsNullOrWhiteSpace() ||
                product.IsNullOrWhiteSpace() || !DateTime.TryParse(date, out dt) ||
                !int.TryParse(version, out ver))
            {
                return Json(new KRiskFlowLoaderResult() { result = "1", message = Message_Type.parameter_Error.GetDescription() });
            }
            #region 執行前檢核
            result = KriskRepository.CheckInfo(date, version);
            if (!result.RETURN_FLAG)
            {                
                return Json(new KRiskFlowLoaderResult() { result = "1", message =result.DESCRIPTION });
            }
            #endregion
            var productInfo = KriskRepository.getProductInfo(productCode);
            var pc = string.Join(",", product.Split(',').Select(x => string.Format("'{0}'", x)));         
            Flow_Apply_Status D04 = new Flow_Apply_Status();
            D04.Report_Date = dt;
            D04.Run_Time_Start = DateTime.Now;
            D04.PRJID = productInfo.Item1;
            D04.FLOWID = productInfo.Item2;
            D04.Group_Product_Code = productCode;
            D04.Version = ver;
            D04.Flow_Result = "1";
            string message = string.Empty;           
            List<KRiskFlowLoaderParm> parms = new List<KRiskFlowLoaderParm>();
            parms.Add(new KRiskFlowLoaderParm() { key = "v_uni_project_name", value = productInfo.Item1 });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_uni_flow_name", value = productInfo.Item2 });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_product_code", value = pc });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_report_date", value = string.Format("'{0}'", dt.ToString("yyyyMMdd")) });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_version", value = version.Trim() });
            HttpClient client = new HttpClient();
            client.BaseAddress = new System.Uri(kriskUrl);
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            client.Timeout = ts;
            try
            {
                HttpResponseMessage response = await client.PostAsJsonAsync(@"KRiskFlowLoader/KRiskFlowLoaderServlet", parms);
                if (response.IsSuccessStatusCode)
                {
                    KRiskFlowLoaderResult data = JsonConvert.DeserializeObject<KRiskFlowLoaderResult>(await response.Content.ReadAsStringAsync());
                    if (data.result == "0")
                    {
                        D04.Flow_Result = "0";
                        message = "Success!";
                        KriskRepository.getBondC07CheckData(dt, ver);
                    }
                    else
                    {
                        if (data.message.IndexOf("唯一索引鍵值") > -1)
                        {
                            message = Message_Type.kriskError.GetDescription(null, "執行重複資料");
                        }
                        else
                        {
                            message = data.message;
                        }
                    }
                }
                else
                {
                    message = string.Format("Status Code : {0}", response.StatusCode);
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            SaveD04Action act = SaveD04;
            act.Invoke(D04);
            return Json(new KRiskFlowLoaderResult() { result = D04.Flow_Result, message = message });
        }

        /// <summary>
        /// 執行減損計算(房貸)
        /// </summary>
        /// <param name="date">reportDate</param>
        /// <param name="product">product</param>
        /// <param name="productCode">product_code</param>
        /// <returns></returns>
        [BrowserEvent("執行減損計算(房貸)")]
        [HttpPost]
        public async Task<JsonResult> KriskMortgageComplete(string date, string product, string productCode,bool deleteFlag)
        {
            DateTime dt = DateTime.MinValue;
            if (date.IsNullOrWhiteSpace() || product.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt))
            {
                return Json(new KRiskFlowLoaderResult() { result = "1", message = Message_Type.parameter_Error.GetDescription() });
            }
            var productInfo = KriskRepository.getProductInfo(productCode);
            var pc = string.Join(",", product.Split(',').Select(x => string.Format("'{0}'", x)));
            Flow_Apply_Status D04 = new Flow_Apply_Status();
            D04.Report_Date = dt;
            D04.Run_Time_Start = DateTime.Now;
            D04.PRJID = productInfo.Item1;
            D04.FLOWID = productInfo.Item2;
            D04.Group_Product_Code = productCode;
            D04.Version = 1;
            D04.Flow_Result = "1";
            string message = string.Empty;
            List<KRiskFlowLoaderParm> parms = new List<KRiskFlowLoaderParm>();
            parms.Add(new KRiskFlowLoaderParm() { key = "v_uni_project_name", value = productInfo.Item1 });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_uni_flow_name", value = productInfo.Item2 });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_product_code", value = pc });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_report_date", value = string.Format("'{0}'", dt.ToString("yyyyMMdd")) });
            parms.Add(new KRiskFlowLoaderParm() { key = "v_etl_version", value = "1" });
            HttpClient client = new HttpClient();
            client.BaseAddress = new System.Uri(kriskUrl);
            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
            client.Timeout = ts;
            string deleteMessage = string.Empty;
            List<EL_Data_Out> C07s = new List<EL_Data_Out>();
            if (deleteFlag)
            {
                var datas = KriskRepository.deleteC07(D04);
                deleteMessage = datas.Item1;
                C07s = datas.Item2;
            }
            bool backFlag = false;
            if (deleteMessage.IsNullOrWhiteSpace())
            {
                try
                {
                    backFlag = true;
                    HttpResponseMessage response = await client.PostAsJsonAsync(@"KRiskFlowLoader/KRiskFlowLoaderServlet", parms);
                    if (response.IsSuccessStatusCode)
                    {
                        KRiskFlowLoaderResult data = JsonConvert.DeserializeObject<KRiskFlowLoaderResult>(await response.Content.ReadAsStringAsync());
                        if (data.result == "0")
                        {
                            D04.Flow_Result = "0";
                            message = "Success!";
                        }
                        else
                        {
                            if (data.message.IndexOf("唯一索引鍵值") > -1)
                            {
                                message = Message_Type.kriskError.GetDescription(null, "執行重複資料");
                            }
                            else
                            {
                                message = data.message;
                            }
                        }
                    }
                    else
                    {
                        message = string.Format("Status Code : {0}", response.StatusCode);
                    }
                }
                catch (Exception ex)
                {
                    message = ex.exceptionMessage();
                }
            }
            else
            {
                message = deleteMessage;
            }
            if (deleteFlag && backFlag && D04.Flow_Result != "0") //復原刪掉C07
            {
                KriskRepository.backC07(C07s);
            }
            SaveD04Action act = SaveD04;
            SaveD02Action actD02 = SaveD02;
            act.Invoke(D04);
            if(D04.Flow_Result == "0")
                actD02.Invoke(dt.ToString("yyyy/MM/dd"), "0");
            return Json(new KRiskFlowLoaderResult() { result = D04.Flow_Result, message = message });
        }

        public JsonResult KriskCheckC07(string date, string product, string productCode)
        {
            MSGReturnModel result = new MSGReturnModel();
            var productInfo = KriskRepository.getProductInfo(productCode);
            DateTime dt = DateTime.MinValue;
            DateTime.TryParse(date, out dt);
            result.RETURN_FLAG = KriskRepository.checkC07(dt, productInfo.Item1, productInfo.Item2);
            return Json(result);
        }

        [HttpPost]
        public JsonResult GetC01(string date, string product)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (date.IsNullOrWhiteSpace() || 
                product.IsNullOrWhiteSpace() ||
                !DateTime.TryParse(date, out dt))
            {
                return Json(result);
            }
            var data = KriskRepository.GetC01Version(product, dt);
            if (data.Any())
            {
                result.Datas = Json(data);
                result.RETURN_FLAG = true;
            }
            return Json(result);
        }

        public void SaveD04(Flow_Apply_Status D04)
        {
            D04.Run_Time_End = DateTime.Now;
            KriskRepository.saveD04(D04);
        }

        public void SaveD02(string reportDate, string result)
        {
            D0Repository.saveD02AfterKRisk(reportDate);
        }

        public class KRiskFlowLoaderParm
        {
            public string key { get; set; }
            public string value { get; set; }
        }

        public class KRiskFlowLoaderResult
        {
            public string result { get; set; }
            public string message { get; set; }
        }
    }
}