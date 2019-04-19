using System;
using System.Collections.Generic;
using System.Web.Mvc;
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
    public class TickerRatingController : CommonController
    {
        private ITickerRatingRepository TickerRatingRepository;

        public TickerRatingController()
        {
            this.TickerRatingRepository = new TickerRatingRepository();
        }

        /// <summary>
        /// TickerRating(Ticker & Rating 維護)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult TickerRatingMaintain()
        {
            List<SelectOption> maintainType = new List<SelectOption>() {
                new SelectOption() {Text="擔保者Ticker",Value="Guarantor_Ticker"},
                new SelectOption() {Text="發行者Ticker",Value="Issuer_Ticker"},
                new SelectOption() {Text="發行者信評",Value="Issuer_Rating"},
                new SelectOption() {Text="擔保者信評",Value="Guarantor_Rating"},
                new SelectOption() {Text="債項信評",Value="Bond_Rating"}
            };

            ViewBag.maintainType = new SelectList(maintainType, "Value", "Text");

            return View();
        }

        #region GetGuarantorTickerData
        [BrowserEvent("查詢擔保者Ticker")]
        [HttpPost]
        public JsonResult GetGuarantorTickerData(string queryType, GuarantorTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = TickerRatingRepository.getGuarantorTicker(queryType, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveGuarantorTicker
        [BrowserEvent("儲存擔保者Ticker")]
        [HttpPost]
        public JsonResult SaveGuarantorTicker(string actionType, GuarantorTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = TickerRatingRepository.saveGuarantorTicker(actionType, dataModel);

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

        #region DeleteGuarantorTicker
        [BrowserEvent("刪除擔保者Ticker")]
        [HttpPost]
        public JsonResult DeleteGuarantorTicker(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = TickerRatingRepository.deleteGuarantorTicker(issuer);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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

        #region GetIssuerTickerData
        [BrowserEvent("查詢發行者Ticker")]
        [HttpPost]
        public JsonResult GetIssuerTickerData(string queryType, IssuerTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = TickerRatingRepository.getIssuerTicker(queryType, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveIssuerTicker
        [BrowserEvent("儲存發行者Ticker")]
        [HttpPost]
        public JsonResult SaveIssuerTicker(string actionType, IssuerTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = TickerRatingRepository.saveIssuerTicker(actionType, dataModel);

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

        #region DeleteIssuerTicker
        [BrowserEvent("刪除發行者Ticker")]
        [HttpPost]
        public JsonResult DeleteIssuerTicker(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = TickerRatingRepository.deleteIssuerTicker(issuer);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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

        #region GetIssuerRatingData
        [BrowserEvent("查詢發行者信評")]
        [HttpPost]
        public JsonResult GetIssuerRatingData(string queryType, IssuerRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = TickerRatingRepository.getIssuerRating(queryType, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveIssuerRating
        [BrowserEvent("儲存發行者信評")]
        [HttpPost]
        public JsonResult SaveIssuerRating(string actionType, IssuerRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = TickerRatingRepository.saveIssuerRating(actionType, dataModel);

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

        #region DeleteIssuerRating
        [BrowserEvent("刪除發行者信評")]
        [HttpPost]
        public JsonResult DeleteIssuerRating(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = TickerRatingRepository.deleteIssuerRating(issuer);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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

        #region GetGuarantorRatingData
        [BrowserEvent("查詢擔保者信評")]
        [HttpPost]
        public JsonResult GetGuarantorRatingData(string queryType, GuarantorRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = TickerRatingRepository.getGuarantorRating(queryType, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveGuarantorRating
        [BrowserEvent("儲存擔保者信評")]
        [HttpPost]
        public JsonResult SaveGuarantorRating(string actionType, GuarantorRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = TickerRatingRepository.saveGuarantorRating(actionType, dataModel);

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

        #region DeleteGuarantorRating
        [BrowserEvent("刪除擔保者信評")]
        [HttpPost]
        public JsonResult DeleteGuarantorRating(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = TickerRatingRepository.deleteGuarantorRating(issuer);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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

        #region GetBondRatingData
        [BrowserEvent("查詢債項信評")]
        [HttpPost]
        public JsonResult GetBondRatingData(string queryType, BondRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = TickerRatingRepository.getBondRating(queryType, dataModel);
                result.RETURN_FLAG = queryData.Item1;
                result.Datas = Json(queryData.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SaveBondRating
        [BrowserEvent("儲存債項信評")]
        [HttpPost]
        public JsonResult SaveBondRating(string actionType, BondRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = TickerRatingRepository.saveBondRating(actionType, dataModel);

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

        #region DeleteBondRating
        [BrowserEvent("刪除債項信評")]
        [HttpPost]
        public JsonResult DeleteBondRating(string bondNumber)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = TickerRatingRepository.deleteBondRating(bondNumber);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.delete_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.delete_Fail.GetDescription(resultDelete.DESCRIPTION);
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