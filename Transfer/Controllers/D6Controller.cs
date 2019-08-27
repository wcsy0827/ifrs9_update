using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Mvc;
using Transfer.Infrastructure;
using Transfer.Models.Interface;
using Transfer.Models.Repository;
using Transfer.Utility;
using Transfer.ViewModels;
using System.Linq;
using static Transfer.Enum.Ref;
using System.IO;

namespace Transfer.Controllers
{
    [LogAuth]
    public class D6Controller : CommonController
    {
        private IA5Repository A5Repository;
        private ID0Repository D0Repository;
        private ID5Repository D5Repository;
        private ID6Repository D6Repository;
        private List<SelectOption> actions = null;
        private string productCode = Assessment_Type.B.GetDescription();
        DateTime startTime = DateTime.MinValue;
        List<SelectOption> flag = EnumUtil.GetValues<Evaluation_Status_Type>()
            .Select(x => new SelectOption()
            {
                Text = x.GetDescription(),
                Value = x.ToString()
            }).ToList();

            ////new List<SelectOption>() {
            ////    new SelectOption() { Text = "全部",Value = "All"},
            ////    new SelectOption() { Text = "已完成評估",Value = "Y"},
            ////    new SelectOption() { Text = "尚未完成評估",Value = "N"}};

        public D6Controller()
        {
            this.A5Repository = new A5Repository();
            this.D0Repository = new D0Repository();
            this.D5Repository = new D5Repository();
            this.D6Repository = new D6Repository();
            startTime = DateTime.Now;
        }

        // GET: D6Controller
        public ActionResult Index()
        {
            return View();
        }

        /// <summary>
        /// D53 (SMF對應表)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D53Detail()
        {
            ViewBag.action = new SelectList(new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="上傳&存檔",Value="upload" }}, "Value", "Text");
            var jqgridInfo = FactoryRegistry.GetInstance(Table_Type.D53).TojqGridData(new int[] { 150, 150, 280, 150, 250 });
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// D60(信評優先選擇參數檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D60()
        {
            var D60Data = D6Repository.getD60All().Item2;

            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(
                             D60Data.Select(x => x.Parm_ID).Distinct()
                             .Select(
                                      x => new SelectOption()
                                      { Text = (x), Value = x }
                                    )
            );
            ViewBag.parmID = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(
                             D60Data.Select(x => x.Rating_Priority).Distinct()
                             .OrderBy(x => x)
                             .Select(
                                      x => new SelectOption()
                                      { Text = (x), Value = x }
                                    )
            );
            ViewBag.ratingPriority = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(
                             D60Data.Select(x => x.Rating_Object).Distinct()
                             .Select(
                                      x => new SelectOption()
                                      { Text = (x), Value = x }
                                    )
            );
            ViewBag.ratingObject = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "1", Value = "1" },
                                              new SelectOption() { Text = "2", Value = "2" },
                                              new SelectOption() { Text = "3", Value = "3" },
                                              new SelectOption() { Text = "4", Value = "4" },
                                              new SelectOption() { Text = "5", Value = "5" },
                                              new SelectOption() { Text = "6", Value = "6" }
                                           };
            ViewBag.Rating_Priority = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "債項", Value = "債項" },
                                              new SelectOption() { Text = "發行人", Value = "發行人" },
                                              new SelectOption() { Text = "保證人", Value = "保證人" },
                                           };
            ViewBag.Rating_Object = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "國內", Value = "國內" },
                                              new SelectOption() { Text = "國外", Value = "國外" }
                                           };
            ViewBag.Rating_Org_Area = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "1:孰高", Value = "1" },
                                              new SelectOption() { Text = "2:孰低", Value = "2" }
                                           };
            ViewBag.Rating_Selection = new SelectList(selectOption, "Value", "Text");

            ProcessStatusList psl = new ProcessStatusList();
            selectOption = psl.statusOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");

            if (getAssessment(productCode, "D60", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(productCode, "D60", SetAssessmentType.Auditor)
                        .Select(x => new SelectOption()
                        { Text = (x.User_Name), Value = x.User_Account }));
            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            return View();
        }

        /// <summary>
        /// D61(減損階段評估參數檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D61()
        {
            List<SelectOption> selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "2", Value = "2" },
                                              new SelectOption() { Text = "3", Value = "3" }
                                           };
            ViewBag.Assessment_Stage = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" }
                                           };
            selectOption.AddRange(new List<SelectOption>() {
                                              new SelectOption() { Text = "量化衡量", Value = "量化衡量" },
                                              new SelectOption() { Text = "質化衡量", Value = "質化衡量" }
                                           });
            ViewBag.Assessment_Kind = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" }
                                           };
            selectOption.AddRange(EnumUtil.GetValues<AssessmentSubKind_Type>()
                        .Select(x => new SelectOption()
                        {
                            Text = x.GetDescription(),
                            Value = x.GetDescription()
                        }));
            ViewBag.Assessment_Sub_Kind = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = ">", Value = ">" },
                                              new SelectOption() { Text = "=", Value = "=" },
                                              new SelectOption() { Text = "<", Value = "<" },
                                              new SelectOption() { Text = ">=", Value = ">=" },
                                              new SelectOption() { Text = "<=", Value = "<=" },
                                           };
            ViewBag.Check_Symbol = new SelectList(selectOption, "Value", "Text");

            ViewBag.IsActive = new SelectList(new List<SelectOption>()
            {
                new SelectOption() { Text = " ", Value = " " },
                new SelectOption() { Text = "Y", Value = "Y" },
                new SelectOption() { Text = "N", Value = "N" }
            }, "Value", "Text");

            ProcessStatusList psl = new ProcessStatusList();
            selectOption = psl.statusOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            if (getAssessment(productCode, "D61", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(
                getAssessment(productCode, "D61", SetAssessmentType.Auditor)
                .Select(x => new SelectOption()
                { Text = (x.User_Name), Value = x.User_Account }));
            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            return View();
        }

        /// <summary>
        /// D68(信用風險低參數檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D68()
        {
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((A5Repository.getA52("SP").Item2)
                        .Select(x => new SelectOption()
                        { Text = x.Rating, Value = x.Rating }));
            ViewBag.Rating_Floor = new SelectList(selectOption, "Value", "Text");

            ProcessStatusList psl = new ProcessStatusList();
            selectOption = psl.statusOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                              new SelectOption() { Text = "", Value = "" },
                                              new SelectOption() { Text = "主權及國營事業債", Value = "主權及國營事業債" },
                                              new SelectOption() { Text = "其他債券", Value = "其他債券" }
                                           };
            ViewBag.Bond_Type = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "Y：是", Value = "Y" },
                                                  new SelectOption() { Text = "N：否", Value = "N" }
                                              };
            ViewBag.Including_Ind = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "1：以上", Value = "1" },
                                                  new SelectOption() { Text = "0：以下", Value = "0" }
                                              };
            ViewBag.Apply_Range = new SelectList(selectOption, "Value", "Text");

            if (getAssessment(productCode, "D68", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(productCode, "D68", SetAssessmentType.Auditor)
                .Select(x => new SelectOption()
                { Text = (x.User_Name), Value = x.User_Account }));
            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            ViewBag.A51Year = new D7Repository().GetA51Year();

            return View();
        }

        /// <summary>
        /// D69(基本要件參數檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D69()
        {
            List<SelectOption> selectOption = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "Y：是", Value = "Y" },
                                                  new SelectOption() { Text = "N：否", Value = "N" }
                                              };

            ViewBag.BasicPass = new SelectList(selectOption, "Value", "Text");
            ViewBag.RatingOriGoodInd = new SelectList(selectOption, "Value", "Text");
            ViewBag.IncludingInd = new SelectList(selectOption, "Value", "Text");
            ViewBag.RatingCurrGoodInd = new SelectList(selectOption, "Value", "Text");
            ViewBag.OriRatingMissingInd = new SelectList(selectOption, "Value", "Text");

            selectOption = new List<SelectOption>() {
                                                  new SelectOption() { Text = "", Value = "" },
                                                  new SelectOption() { Text = "1：以上", Value = "1" },
                                                  new SelectOption() { Text = "0：以下", Value = "0" }
                                              };

            ViewBag.ApplyRange = new SelectList(selectOption, "Value", "Text");

            ProcessStatusList psl = new ProcessStatusList();
            selectOption = psl.statusOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.Status = new SelectList(selectOption, "Value", "Text");

            if (getAssessment(productCode, "D69", SetAssessmentType.Presented)
                .Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name)))
            {
                ViewBag.IsSender = "Y";
            }
            else
            {
                ViewBag.IsSender = "N";
            }

            selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange(getAssessment(productCode, "D69", SetAssessmentType.Auditor)
                .Select(x => new SelectOption()
                { Text = (x.User_Name), Value = x.User_Account }));

            ViewBag.Auditor = new SelectList(selectOption, "Value", "Text");

            ViewBag.UserAccount = AccountController.CurrentUserInfo.Name;

            IsActiveList ial = new IsActiveList();
            selectOption = ial.isActiveOption;
            selectOption.Insert(0, new SelectOption() { Text = "", Value = "" });
            ViewBag.IsActive = new SelectList(selectOption, "Value", "Text");

            return View();
        }

        /// <summary>
        /// D62(必要條件評估檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D62()
        {
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="執行評估",Value="transfer" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            List<SelectOption> YN = new List<SelectOption>() {
                new SelectOption() {Text="",Value="" },
                new SelectOption() {Text="Y：是",Value="Y" },
                new SelectOption() {Text="N：否",Value="N" }
            };
            ViewBag.YN = new SelectList(YN, "Value", "Text");

            return View();
        }

        /// <summary>
        /// D63(量化評估需求資訊檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D63()
        {
            ViewBag.flag = new SelectList(flag, "Value", "Text");
            ViewBag.assessmentSubKinds = new SelectList(getAssessmentSubKind_Type(AssessmentKind_Type.Quantify), "Value", "Text");
            ViewBag.action = new SelectList(actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="載入觀察名單資料",Value="transfer" }}, "Value", "Text");
            var jqgridInfo = new D63ViewModel().TojqGridData(new int[] { 70,170 }, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            jqgridInfo.colModel.hideColModel(new List<string>() { "Quantitative_Pass_Confirm", "Send_to_Auditor" });
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            ViewBag.AssessorFlag = getAssessment(productCode, Table_Type.D63.ToString(),
                SetAssessmentType.Presented).Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name));
            return View();
        }

        /// D64(量化評估複核)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D64Audit()
        {
            ViewBag.Status = new SelectList(new List<SelectOption>() {
                new SelectOption() { Text = Evaluation_Status_Type.All.GetDescription(), Value = Evaluation_Status_Type.All.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.Review.GetDescription(), Value = Evaluation_Status_Type.Review.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.ReviewCompleted.GetDescription(), Value = Evaluation_Status_Type.ReviewCompleted.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.Reject.GetDescription(), Value = Evaluation_Status_Type.Reject.ToString() },
            }, "Value", "Text");
            var jqgridInfo = new D63ViewModel().TojqGridData(new int[] { 70, 170 }, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            jqgridInfo.colModel.hideColModel(new List<string>() { "Pass_Confirm_Flag", "Result_Version_Confirm_Flag" , "Send_to_Auditor" });
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            var _User_Account = AccountController.CurrentUserInfo.Name;
            ViewBag.Account = _User_Account;
            ViewBag.AssessorFlag = getAssessment(productCode, Table_Type.D63.ToString(),
                SetAssessmentType.Auditor).Any(x => x.User_Account.Contains(_User_Account));
            return View();
        }

        /// <summary>
        /// D65Assesment(質化評估需求資訊檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D65Assesment()
        {
            ViewBag.flag = new SelectList(flag, "Value", "Text");
            ViewBag.assessmentSubKinds = new SelectList(getAssessmentSubKind_Type(AssessmentKind_Type.Qualitative), "Value", "Text");
            var jqgridInfo = new D65ViewModel().TojqGridData(new int[] { 100, 170 }, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            jqgridInfo.colModel.hideColModel(new List<string>() { "Quantitative_Pass_Confirm",  "Assessment_Kind", "Send_to_Auditor" });
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            ViewBag.action = new SelectList(actions = new List<SelectOption>() {
                new SelectOption() { Text="查詢",Value="search" },
                new SelectOption() { Text="新增質化評估個案",Value="case"},
                new SelectOption() { Text="篩選需求質化評估資料" ,Value="transfer" }},
                "Value", "Text");
            ViewBag.AssessorFlag = getAssessment(productCode, Table_Type.D65.ToString(),
                SetAssessmentType.Presented).Any(x => x.User_Account.Contains(AccountController.CurrentUserInfo.Name));
            return View();
        }

        /// <summary>
        /// D65Review(質化評估複核)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D65Review()
        {
            ViewBag.Status = new SelectList(new List<SelectOption>() {
                new SelectOption() { Text = Evaluation_Status_Type.All.GetDescription(), Value = Evaluation_Status_Type.All.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.Review.GetDescription(), Value = Evaluation_Status_Type.Review.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.ReviewCompleted.GetDescription(), Value = Evaluation_Status_Type.ReviewCompleted.ToString() },
                new SelectOption() { Text = Evaluation_Status_Type.Reject.GetDescription(), Value = Evaluation_Status_Type.Reject.ToString() },
            }, "Value", "Text");
            var jqgridInfo = new D65ViewModel().TojqGridData(new int[] { 70, 170 }, null, true);
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            jqgridInfo.colModel.hideColModel(new List<string>() { "Pass_Confirm_Flag", "Result_Version_Confirm_Flag",  "Assessment_Kind", "Send_to_Auditor" });
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            var _User_Account = AccountController.CurrentUserInfo.Name;
            ViewBag.Account = _User_Account;
            ViewBag.AssessorFlag = getAssessment(productCode, Table_Type.D65.ToString(),
                SetAssessmentType.Auditor).Any(x => x.User_Account.Contains(_User_Account));
            return View();
        }

        [UserAuth]
        public ActionResult D54Insert()
        {
            return View();
        }

        [UserAuth]
        public ActionResult D54Search()
        {
            ViewBag.GPC = new SelectList(D5Repository.getD54GroupProductCode(), "Value", "Text");
            var jqgridInfo = new D54ViewModel().TojqGridData(new[] { 250, 300, 150, 85, 110, 70, 150, 120, 70, 150, 170, 90, 150, 200, 185, 240, 160, 85, 270, 250, 190, 260, 240, 170, 110 });
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            return View();
        }

        /// <summary>
        /// D67(信評警示記錄檔)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D67()
        {
            List<SelectOption> actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="執行評估",Value="transfer" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");

            List<SelectOption> isComplete = new List<SelectOption>() {
                new SelectOption() {Text="",Value="" },
                new SelectOption() {Text="完成",Value="Y" },
                new SelectOption() {Text="未完成(有空值的資料)",Value="N" }};
            ViewBag.isComplete = new SelectList(isComplete, "Value", "Text");

            List<SelectOption> YN = new List<SelectOption>() {
                new SelectOption() {Text="",Value="" },
                new SelectOption() {Text="Y：是",Value="Y" },
                //new SelectOption() {Text="N：否",Value="N" }
            };
            ViewBag.YN = new SelectList(YN, "Value", "Text");

            return View();
        }

        /// <summary>
        /// 減損作業狀態查詢 
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D6Check()
        {
            ViewBag.Impairment = new SelectList(
                EnumUtil.GetValues<Impairment_Operations_Type>()
                .Select(x => new SelectOption()
                {
                    Text = x.GetDescription(),
                    Value = x.ToString()
                }).ToList(), "Value", "Text");
           
            return View();
        }

        /// <summary>
        /// 寄信通知查詢
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult D6MailLog()
        {
            return View();
        }

        /// <summary>
        /// 減損報表資料刪除作業(複核者使用for報表重新產製)
        /// </summary>
        /// <returns></returns>
        [UserAuth]
        public ActionResult ReEL()
        {
            actions = new List<SelectOption>() {
                new SelectOption() {Text="查詢",Value="search" },
                new SelectOption() {Text="減損報表資料刪除作業",Value="reSet" }};
            ViewBag.action = new SelectList(actions, "Value", "Text");
            var jqgridInfo = new ReELViewModel().TojqGridData(new int[] { 85, 300, 150, 150, 300 });
            ViewBag.jqgridColNames = jqgridInfo.colNames;
            ViewBag.jqgridColModel = jqgridInfo.colModel;
            List<SelectOption> selectOption = new List<SelectOption>();
            selectOption.Add(new SelectOption() { Text = "", Value = "" });
            selectOption.AddRange((D0Repository.getGroupProductByDebtType(GroupProductCode.B.GetDescription()).Item2)
                .Select(x => new SelectOption()
                { Text = (x.Group_Product_Code + " " + x.Group_Product_Name), Value = x.Group_Product_Code }));

            ViewBag.GroupProduct = new SelectList(selectOption, "Value", "Text");
            return View();
        }

        #region GetD60PRR
        [HttpPost]
        public JsonResult GetD60PRR(string selectName)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD60All();
                result.RETURN_FLAG = queryResult.Item1;

                List<SelectOption> selectOption = new List<SelectOption>();
                selectOption.Add(new SelectOption() { Text = "", Value = "" });

                switch (selectName)
                {
                    case "Parm_ID":
                        selectOption.AddRange(
                                         queryResult.Item2.Select(x => x.Parm_ID).Distinct()
                                         .Select(
                                                  x => new SelectOption()
                                                  { Text = (x), Value = x }
                                                )
                        );
                        break;
                    case "Rating_Priority":
                        selectOption.AddRange(
                                         queryResult.Item2.Select(x => x.Rating_Priority).Distinct()
                                         .OrderBy(x => x)
                                         .Select(
                                                  x => new SelectOption()
                                                  { Text = (x), Value = x }
                                                )
                        );
                        break;
                    case "Rating_Object":
                        selectOption.AddRange(
                                         queryResult.Item2.Select(x => x.Rating_Object).Distinct()
                                         .Select(
                                                  x => new SelectOption()
                                                  { Text = (x), Value = x }
                                                )
                        );
                        break;
                    default:
                        break;
                }

                result.Datas = Json(selectOption);

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

        #region GetD60All
        [HttpPost]
        public JsonResult GetD60AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD60All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D60DbfileData); //清除
                Cache.Set(CacheList.D60DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheD60Data
        [HttpPost]
        public JsonResult GetCacheD60Data(jqGridParam jdata)
        {
            List<D60ViewModel> data = new List<D60ViewModel>();
            if (Cache.IsSet(CacheList.D60DbfileData))
            {
                data = (List<D60ViewModel>)Cache.Get(CacheList.D60DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetD60Data
        [BrowserEvent("查詢D60資料")]
        [HttpPost]
        public JsonResult GetD60Data(D60ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD60(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D60DbfileData); //清除
                Cache.Set(CacheList.D60DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveD60
        [BrowserEvent("儲存D60資料")]
        [HttpPost]
        public JsonResult SaveD60(string actionType, D60ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = D6Repository.saveD60(actionType, dataModel);

            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region DeleteD60
        [BrowserEvent("D60資料設為失效")]
        [HttpPost]
        public JsonResult DeleteD60(string parmID)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = D6Repository.deleteD60(parmID);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SendD60ToAudit
        [BrowserEvent("D60呈送複核")]
        [HttpPost]
        public JsonResult SendD60ToAudit(string parmID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {

                result = D6Repository.sendD60ToAudit(parmID, auditor);

            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region D60Audit
        [BrowserEvent("D60 複核確認/退回")]
        [HttpPost]
        public JsonResult D60Audit(string parmID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = D6Repository.D60Audit(parmID, status);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetD61All
        [HttpPost]
        public JsonResult GetD61AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD61All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D61DbfileData); //清除
                Cache.Set(CacheList.D61DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheD61Data
        [HttpPost]
        public JsonResult GetCacheD61Data(jqGridParam jdata)
        {
            List<D61ViewModel> data = new List<D61ViewModel>();
            if (Cache.IsSet(CacheList.D61DbfileData))
            {
                data = (List<D61ViewModel>)Cache.Get(CacheList.D61DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetD61Data
        [BrowserEvent("查詢D61資料")]
        [HttpPost]
        public JsonResult GetD61Data(D61ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD61(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D61DbfileData); //清除
                Cache.Set(CacheList.D61DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveD61
        [BrowserEvent("儲存D61資料")]
        [HttpPost]
        public JsonResult SaveD61(string actionType, D61ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = D6Repository.saveD61(actionType, dataModel);

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

        #region DeleteD61
        [BrowserEvent("D61刪除資料")]
        [HttpPost]
        public JsonResult DeleteD61(string checkItemCode,string Id)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = D6Repository.deleteD61(checkItemCode, Id);

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

        #region SendD61ToAudit
        [BrowserEvent("D61呈送複核")]
        [HttpPost]
        public JsonResult SendD61ToAudit(string checkItemCode, string auditor, string Ids)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                List<D61ViewModel> dataList = new List<D61ViewModel>();
                string[] arrayId = Ids.Split(',');
                for (int i = 0; i < arrayId.Length; i++)
                {
                    if (!arrayId[i].IsNullOrWhiteSpace())
                    {
                        D61ViewModel dataModel = new D61ViewModel();

                        dataModel.Id = arrayId[i];
                        dataModel.Auditor = auditor;

                        dataList.Add(dataModel);
                    }
                }
                MSGReturnModel resultSendToAudit = D6Repository.sendD61ToAudit(dataList);

                result.RETURN_FLAG = resultSendToAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription(resultSendToAudit.DESCRIPTION);
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

        #region D61Audit
        [BrowserEvent("D61複核確認/退回")]
        [HttpPost]
        public JsonResult D61Audit(string checkItemCode, string status ,string Ids)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                List<D61ViewModel> dataList = new List<D61ViewModel>();
                string[] arrayId = Ids.Split(',');
                for (int i = 0; i < arrayId.Length; i++)
                {
                    if (!arrayId[i].IsNullOrWhiteSpace())
                    {
                        D61ViewModel dataModel = new D61ViewModel();

                        dataModel.Id = arrayId[i];
                        dataModel.Status = status;

                        dataList.Add(dataModel);
                    }
                }

                MSGReturnModel resultAudit = D6Repository.D61Audit(dataList);

                result.RETURN_FLAG = resultAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.Audit_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.Audit_Fail.GetDescription(resultAudit.DESCRIPTION);
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

        #region GetD68All
        [HttpPost]
        public JsonResult GetD68AllData()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD68All();

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D68DbfileData); //清除
                Cache.Set(CacheList.D68DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region GetCacheD68Data
        [HttpPost]
        public JsonResult GetCacheD68Data(jqGridParam jdata)
        {
            List<D68ViewModel> data = new List<D68ViewModel>();
            if (Cache.IsSet(CacheList.D68DbfileData))
            {
                data = (List<D68ViewModel>)Cache.Get(CacheList.D68DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data));
        }
        #endregion

        #region GetD68Data
        [BrowserEvent("查詢D68資料")]
        [HttpPost]
        public JsonResult GetD68Data(D68ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryResult = D6Repository.getD68(dataModel);

                result.RETURN_FLAG = queryResult.Item1;

                Cache.Invalidate(CacheList.D68DbfileData); //清除
                Cache.Set(CacheList.D68DbfileData, queryResult.Item2); //把資料存到 Cache

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

        #region SaveD68
        [BrowserEvent("儲存D68資料")]
        [HttpPost]
        public JsonResult SaveD68(string actionType, D68ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = D6Repository.saveD68(actionType, dataModel);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region DeleteD68
        [BrowserEvent("D68設為失效")]
        [HttpPost]
        public JsonResult DeleteD68(string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultDelete = D6Repository.deleteD68(ruleID);

                result.RETURN_FLAG = resultDelete.RETURN_FLAG;
                result.DESCRIPTION = resultDelete.DESCRIPTION;
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region SendD68ToAudit
        [BrowserEvent("D68呈送複核")]
        [HttpPost]
        public JsonResult SendD68ToAudit(string ruleID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = D6Repository.sendD68ToAudit(ruleID, auditor);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region D68Audit
        [BrowserEvent("D68 複核確認/退回")]
        [HttpPost]
        public JsonResult D68Audit(string ruleID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                result = D6Repository.D68Audit(ruleID, status);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }
        #endregion

        #region GetD69AllData

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD69AllData(D69ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = "";

            try
            {
                var D69Data = D6Repository.getD69All();
                result.RETURN_FLAG = D69Data.Item1;
                result.Datas = Json(D69Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion GetD69AllData

        #region GetD69Data

        /// <summary>
        /// 前端抓資料時呼叫
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [BrowserEvent("查詢D69資料")]
        [HttpPost]
        public JsonResult GetD69Data(D69ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D69");

            try
            {
                var D69Data = D6Repository.getD69(dataModel);
                result.RETURN_FLAG = D69Data.Item1;
                result.Datas = Json(D69Data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion GetD69Data

        #region SaveD69
        /// <summary>
        /// 新增、俢改
        /// </summary>
        /// <param name="dataModel"></param>
        /// <returns></returns>
        [BrowserEvent("儲存D69資料")]
        [HttpPost]
        public JsonResult SaveD69(D69ViewModel data)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D69", result.DESCRIPTION);

            try
            {
                MSGReturnModel resultSave = D6Repository.saveD69(data);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription("D69");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription("D69", resultSave.DESCRIPTION);
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

        #region DeleteD69
        [BrowserEvent("D69資料設為失效")]
        [HttpPost]
        public JsonResult DeleteD69(string ruleID)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D69", result.DESCRIPTION);

            try
            {
                result = D6Repository.deleteD69(ruleID);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion DeleteD69

        #region SendD69ToAudit

        /// <summary>
        /// 呈送複核
        [BrowserEvent("呈送複核D69")]
        [HttpPost]
        public JsonResult SendD69ToAudit(string ruleID, string auditor)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D69", result.DESCRIPTION);

            try
            {
                List<D69ViewModel> dataList = new List<D69ViewModel>();

                string[] arrayRuleID = ruleID.Split(',');
                for (int i = 0; i < arrayRuleID.Length; i++)
                {
                    D69ViewModel dataModel = new D69ViewModel();
                    dataModel.Rule_ID = arrayRuleID[i];
                    dataModel.Auditor = auditor;

                    dataList.Add(dataModel);
                }

                MSGReturnModel resultSendToAudit = D6Repository.sendD69ToAudit(dataList);

                result.RETURN_FLAG = resultSendToAudit.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.send_To_Audit_Success.GetDescription("D69");

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.send_To_Audit_Fail.GetDescription("D69", resultSendToAudit.DESCRIPTION);
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

        #region D69Audit
        [BrowserEvent("D69複核確認")]
        [HttpPost]
        public JsonResult D69Audit(string ruleID, string status)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D69", result.DESCRIPTION);

            try
            {
                List<D69ViewModel> dataList = new List<D69ViewModel>();

                string[] arrayRuleID = ruleID.Split(',');
                for (int i = 0; i < arrayRuleID.Length; i++)
                {
                    D69ViewModel dataModel = new D69ViewModel();
                    dataModel.Rule_ID = arrayRuleID[i];
                    dataModel.Status = status;

                    dataList.Add(dataModel);
                }
                result = D6Repository.D69Audit(dataList);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }

            return Json(result);
        }

        #endregion

        #region GetD62Data
        [BrowserEvent("查詢D62資料")]
        [HttpPost]
        public JsonResult GetD62Data(string reportDateStart, string reportDateEnd,
                                     string referenceNbr, string bondNumber,
                                     string basicPass,
                                     string watchIND, string warningIND,
                                     string chgInSpreadIND, string beforeHasChgInSpread)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription("D62");

            try
            {
                var queryData = D6Repository.getD62(reportDateStart, reportDateEnd,
                                                    referenceNbr, bondNumber,
                                                    basicPass,
                                                    watchIND, warningIND,
                                                    chgInSpreadIND, beforeHasChgInSpread);
                result.RETURN_FLAG = (queryData.Item1 == "查無資料" ? false : true);
                Cache.Invalidate(CacheList.D62DbfileData); //清除
                Cache.Set(CacheList.D62DbfileData, queryData.Item2); //把資料存到 Cache
                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
                else
                {
                    result.DESCRIPTION = queryData.Item1;
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

        #region GetCacheD62Data
        [HttpPost]
        public JsonResult GetCacheD62Data(jqGridParam jdata)
        {
            List<D62ViewModel> data = new List<D62ViewModel>();
            if (Cache.IsSet(CacheList.D62DbfileData))
            {
                data = (List<D62ViewModel>)Cache.Get(CacheList.D62DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true, new List<string>() { "Reference_Nbr", "Lots", "Version", "Map_Rule_Id_D70", "Map_Rule_Id_D71" }));
        }
        #endregion

        #region 匯出 Excel

        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("D62匯出Excel")]
        [HttpPost]
        public ActionResult GetD62Excel(string reportDateStart, string reportDateEnd,
                                        string referenceNbr, string bondNumber,
                                        string basicPass,
                                        string watchIND, string warningIND,
                                        string chgInSpreadIND, string beforeHasChgInSpread)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                var data = D6Repository.getD62(reportDateStart, reportDateEnd,
                                               referenceNbr, bondNumber,
                                               basicPass,
                                               watchIND, warningIND,
                                               chgInSpreadIND, beforeHasChgInSpread);

                path = "D62".GetExelName();
                result = D6Repository.DownLoadD62Excel(ExcelLocation(path), data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region SaveD62
        [BrowserEvent("D62執行評估")]
        [HttpPost]
        public JsonResult SaveD62(string reportDate, string version)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription("D62", result.DESCRIPTION);

            try
            {
                MSGReturnModel resultSave = D6Repository.saveD62(reportDate, version);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription("D62", resultSave.DESCRIPTION);
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

        #region getD62ByChg_In_Spread
        [HttpPost]
        public JsonResult getD62ByChg_In_Spread(D62ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = D6Repository.getD62ByChg_In_Spread(dataModel);

                result.RETURN_FLAG = (queryData.Item1 != "" ? false : true);

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = queryData.Item1;
                }
                else
                {
                    result.Datas = Json(queryData.Item2);
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

        #region modifyD62
        [BrowserEvent("修改D62資料")]
        [HttpPost]
        public JsonResult modifyD62(D62ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                dataModel.Watch_IND_Override = dataModel.Watch_IND;
                dataModel.Watch_IND_Override_Date = DateTime.Now.ToString("yyyy/MM/dd");
                dataModel.Watch_IND_Override_User = AccountController.CurrentUserInfo.Name;

                dataModel.Warning_IND_Override = dataModel.Warning_IND;
                dataModel.Warning_IND_Override_Date = DateTime.Now.ToString("yyyy/MM/dd");
                dataModel.Warning_IND_Override_User = AccountController.CurrentUserInfo.Name;

                MSGReturnModel modifyResult = D6Repository.modifyD62(dataModel);

                result.RETURN_FLAG = modifyResult.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription();

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(modifyResult.DESCRIPTION);
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

        #region Get Cache Data
        /// <summary>
        /// Get Cache Data
        /// </summary>
        /// <param name="jdata"></param>
        /// <param name="table">table name</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCacheData(jqGridParam jdata, string table)
        {
            List<D63ViewModel> D63Data = new List<D63ViewModel>();
            switch (table)
            {
                case "D63Transfer":
                    if (Cache.IsSet(CacheList.D63DbfileTransferData))
                        D63Data = (List<D63ViewModel>)Cache.Get(CacheList.D63DbfileTransferData);  //從Cache 抓資料
                    break;
                case "D63":
                    if (Cache.IsSet(CacheList.D63DbfileData))
                        D63Data = (List<D63ViewModel>)Cache.Get(CacheList.D63DbfileData);  //從Cache 抓資料
                    break;
                case "D63History":
                    if (Cache.IsSet(CacheList.D63DbfileHistoryData))
                        D63Data = (List<D63ViewModel>)Cache.Get(CacheList.D63DbfileHistoryData);
                    break;
                case "D64":
                    if (Cache.IsSet(CacheList.D64DbfileData))
                    {
                        var data = (List<D64ViewModel>)Cache.Get(CacheList.D64DbfileData);  //從Cache 抓資料
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D65Transfer":
                    if (Cache.IsSet(CacheList.D65DbfileTransferData))
                    {
                        var data = (List<D65ViewModel>)Cache.Get(CacheList.D65DbfileTransferData);  //從Cache 抓資料
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D65":
                    if (Cache.IsSet(CacheList.D65DbfileData))
                    {
                        var data = (List<D65ViewModel>)Cache.Get(CacheList.D65DbfileData);  //從Cache 抓資料
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D65History":
                    if (Cache.IsSet(CacheList.D65DbfileHistoryData))
                    {
                        var data = (List<D65ViewModel>)Cache.Get(CacheList.D65DbfileHistoryData);
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D66":
                    if (Cache.IsSet(CacheList.D66DbfileData))
                    {
                        var data = (List<D66ViewModel>)Cache.Get(CacheList.D66DbfileData);
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D54InsertSearchData":
                    if (Cache.IsSet(CacheList.D54InsertSearchData))
                    {
                        var data = (List<D54ViewModel>)Cache.Get(CacheList.D54InsertSearchData);
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "D54DbfileData":
                    if (Cache.IsSet(CacheList.D54DbfileData))
                    {
                        var data = (List<D54ViewModel>)Cache.Get(CacheList.D54DbfileData);
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
                case "ReEL":
                    if (Cache.IsSet(CacheList.ReELDbfileData))
                    {
                        var data = (List<ReELViewModel>)Cache.Get(CacheList.ReELDbfileData);
                        return Json(jdata.modelToJqgridResult(data));
                    }
                    break;
            }
            return Json(jdata.modelToJqgridResult(D63Data));
        }
        #endregion       


        #region 匯出 Excel

        /// <summary>
        /// 匯出 Excel
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("D64匯出Excel")]
        [HttpPost]
        public ActionResult GetD64Excel(string reportDate, string referenceNbr, string bondNumber, string isComplete)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                //var data = D6Repository.getD64(reportDate, referenceNbr, bondNumber, isComplete);

                //path = "D64".GetExelName();
                //result = D6Repository.DownLoadD64Excel(ExcelLocation(path), data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region getAssessmentStage
        /// <summary>
        /// get Assessment_Stage
        /// </summary>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        public JsonResult getAssessmentStage(
            string Assessment_Kind,
            string Assessment_Sub_Kind
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = D6Repository.getAssessmentStage(Assessment_Kind, Assessment_Sub_Kind);
            result.RETURN_FLAG = data.Any();
            if (result.RETURN_FLAG)
                result.Datas = Json(data);
            return Json(result);
        }
        #endregion

        #region get Check_Item_Code
        /// <summary>
        /// get Check_Item_Code
        /// </summary>
        /// <param name="Assessment_Stage"></param>
        /// <param name="Assessment_Kind"></param>
        /// <param name="Assessment_Sub_Kind"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetCheckItemCode(
            string Assessment_Stage,
            string Assessment_Kind,
            string Assessment_Sub_Kind)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var data = D6Repository.getCheckItemCode(Assessment_Stage, Assessment_Kind, Assessment_Sub_Kind);
            if (data.Any())
            {
                result.RETURN_FLAG = true;
                result.Datas = Json(data);
            }
            return Json(result);
        }
        #endregion

        #region QualitativeFile
        [BrowserEvent("上傳評估資料")]
        [HttpPost]
        public JsonResult UpLoadQualitativeFile()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 前端無傳送檔案進來

                if (!Request.Files.AllKeys.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                    return Json(result);
                }

                #endregion 前端無傳送檔案進來

                var FileModel = Request.Files["UploadedFile"];
                string Check_Reference = Request.Form["Check_Reference"];
                string sameFlag = Request.Form["sameFlag"];
                string type = Request.Form["type"];

                #region 前端檔案驗證失敗
                string fileName = Path.GetFileName(FileModel.FileName);
                string pathType = Path.GetExtension(FileModel.FileName)
                       .Substring(1); //上傳的檔案類型

                ////ModelState
                //List<string> pathTypes = new List<string>()
                //{
                //    "xlsx","xls","docx","pdf"
                //};
                //if (FileModel.ContentLength == 0 || !pathTypes.Contains(pathType.ToLower()))
                //{
                //    result.RETURN_FLAG = false;
                //    result.DESCRIPTION = "無資料或檔案類型不為Excel,Word,PDF !";
                //    return Json(result);
                //}

                #endregion 前端檔案驗證失敗(驗證)

                #region 上傳檔案


                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                Table_Type _type = Table_Type.D64;
                string projectFile = string.Empty;
                if (type == "D64")
                {
                    _type = Table_Type.D64;
                    projectFile = Server.MapPath("~/" + SetFile.QuantifyFile); //D64專案資料夾
                }
                else if (type == "D66")
                {
                    _type = Table_Type.D66;
                    projectFile = Server.MapPath("~/" + SetFile.QualitativeFile); //D66專案資料夾
                }
                
                FileRelated.createFile(projectFile); //檢查是否有資料夾,如果沒有就新增

                string projectFileSub = Path.Combine(projectFile, Check_Reference);
                FileRelated.createFile(projectFileSub); //檢查是否有Check_Reference資料夾,如果沒有就新增

                string path = Path.Combine(projectFileSub, fileName);


                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel);
                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                if (sameFlag != "Y") //Y(有重複檔名) or N(需加入資料表)  
                    result = D6Repository.SaveQualitativeFile(Check_Reference, fileName,_type, path);

                if (result.RETURN_FLAG)
                {
                    if (type == "D64")
                        result.Datas = Json(D6Repository.getQuantifyFile(Check_Reference));
                    if (type == "D66")
                        result.Datas = Json(D6Repository.getQualitativeFile(Check_Reference));
                }
                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }
            return Json(result);
        }

        [HttpPost]
        public JsonResult GetQuantifyFile(string Check_Reference)
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = D6Repository.getQuantifyFile(Check_Reference);
            result.RETURN_FLAG = data.Any();
            if (!result.RETURN_FLAG)
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            else
                result.Datas = Json(data);
            return Json(result);
        }

        [HttpPost]
        public JsonResult GetQualitativeFile(string Check_Reference)
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = D6Repository.getQualitativeFile(Check_Reference);
            result.RETURN_FLAG = data.Any();
            if (!result.RETURN_FLAG)
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            else
                result.Datas = Json(data);
            return Json(result);
        }

        [HttpGet]
        public ActionResult DLQuantifyFile(string Check_Reference, string fileName)
        {
            try
            {
                string path = Path.Combine(Path.Combine(Server.MapPath("~/" + SetFile.QuantifyFile), Check_Reference), fileName);
                //return File(path, "application/octet-stream", fileName);        
                if (System.IO.File.Exists(path))
                    return File(path, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
                else
                    return Content("檔案已遺失!");
            }
            catch (Exception ex)
            {
                return Content(ex.exceptionMessage());
            }
        }

        [HttpGet]
        public ActionResult DLQualitativeFile(string Check_Reference, string fileName)
        {
            try
            {
                string path = Path.Combine(Path.Combine(Server.MapPath("~/" + SetFile.QualitativeFile), Check_Reference), fileName);
                //return File(path, "application/octet-stream", fileName);        
                if (System.IO.File.Exists(path))
                    return File(path, System.Net.Mime.MediaTypeNames.Application.Octet, fileName);
                else
                    return Content("檔案已遺失!");
            }
            catch (Exception ex)
            {
                return Content(ex.exceptionMessage());
            }
        }

        #region DelFile
        /// <summary>
        /// 刪除D64orD66附件檔案
        /// </summary>
        /// <param name="type">D64 or D66</param>
        /// <param name="Check_Reference">Check_Reference</param>
        /// <param name="fileName">File_Name</param>
        /// <returns></returns>
        [BrowserEvent("刪除D64orD66附件檔案")]
        [HttpPost]
        public JsonResult DelFile(
            string type,
            string Check_Reference,
            string fileName
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result = D6Repository.DelQuantifyAndQualitativeFile(type, Check_Reference, fileName);
            if (result.RETURN_FLAG)
            {
                if (type == "D64")
                    result.Datas = Json(D6Repository.getQuantifyFile(Check_Reference));
                if (type == "D66")
                    result.Datas = Json(D6Repository.getQualitativeFile(Check_Reference));
            }
            return Json(result);
        }
        #endregion

        #endregion

        #region get 評估結果說明 或 備註
        /// <summary>
        /// get 評估結果說明 或 備註
        /// </summary>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Vresion"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="Check_Item_Code"></param>
        /// <param name="type"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetTextArea(
            string ReportDate,
            string Reference_Nbr,
            string Vresion,
            string Assessment_Result_Version,
            string Check_Item_Code,
            string type,
            string table)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime dt = DateTime.MinValue;
            int v = 0;
            int arv = 0;
            if (!DateTime.TryParse(ReportDate, out dt) ||
                !Int32.TryParse(Vresion, out v) ||
                !Int32.TryParse(Assessment_Result_Version, out arv))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                Table_Type _table = EnumUtil.GetValues<Table_Type>()
                    .FirstOrDefault(x => x.ToString() == table);
                result.Datas = Json(D6Repository.GetTextArea(dt, Reference_Nbr, v, arv, Check_Item_Code, type, _table));
                result.RETURN_FLAG = true;
            }
            return Json(result);
        }
        #endregion

        #region save 評估結果說明 或 備註
        /// <summary>
        /// save 評估結果說明 或 備註
        /// </summary>
        /// <param name="value"></param>
        /// <param name="ReportDate"></param>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Vresion"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="Check_Item_Code"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult SaveTextArea(
            string value,
            string ReportDate,
            string Reference_Nbr,
            string Vresion,
            string Assessment_Result_Version,
            string Check_Item_Code,
            string type,
            string table)
        {
            MSGReturnModel result = new MSGReturnModel();
            DateTime dt = DateTime.MinValue;
            int v = 0;
            int arv = 0;
            if (!DateTime.TryParse(ReportDate, out dt) ||
                !Int32.TryParse(Vresion, out v) ||
                !Int32.TryParse(Assessment_Result_Version, out arv))
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                var _table = EnumUtil.GetValues<Table_Type>()
                    .FirstOrDefault(x => x.ToString() == table);
                result = D6Repository.SaveTextArea(value, dt, Reference_Nbr, v, arv, Check_Item_Code, type, _table);
            }
            return Json(result);
        }
        #endregion

        #region get 減損試算結果
        /// <summary>
        /// get 減損試算結果
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetChg_In_Spread(string Reference_Nbr)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var data = D6Repository.getChgInSpread(Reference_Nbr);
            if (data.Any())
            {
                result.RETURN_FLAG = true;
                result.Datas = Json(data);
            }
            return Json(result);
        }
        #endregion

        #region 查詢預計調整資料
        /// <summary>
        /// 查詢預計調整資料
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        [BrowserEvent("查詢預計調整資料")]
        [HttpPost]
        public JsonResult D54InsertSearch(string dt)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = true;
            List<D54ViewModel> D54s = D5Repository.getD54InsertSearch(dt);
            List<D54Group> D54g = new List<D54Group>();
            Cache.Invalidate(CacheList.D54InsertSearchData); //清除

            if (D54s.Any())
            {
                if (D54s.First().Bond_Number == null)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = $@"以下帳戶編號:{string.Join(",", D54s.Select(x => x.Reference_Nbr))}，落入觀察名單或於質化評估手動新增，但未經過D63量化評估覆核通過或D65質化評估覆核通過，最後減損階段未確認";//PGE需求，修改提示訊息
                    return Json(result);
                }
                else
                {
                    Cache.Set(CacheList.D54InsertSearchData, D54s); //把資料存到 Cache
                    var P2 = D54s.Where(x => x.Impairment_Stage == "2");
                    var P3 = D54s.Where(x => x.Impairment_Stage == "3");
                    D54g.Add(new D54Group()
                    {
                        Product_Code = "調整為減損階段2",
                        Bond_A = P2.Count(x => x.Product_Code == "Bond_A").ToString(),
                        Bond_B = P2.Count(x => x.Product_Code == "Bond_B").ToString(),
                        Bond_P = P2.Count(x => x.Product_Code == "Bond_P").ToString()
                    });
                    D54g.Add(new D54Group()
                    {
                        Product_Code = "調整為減損階段3",
                        Bond_A = P3.Count(x => x.Product_Code == "Bond_A").ToString(),
                        Bond_B = P3.Count(x => x.Product_Code == "Bond_B").ToString(),
                        Bond_P = P3.Count(x => x.Product_Code == "Bond_P").ToString()
                    });
                    result.Datas = Json(D54g);
                }
            }
            else
            {
                D54g.Add(new D54Group()
                {
                    Product_Code = "調整為減損階段2",
                    Bond_A = "0",
                    Bond_B = "0",
                    Bond_P = "0"
                });
                D54g.Add(new D54Group()
                {
                    Product_Code = "調整為減損階段3",
                    Bond_A = "0",
                    Bond_B = "0",
                    Bond_P = "0"
                });
                result.Datas = Json(D54g);
            }
            return Json(result);
        }
        #endregion
        #region PGE需求延伸，調整查詢預計調整資料時的檢核
        [HttpPost]
        public JsonResult TransferD54Check(string dt)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D5Repository.checkD54(dt);
            return Json(result);
        }
        [HttpPost]
        public JsonResult CheckC10(string dt)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D5Repository.CheckC10data(dt);
            return Json(result);
        }
        #endregion      

        #region 減損階段確認
        /// <summary>
        /// 減損階段確認
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        [HttpPost]
        [BrowserEvent("最終減損階段確認")]
        public JsonResult TransferD54(string dt)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D5Repository.SaveD54(dt);
            return Json(result);

        }
        #endregion

        #region getD54
        /// <summary>
        /// 最終減損金額查詢
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("最終減損金額查詢")]
        [HttpPost]
        public JsonResult GetD54(string reportdate, string groupProductCode, string bondNumber, string refNumber)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (!DateTime.TryParse(reportdate, out dt) || groupProductCode.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
            }
            else
            {
                var data = D5Repository.getD54(dt, groupProductCode, bondNumber, refNumber);
                if (data.Any())
                {
                    Cache.Invalidate(CacheList.D54DbfileData); //清除
                    Cache.Set(CacheList.D54DbfileData, data); //把資料存到 Cache
                    result.RETURN_FLAG = true;
                }
                else
                {
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                }
            }
            return Json(result);
        }
        #endregion

        #region getD54Excel
        [BrowserEvent("下載D54Excel檔案")]
        [HttpPost]
        public JsonResult GetD54Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            if (Cache.IsSet(CacheList.D54DbfileData))
            {
                var D54 = Excel_DownloadName.D54.ToString();
                var D54Data = (List<D54ViewModel>)Cache.Get(CacheList.D54DbfileData);  //從Cache 抓資料
                result = D6Repository.DownLoadExcel(D54, ExcelLocation(D54.GetExelName()), D54Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            return Json(result);
        }
        #endregion

        #region 複核相關
        #region D63 相關

        #region SearchD63
        /// <summary>
        /// 查詢D63(量化評估需求資訊檔)資料
        /// </summary>
        /// <param name="date">報導日</param>
        /// <param name="watch">觀察名單</param>
        /// <param name="assessmentSubKind">評估次分類</param>
        /// <param name="bondNumber">債券編號</param>
        /// <param name="index">指標數值</param>
        /// <returns></returns>
        [BrowserEvent("查詢D63量化評估結果")]
        [HttpPost]
        public JsonResult SearchD63(
            string date,
            string assessmentSubKind,
            string bondNumber,
            string index,
            bool Send_to_AuditorFlag = false,
            string referenceNbr = "")
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (!(assessmentSubKind == "All" ||
                EnumUtil.GetValues<AssessmentSubKind_Type>()
                .Any(x => x.ToString() == assessmentSubKind)) ||
                !DateTime.TryParse(date, out dt))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            string type = "All";
            if (assessmentSubKind != "All")
                type = EnumUtil.GetValues<AssessmentSubKind_Type>()
                .First(x => x.ToString() == assessmentSubKind)
                .GetDescription();
            Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                .FirstOrDefault(x => x.ToString() == index);
            var datas = D6Repository.getD63(dt, bondNumber, _index, type, Send_to_AuditorFlag, referenceNbr);
            if (datas.Any())
            {
                Cache.Invalidate(CacheList.D63SearchCache);//清除
                Cache.Set(CacheList.D63SearchCache, new D6SearchCache() {
                    dt = dt,
                    bondNumber = bondNumber,
                    EST = _index,
                    type = type,
                    Send_to_AuditorFlag = Send_to_AuditorFlag,
                    referenceNbr = referenceNbr
                });//把資料存到 Cache

                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D63DbfileData); //清除
                Cache.Set(CacheList.D63DbfileData, datas); //把資料存到 Cache
                return Json(result);
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Data.GetDescription();
                return Json(result);
            }
        }
        #endregion

        #region getD63 History
        /// <summary>
        /// 查詢D63
        /// </summary>
        /// <param name="referenceNbr"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD63History(
            string referenceNbr,
            string version
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            int ver = 0;
            Int32.TryParse(version, out ver);
            var datas = D6Repository.getD63History(referenceNbr);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D63DbfileHistoryData); //清除
                Cache.Set(CacheList.D63DbfileHistoryData, datas); //把資料存到 Cache
                return Json(result);
            }
            result.DESCRIPTION = "尚未執行評估";
            return Json(result);
        }
        #endregion

        #region getD64
        /// <summary>
        /// Get D64
        /// </summary>
        /// <param name="referenceNbr">ReferenceNbr</param>
        /// <param name="version">Assessment_Result_Version</param>
        /// <returns></returns>
        [HttpPost]
        public JsonResult GetD64(
            string referenceNbr,
            string version
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Data.GetDescription();
            int ver = 0;
            Int32.TryParse(version, out ver);
            var datas = D6Repository.getD64(referenceNbr, ver);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D64DbfileData); //清除
                Cache.Set(CacheList.D64DbfileData, datas); //把資料存到 Cache

                var D63SearchCache = (D6SearchCache)Cache.Get(CacheList.D63SearchCache);
                var _datas = new List<D63ViewModel>();
                if (D63SearchCache != null)
                {
                     _datas = D6Repository.getD63(
                        D63SearchCache.dt,
                        D63SearchCache.bondNumber,
                        D63SearchCache.EST,
                        D63SearchCache.type,
                        D63SearchCache.Send_to_AuditorFlag,
                        D63SearchCache.referenceNbr);
                    if (datas.Any())
                    {
                        Cache.Invalidate(CacheList.D63DbfileData); //清除
                        Cache.Set(CacheList.D63DbfileData, _datas); //把資料存到 Cache
                    }
                }
            }
            return Json(result);
        }
        #endregion

        #region SaveD63
        /// <summary>
        /// 推送D63評估資料
        /// </summary>
        /// <param name="action"></param>
        /// <param name="Auditor"></param>
        /// <param name="model"></param>
        /// <param name="assessmentSubKind"></param>
        /// <param name="bondNumber"></param>
        /// <param name="index"></param>
        /// <returns></returns>
        [BrowserEvent("推送D63評估資料")]
        [HttpPost]
        public JsonResult SaveD63(
            string action,
            string Auditor,
            D63ViewModel model,
            string assessmentSubKind,
            string bondNumber,
            string index
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.SaveD63(action, Auditor, model);
            if (result.RETURN_FLAG)
            {
                Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                    .FirstOrDefault(x => x.ToString() == index);
                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(model.Report_Date, out dt);
                string type = "All";
                if (assessmentSubKind != "All")
                    type = EnumUtil.GetValues<AssessmentSubKind_Type>()
                    .First(x => x.ToString() == assessmentSubKind)
                    .GetDescription();
                var datas = D6Repository.getD63(dt, bondNumber, _index, type);
                if (datas.Any())
                {
                    Cache.Invalidate(CacheList.D63DbfileData); //清除
                    Cache.Set(CacheList.D63DbfileData, datas); //把資料存到 Cache
                }
            }
            return Json(result);
        }
        #endregion

        #region Transfer D63
        /// <summary>
        /// D63 篩選需求量化評估資料
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        [BrowserEvent("篩選需求量化評估資料")]
        [HttpPost]
        public JsonResult TransferD63(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.TransferD63(reportDate);
            if (result.RETURN_FLAG)
            {
                Cache.Invalidate(CacheList.D63DbfileTransferData); //清除
                Cache.Set(CacheList.D63DbfileTransferData, D6Repository.TransferD63Data(reportDate)); //把資料存到 Cache
            }
            return Json(result);
        }
        #endregion

        #region D63 複核系列動作
        /// <summary>
        ///  D63 複核系列動作
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        [BrowserEvent("D63 複核系列動作")]
        [HttpPost]
        public JsonResult UpdateD63(
            string Reference_Nbr,
            int Assessment_Result_Version,
            string action
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                .FirstOrDefault(x => x.GetDescription() == action);
            Table_Type table = Table_Type.D63;
            //變更動作
            if (_index == Evaluation_Status_Type.NotReview)
                _index = Evaluation_Status_Type.Review;
            else if (_index == Evaluation_Status_Type.Review)
                _index = Evaluation_Status_Type.NotReview;
            else if (_index == Evaluation_Status_Type.Reject || _index == Evaluation_Status_Type.ReviewCompleted)
                table = Table_Type.D64;
            result = D6Repository.UpdateD63(Reference_Nbr, Assessment_Result_Version, _index);
            if (result.RETURN_FLAG)
            {
                if (!result.REASON_CODE.IsNullOrWhiteSpace())
                {
                    TxtLog.senMailLog(result.REASON_CODE, mailLogLocation("D63Mail.txt")); //寫txt Log
                }
                var datas = D6Repository.getD63History(Reference_Nbr);
                if (datas.Any())
                {
                    if (table == Table_Type.D63)
                    {                   
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.D63DbfileHistoryData); //清除
                        Cache.Set(CacheList.D63DbfileHistoryData, datas); //把資料存到 Cache
                        //return Json(result);                           
                    }
                    bool D64Flag = false;
                    if (table == Table_Type.D64)
                    {
                        D64Flag = true;
                    }
                    DateTime dt = DateTime.MinValue;
                    DateTime.TryParse(datas.First().Report_Date, out dt);

                    var D63SearchCache = (D6SearchCache)Cache.Get(CacheList.D63SearchCache);
                    var _datas = new List<D63ViewModel>();
                    if (D63SearchCache != null)
                    {
                        _datas = D6Repository.getD63(
                           D63SearchCache.dt,
                           D63SearchCache.bondNumber,
                           D63SearchCache.EST,
                           D63SearchCache.type,
                           D63SearchCache.Send_to_AuditorFlag,
                           D63SearchCache.referenceNbr);
                    }
                    else
                    {
                        _datas = D6Repository.getD63(dt, null, Evaluation_Status_Type.All, "All", D64Flag);
                    }
                    if (_datas.Any())
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.D63DbfileData); //清除
                        Cache.Set(CacheList.D63DbfileData, _datas); //把資料存到 Cache
                        return Json(result);
                    }
                }
            }
            return Json(result);
        }
        #endregion

        /// <summary>
        ///下載D63,D64Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載D63&D64Excel檔案")]
        [HttpPost]
        public JsonResult GetD63Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.D63DbfileData))
            {
                var D63 = Excel_DownloadName.D63.ToString();
                var D64 = Excel_DownloadName.D64.ToString();
                var D63Data = (List<D63ViewModel>)Cache.Get(CacheList.D63DbfileData);  //從Cache 抓資料
                result = D6Repository.DownLoadExcel2(D63, ExcelLocation(D63.GetExelName()), ExcelLocation(D64.GetExelName()), D63Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }

        #endregion

        #region D65 相關

        #region SearchD65
        /// <summary>
        /// 查詢D65(質化評估結果檔)資料
        /// </summary>
        /// <param name="date"></param>
        /// <param name="assessmentSubKind"></param>
        /// <param name="bondNumber"></param>
        /// <param name="referenceNbr"></param>
        /// <param name="index"></param>
        /// <param name="Send_to_AuditorFlag"></param>
        /// <returns></returns>
        [BrowserEvent("查詢D65質化評估結果")]
        [HttpPost]
        public JsonResult SearchD65(
            string date,
            string assessmentSubKind,
            string bondNumber,
            string referenceNbr,
            string index,
            bool Send_to_AuditorFlag = false)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            DateTime dt = DateTime.MinValue;
            if (!(assessmentSubKind == "All" ||
                EnumUtil.GetValues<AssessmentSubKind_Type>()
                .Any(x => x.ToString() == assessmentSubKind)) ||
                !DateTime.TryParse(date, out dt))
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                return Json(result);
            }
            string type = "All";
            if (assessmentSubKind != "All")
                type = EnumUtil.GetValues<AssessmentSubKind_Type>()
                .First(x => x.ToString() == assessmentSubKind)
                .GetDescription();
            Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                .FirstOrDefault(x => x.ToString() == index);
            var datas = D6Repository.getD65(dt, bondNumber, referenceNbr, _index, type, Send_to_AuditorFlag);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D65DbfileData); //清除
                Cache.Set(CacheList.D65DbfileData, datas); //把資料存到 Cache
                Cache.Invalidate(CacheList.D65SearchCache);//清除
                Cache.Set(CacheList.D65SearchCache, new D6SearchCache()
                {
                    dt = dt,
                    bondNumber = bondNumber,
                    EST = _index,
                    type = type,
                    Send_to_AuditorFlag = Send_to_AuditorFlag,
                    referenceNbr = referenceNbr
                });//把資料存到 Cache
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Data.GetDescription();
            }
            return Json(result);
        }
        #endregion

        #region GetD66
        /// <summary>
        /// Get D66
        /// </summary>
        /// <param name="referenceNbr">ReferenceNbr</param>
        /// <param name="version">Assessment_Result_Version</param>
        /// <returns></returns>
        [BrowserEvent("查詢D66(質化衡量複核資料)")]
        [HttpPost]
        public JsonResult GetD66(
            string referenceNbr,
            int version
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;         
            var datas = D6Repository.getD66(referenceNbr, version);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D66DbfileData); //清除
                Cache.Set(CacheList.D66DbfileData, datas); //把資料存到 Cache

                var D65SearchCache = (D6SearchCache)Cache.Get(CacheList.D65SearchCache);
                var _datas = new List<D65ViewModel>();
                if (D65SearchCache != null)
                {
                    _datas =
                        D6Repository.getD65(
                            D65SearchCache.dt, 
                            D65SearchCache.bondNumber, 
                            D65SearchCache.referenceNbr, 
                            D65SearchCache.EST, 
                            D65SearchCache.type,
                            D65SearchCache.Send_to_AuditorFlag);
                    if (datas.Any())
                    {
                        Cache.Invalidate(CacheList.D65DbfileData); //清除
                        Cache.Set(CacheList.D65DbfileData, _datas); //把資料存到 Cache
                    }
                }
            }
            if(!result.RETURN_FLAG)
                result.DESCRIPTION = Message_Type.not_Find_Data.GetDescription();
            return Json(result);
        }
        #endregion

        #region getD65 History
        /// <summary>
        /// 查詢D65複核版本
        /// </summary>
        /// <param name="referenceNbr"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        [BrowserEvent("查詢D65(質化衡量評估版本)")]
        [HttpPost]
        public JsonResult GetD65History(
            string referenceNbr,
            string version
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            int ver = 0;
            Int32.TryParse(version, out ver);
            var datas = D6Repository.getD65History(referenceNbr);
            if (datas.Any())
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D65DbfileHistoryData); //清除
                Cache.Set(CacheList.D65DbfileHistoryData, datas); //把資料存到 Cache
                return Json(result);
            }
            result.DESCRIPTION = "尚未執行評估";
            return Json(result);
        }
        #endregion

        #region SaveD65
        /// <summary>
        /// 推送D65評估資料
        /// </summary>
        /// <param name="Auditor"></param>
        /// <param name="D65Model"></param>
        /// <param name="D66Model"></param>
        /// <returns></returns>
        [BrowserEvent("推送D65評估資料")]
        [HttpPost]
        public JsonResult SaveD65(
            string Auditor,
            D65ViewModel D65Model,
            List<D66ViewModel> D66Model,
            string assessmentSubKind,
            string bondNumber,
            string index,
            string referenceNbr  
            )
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.SaveD65(Auditor, D65Model, D66Model);
            if (result.RETURN_FLAG)
            {
                DateTime dt = DateTime.MinValue;
                DateTime.TryParse(D65Model.Report_Date, out dt);
                Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                    .FirstOrDefault(x => x.ToString() == index);
                string type = "All";
                if (assessmentSubKind != "All")
                    type = EnumUtil.GetValues<AssessmentSubKind_Type>()
                    .First(x => x.ToString() == assessmentSubKind)
                    .GetDescription();
                var datas = D6Repository.getD65(dt, bondNumber, referenceNbr, _index, type);
                if (datas.Any())
                {
                    Cache.Invalidate(CacheList.D65DbfileData); //清除
                    Cache.Set(CacheList.D65DbfileData, datas); //把資料存到 Cache
                }
            }
            return Json(result);
        }
        #endregion

        #region Transfer D65
        /// <summary>
        /// D65 篩選需求質化評估資料
        /// </summary>
        /// <param name="reportDate"></param>
        /// <returns></returns>
        [BrowserEvent("篩選需求質化評估資料")]
        [HttpPost]
        public JsonResult TransferD65(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.TransferD65(reportDate);
            if (result.RETURN_FLAG)
            {
                Cache.Invalidate(CacheList.D65DbfileTransferData); //清除
                Cache.Set(CacheList.D65DbfileTransferData, D6Repository.TransferD65Data(reportDate)); //把資料存到 Cache
            }
            return Json(result);
        }
        #endregion

        #region D63 複核系列動作
        /// <summary>
        ///  D65 複核系列動作
        /// </summary>
        /// <param name="Reference_Nbr"></param>
        /// <param name="Assessment_Result_Version"></param>
        /// <param name="action"></param>
        /// <returns></returns>
        [BrowserEvent("D65 複核系列動作")]
        [HttpPost]
        public JsonResult UpdateD65(
            string Reference_Nbr,
            int Assessment_Result_Version,
            string action)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            Evaluation_Status_Type _index = EnumUtil.GetValues<Evaluation_Status_Type>()
                .FirstOrDefault(x => x.GetDescription() == action);
            Table_Type table = Table_Type.D65;
            //變更動作
            if (_index == Evaluation_Status_Type.NotReview)
                _index = Evaluation_Status_Type.Review;
            else if (_index == Evaluation_Status_Type.Review)
                _index = Evaluation_Status_Type.NotReview;
            else if (_index == Evaluation_Status_Type.Reject || _index == Evaluation_Status_Type.ReviewCompleted)
                table = Table_Type.D66;
            result = D6Repository.UpdateD65(Reference_Nbr, Assessment_Result_Version, _index);
            if (result.RETURN_FLAG)
            {
                var datas = D6Repository.getD65History(Reference_Nbr);
                if (datas.Any())
                {
                    if (table == Table_Type.D65)
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.D65DbfileHistoryData); //清除
                        Cache.Set(CacheList.D65DbfileHistoryData, datas); //把資料存到 Cache                         
                    }
                    bool D66Flag = false;
                    if (table == Table_Type.D66)
                    {
                        D66Flag = true;
                    }
                    DateTime dt = DateTime.MinValue;
                    DateTime.TryParse(datas.First().Report_Date, out dt);
                    var D65SearchCache = (D6SearchCache)Cache.Get(CacheList.D65SearchCache);
                    var _datas = new List<D65ViewModel>();
                    if (D65SearchCache != null)
                    {
                        _datas =
                            D6Repository.getD65(
                                D65SearchCache.dt,
                                D65SearchCache.bondNumber,
                                D65SearchCache.referenceNbr,
                                D65SearchCache.EST,
                                D65SearchCache.type,
                                D65SearchCache.Send_to_AuditorFlag);
                    }
                    else
                    {
                        _datas = D6Repository.getD65(dt, null, null, Evaluation_Status_Type.All, "All", D66Flag);
                    }
                    if (_datas.Any())
                    {
                        result.RETURN_FLAG = true;
                        Cache.Invalidate(CacheList.D65DbfileData); //清除
                        Cache.Set(CacheList.D65DbfileData, _datas); //把資料存到 Cache
                        return Json(result);
                    }
                }
            }
            return Json(result);
        }
        #endregion

        #region 下載D65&D66Excel檔案
        /// <summary>
        ///下載D65,D66Excel檔案
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("下載D65&D66Excel檔案")]
        [HttpPost]
        public JsonResult GetD65Excel()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;

            if (Cache.IsSet(CacheList.D65DbfileData))
            {
                var D65 = Excel_DownloadName.D65.ToString();
                var D66 = Excel_DownloadName.D66.ToString();
                var D65Data = (List<D65ViewModel>)Cache.Get(CacheList.D65DbfileData);  //從Cache 抓資料
                result = D6Repository.DownLoadExcel2(D65, ExcelLocation(D65.GetExelName()), ExcelLocation(D66.GetExelName()), D65Data);
            }
            else
            {
                result.DESCRIPTION = Message_Type.time_Out.GetDescription();
            }
            return Json(result);
        }
        #endregion


        #endregion

        [HttpPost]
        public JsonResult GetAuditor(string tableId)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var users = getAssessment(productCode, tableId, SetAssessmentType.Auditor)
                .Select(x => new SelectOption()
                {
                    Text = $"{x.User_Account}({x.User_Name})" ,
                    Value = x.User_Account
                }).ToList();
            if (users.Any())
            {
                result.RETURN_FLAG = true;
                //users.Insert(0, new SelectOption() { Text = " ", Value = " " });
                result.Datas = Json(users);
            }
            return Json(result);
        }



        #endregion

        #region 減損報表資料刪除作業(複核者使用for報表重新產製)

        #region 取得產品組合

        /// <summary>
        /// 查詢 flowId 
        /// </summary>
        /// <param name="date">基準日</param>
        /// <returns></returns>
        [BrowserEvent("查詢 flowId")]
        [HttpPost]
        public JsonResult GetFlowId(string date)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            List<string> data = new List<string>();
            data = D6Repository.getFlowIDs(date);
            if (data.Any())
            {
                result.RETURN_FLAG = true;
                result.Datas = Json(data);
            }
            return Json(result);
        }
        #endregion

        #region 查詢ReEL

        /// <summary>
        /// 查詢ReEL
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="flowId"></param>
        /// <returns></returns>
        [BrowserEvent("查詢ReEL")]
        [HttpPost]
        public JsonResult GetReEL(string reportDate,string Group_Product_Code)
        {
            MSGReturnModel result = new MSGReturnModel();
            var data = D6Repository.getReEL(reportDate, Group_Product_Code);
            if (data.Any())
            {
                Cache.Invalidate(CacheList.ReELDbfileData); //清除
                Cache.Set(CacheList.ReELDbfileData, data); //把資料存到 Cache
                result.RETURN_FLAG = true;
            }
            else
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            }
            return Json(result);
        }
        #endregion

        #region 執行減損報表資料刪除作業

        /// <summary>
        /// 執行減損報表資料刪除作業
        /// </summary>
        /// <param name="reportDate"></param>
        /// <param name="flowId"></param>
        /// <param name="msg"></param>
        /// <returns></returns>
        [BrowserEvent("執行減損報表資料刪除作業")]
        [HttpPost]
        public JsonResult ReSetEL(string reportDate, string flowId, string msg)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.ReEL(reportDate, flowId, msg);
            return Json(result);
        }

        #endregion

        #endregion

        private class D54Group
        {
            public string Product_Code { get; set; }
            public string Bond_A { get; set; }
            public string Bond_B { get; set; }
            public string Bond_P { get; set; }
        }

        #region private function
        private List<SelectOption> getAssessmentSubKind_Type(AssessmentKind_Type type)
        {
            List<SelectOption> assessmentSubKind = new List<SelectOption>();
            assessmentSubKind = D6Repository.getAssessmentSubKind(type);
            return assessmentSubKind;
        }
       
        public class D6SearchCache
        {
            public DateTime dt { get; set; }
            public string bondNumber { get; set; }
            public Evaluation_Status_Type EST { get; set; }
            public string type { get; set; }
            public bool Send_to_AuditorFlag { get; set;}
            public string referenceNbr { get; set; }
        }

        #endregion

        #region GetD67Data
        [BrowserEvent("查詢D67資料")]
        [HttpPost]
        public JsonResult GetD67Data(D67ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            try
            {
                var queryData = D6Repository.getD67(dataModel);
                result.RETURN_FLAG = queryData.Item1;
                Cache.Invalidate(CacheList.D67DbfileData); //清除
                Cache.Set(CacheList.D67DbfileData, queryData.Item2); //把資料存到 Cache
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

        #region GetCacheD67Data
        [HttpPost]
        public JsonResult GetCacheD67Data(jqGridParam jdata)
        {
            List<D67ViewModel> data = new List<D67ViewModel>();
            if (Cache.IsSet(CacheList.D67DbfileData))
            {
                data = (List<D67ViewModel>)Cache.Get(CacheList.D67DbfileData);
            }

            return Json(jdata.modelToJqgridResult(data, true));
        }
        #endregion

        #region GetD67Excel
        [BrowserEvent("D67匯出")]
        [HttpPost]
        public ActionResult GetD67Excel(D67ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            string path = string.Empty;

            try
            {
                var data = D6Repository.getD67(dataModel);

                path = "D67".GetExelName();
                result = D6Repository.DownLoadD67Excel(ExcelLocation(path), data.Item2);
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.download_Fail
                                     .GetDescription(ex.Message);
            }

            return Json(result);
        }

        #endregion

        #region SaveD67
        [BrowserEvent("D67執行評估")]
        [HttpPost]
        public JsonResult SaveD67(string reportDate)
        {
            MSGReturnModel result = new MSGReturnModel();

            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();

            try
            {
                MSGReturnModel resultSave = D6Repository.saveD67(reportDate);

                result.RETURN_FLAG = resultSave.RETURN_FLAG;
                result.DESCRIPTION = resultSave.DESCRIPTION;

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

        #region D53 相關
        /// <summary>
        /// 查詢A46資料
        /// </summary>
        /// <param name="searchAll"></param>
        /// <returns></returns>
        [BrowserEvent("查詢D53資料")]
        [HttpPost]
        public JsonResult GetD53Data()
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var Datas = D5Repository.GetD53();
            if (!Datas.Item1.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = Datas.Item1;
            }
            else if (!Datas.Item2.Any())
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(Table_Type.D53.tableNameGetDescription());
            }
            else
            {
                result.RETURN_FLAG = true;
                Cache.Invalidate(CacheList.D53DbfileData); //清除
                Cache.Set(CacheList.D53DbfileData, Datas.Item2); //把資料存到 Cache
            }
            return Json(result);
        }

        /// <summary>
        /// 轉檔把Excel 資料存到 DB
        /// </summary>
        /// <returns></returns>
        [BrowserEvent("D53 Excel檔存到DB")]
        [HttpPost]
        public JsonResult TransferD53()
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 抓Excel檔案 轉成 model

                // Excel 檔案位置

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads);

                string fileName = string.Empty;
                if (Cache.IsSet(CacheList.D53ExcelName))
                    fileName = (string)Cache.Get(CacheList.D53ExcelName);  //從Cache 抓資料

                if (fileName.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.time_Out.GetDescription();
                    return Json(result);
                }

                string path = Path.Combine(projectFile, fileName);

                List<D53ViewModel> dataModel = new List<D53ViewModel>();

                string errorMessage = string.Empty;

                using (FileStream stream = System.IO.File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //Excel轉成 Exhibit10Model
                    string pathType = path.Split('.')[1]; //抓副檔名
                    var data = D5Repository.GetD53Excel(pathType, stream);
                    dataModel = data.Item2;
                    errorMessage = data.Item1;
                }
                if (!errorMessage.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = errorMessage;
                    return Json(result);
                }
                #endregion 抓Excel檔案 轉成 model

                #region txtlog 檔案名稱

                string txtpath = SetFile.D53TransferTxtLog; //預設txt名稱

                #endregion txtlog 檔案名稱

                #region save SMF_Info(D53)

                MSGReturnModel resultD53 = D5Repository.SaveD53(dataModel); //save to DB
                CommonFunction.saveLog(Table_Type.D53,
                    fileName, SetFile.ProgramName, resultD53.RETURN_FLAG,
                    Debt_Type.B.ToString(), startTime, DateTime.Now, AccountController.CurrentUserInfo.Name); //寫sql Log
                TxtLog.txtLog(Table_Type.D53, resultD53.RETURN_FLAG, startTime, txtLocation(txtpath)); //寫txt Log

                #endregion save SMF_Info(D53)

                result.RETURN_FLAG = resultD53.RETURN_FLAG;
                result.DESCRIPTION = Message_Type.save_Success.GetDescription(Table_Type.D53.ToString());

                if (!result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.save_Fail
                        .GetDescription(Table_Type.D53.ToString(), resultD53.DESCRIPTION);
                }
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.exceptionMessage();
            }
            return Json(result);
        }

        /// <summary>
        /// 選擇檔案後點選資料上傳觸發(D53)
        /// </summary>
        /// <param name="FileModel"></param>
        /// <returns></returns>
        [BrowserEvent("D53上傳檔案")]
        [HttpPost]
        public JsonResult UploadD53(ValidateFiles FileModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                #region 前端無傳送檔案進來

                if (FileModel.File == null)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.upload_Not_Find.GetDescription();
                    return Json(result);
                }

                #endregion 前端無傳送檔案進來

                #region 前端檔案大小不符或不為Excel檔案(驗證)

                //ModelState
                if (!ModelState.IsValid)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.excel_Validate.GetDescription();
                    return Json(result);
                }

                #endregion 前端檔案大小不符或不為Excel檔案(驗證)

                #region 上傳檔案

                string pathType = Path.GetExtension(FileModel.File.FileName)
                                       .Substring(1); //上傳的檔案類型

                var fileName = string.Format("{0}.{1}",
                    Excel_UploadName.D53.GetDescription(),
                    pathType); //固定轉成此名稱

                Cache.Invalidate(CacheList.D53ExcelName); //清除 Cache
                Cache.Set(CacheList.D53ExcelName, fileName); //把資料存到 Cache

                #region 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                string projectFile = Server.MapPath("~/" + SetFile.FileUploads); //專案資料夾
                string path = Path.Combine(projectFile, fileName);

                FileRelated.createFile(projectFile); //檢查是否有FileUploads資料夾,如果沒有就新增

                //呼叫上傳檔案 function
                result = FileRelated.FileUpLoadinPath(path, FileModel.File);
                if (!result.RETURN_FLAG)
                    return Json(result);

                #endregion 檢查是否有FileUploads資料夾,如果沒有就新增 並加入 excel 檔案

                #region 讀取Excel資料
                var data = D5Repository.GetD53Excel(pathType, FileModel.File.InputStream);
                if (data.Item1.IsNullOrWhiteSpace())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.D53ExcelfileData); //清除 Cache
                    Cache.Set(CacheList.D53ExcelfileData, data.Item2); //把資料存到 Cache
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = data.Item1;
                }

                #endregion 讀取Excel資料 

                #endregion 上傳檔案
            }
            catch (Exception ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;
            }
            return Json(result);
        }

        #region 下載 D53範例 Excel
        [HttpGet]
        public ActionResult DownloadD53TempExcel(string type)
        {
            try
            {
                string path = string.Empty;
                if (EnumUtil.GetValues<Excel_DownloadName>()
                    .Any(x => x.ToString().Equals(type)))
                {
                    path = type.GetExelName();
                    return File(ExcelLocation(path), "application/octet-stream", path);
                }
            }
            catch
            {
            }
            return new HttpStatusCodeResult(System.Net.HttpStatusCode.Forbidden);
        }
        #endregion

        #endregion

        #region D6Check 相關

        [BrowserEvent("查詢減損作業狀態")]
        [HttpPost]
        public JsonResult GetD6Check(string reportdate, string Impairment)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
            List<D6CheckViewModel> datas = D6Repository.getD6Check(reportdate);
            if (datas.Any())
            {
                if (Impairment != "All")
                {
                    var Job_Details = EnumUtil.GetValues<Impairment_Operations_Type>().FirstOrDefault(x => x.ToString() == Impairment).GetDescription();
                    datas = datas.Where(x => x.Job_Details == Job_Details).ToList();                
                }
                if (datas.Any())
                {
                    result.RETURN_FLAG = true;
                    result.Datas = Json(datas);
                }
            }
            return Json(result);
        }
        #endregion

        #region D6Mail 相關

        public JsonResult GetD6MailMsg()
        {
            string result = string.Empty;

            result = TxtLog.getLog(mailLogLocation("D63Mail.txt")); //寫txt Log

            return Json(result);
        }

        #endregion

        #region Extra_Case

        public JsonResult GetA41Version(string date)
        {
            int version = 0;
            version = D6Repository.GetA41Version(date);
            return Json(version); 
        }


        #region 190619 PGE需求新增
        public JsonResult GetA41AssessmentCheck(string date, int version)
        {
            int number = 0;
            number= D6Repository.GetA41AssessmentCheck(date, version).Count();
            return Json(number);
        }
        public JsonResult AutoAddExtraCase(string reportDate, int version)
        {
            var ExtraCase = D6Repository.GetA41AssessmentCheck(reportDate, version);       
            MSGReturnModel result = new MSGReturnModel();            
            result = D6Repository.AutoInsertD65ExtraCase(ExtraCase,reportDate);
            return Json(result);
        }
        #endregion

        public JsonResult AddExtraCase(string reportDate, int version, string bondNumber, string Lots, string portfolio_Name)
        {
            MSGReturnModel result = new MSGReturnModel();
            result = D6Repository.InsertD65ByExtraCase(reportDate, version, bondNumber?.Trim(), Lots?.Trim(), portfolio_Name?.Trim());
            return Json(result);
        }

        public JsonResult DelExtraCase(string referenceNbr,int version)
        {
            MSGReturnModel result = new MSGReturnModel();

            result = D6Repository.DeleteD65ByExtraCase(referenceNbr, version);
            if (result.RETURN_FLAG)
            {
                var D65SearchCache = (D6SearchCache)Cache.Get(CacheList.D65SearchCache);
                var _datas = new List<D65ViewModel>();
                if (D65SearchCache != null)
                {
                    _datas =
                        D6Repository.getD65(
                            D65SearchCache.dt,
                            D65SearchCache.bondNumber,
                            D65SearchCache.referenceNbr,
                            D65SearchCache.EST,
                            D65SearchCache.type,
                            D65SearchCache.Send_to_AuditorFlag);
                }
                if (_datas.Any())
                {
                    result.RETURN_FLAG = true;
                    Cache.Invalidate(CacheList.D65DbfileData); //清除
                    Cache.Set(CacheList.D65DbfileData, _datas); //把資料存到 Cache
                    return Json(result);
                }
            }
            return Json(result);
        }

        #endregion
    }
}