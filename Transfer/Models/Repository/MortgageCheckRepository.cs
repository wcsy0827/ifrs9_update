using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using Transfer.Models.Abstract;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class MortgageCheckRepository<T> : CheckDataAbstract<T> where T : class
    {
        public MortgageCheckRepository(IEnumerable<T> data, Check_Table_Type _event) : base(data, _event)
        {
        }

        protected override void Set()
        {
            _resources.Add(Check_Table_Type.Mortgage_B01_Before_Check, B01TransferCheck_Before);
            _resources.Add(Check_Table_Type.Mortgage_B01_Transfer_Check, B01dbModelCheck);
            _resources.Add(Check_Table_Type.Mortgage_C01_Before_Check, C01TransferCheck_Before);
            _resources.Add(Check_Table_Type.Mortgage_C01_Transfer_Check, C01dbModelCheck);
            _resources.Add(Check_Table_Type.Mortgage_C02_Before_Check, C02TransferCheck_Before);
            _resources.Add(Check_Table_Type.Mortgage_C02_Transfer_Check, C02dbModelCheck);
        }

        #region B01TransferCheck_Before
        private List<messageTable> B01TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                messageTable mt = new messageTable()
                {
                    title = @"檢查A01(Loan_IAS39_Info)、A02(Loan_Account_Info)資料是否為空值",
                    successStr = @"來源資料A01(Loan_IAS39_Info)、A02(Loan_Account_Info)內重要欄位皆非空值"
                };

                var data = _data.Cast<Loan_IAS39_Info>().ToList();

                StringBuilder sbEIR = new StringBuilder();
                StringBuilder sbPrincipal = new StringBuilder();
                StringBuilder sbInterest_Receivable = new StringBuilder();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                    if(x.EIR.ToString() == null || x.EIR.ToString() == "")
                    {
                        sbEIR.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Principal.ToString() == null || x.Principal.ToString() == "")
                    {
                        sbPrincipal.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Interest_Receivable.ToString() == null || x.Interest_Receivable.ToString() == "")
                    {
                        sbInterest_Receivable.Append(_parameter + Environment.NewLine);
                    }
                });

                StringBuilder sbDelinquent_Days = new StringBuilder();
                StringBuilder sbCurrent_Rating_Code = new StringBuilder();
                StringBuilder sbCurrent_Int_Rate = new StringBuilder();
                StringBuilder sbLexp_Date = new StringBuilder();
                StringBuilder sbCurrent_Lgd = new StringBuilder();
                var _first = data.FirstOrDefault();
                List<Loan_Account_Info> A02data = new List<Loan_Account_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (_first != null)
                    {
                        A02data = db.Loan_Account_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        A02data.ForEach(x =>
                        {
                            var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                            if (x.Delinquent_Days.ToString() == null || x.Delinquent_Days.ToString() == "")
                            {
                                sbDelinquent_Days.Append(_parameter + Environment.NewLine);
                            }

                            if (x.Current_Rating_Code == null || x.Current_Rating_Code == "")
                            {
                                sbCurrent_Rating_Code.Append(_parameter + Environment.NewLine);
                            }

                            if (x.Current_Int_Rate.ToString() == null || x.Current_Int_Rate.ToString() == "")
                            {
                                sbCurrent_Int_Rate.Append(_parameter + Environment.NewLine);
                            }

                            if (x.Lexp_Date == null || x.Lexp_Date == "")
                            {
                                sbLexp_Date.Append(_parameter + Environment.NewLine);
                            }

                            if (x.Current_LGD.ToString() == null || x.Current_LGD.ToString() == "")
                            {
                                sbCurrent_Lgd.Append(_parameter + Environment.NewLine);
                            }
                        });
                    }
                }

                if (sbEIR.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "有效利率 EIR 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbEIR.ToString();
                }

                if (sbPrincipal.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "剩餘本金餘額 Principal 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbPrincipal.ToString();
                }

                if (sbInterest_Receivable.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "應收利息 Interest_Receivable 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbInterest_Receivable.ToString();
                }

                if (sbDelinquent_Days.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "違約天數 Delinquent_Days 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbDelinquent_Days.ToString();
                }

                if (sbCurrent_Rating_Code.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "信用評等 Current_Rating_Code 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Rating_Code.ToString();
                }

                if (sbCurrent_Int_Rate.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "合約利率 Current_Int_Rate 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate.ToString();
                }

                if (sbLexp_Date.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "到期日 Lexp_Date 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbLexp_Date.ToString();
                }

                if (sbCurrent_Lgd.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "違約損失率 Current_Lgd 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Lgd.ToString();
                }
            }

            return result;
        }
        #endregion

        #region B01dbModelCheck
        private List<messageTable> B01dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<IFRS9_Main>().ToList();

                messageTable mt = new messageTable()
                {
                    title = @"彙總資訊",
                    successStr = @"轉檔完成"
                };

                StringBuilder sbEIR = new StringBuilder();
                StringBuilder sbPrincipal = new StringBuilder();
                StringBuilder sbInterest_Receivable = new StringBuilder();
                StringBuilder sbDelinquent_Days = new StringBuilder();
                StringBuilder sbCurrent_Rating_Code = new StringBuilder();
                StringBuilder sbCurrent_Int_Rate = new StringBuilder();
                StringBuilder sbLexp_Date = new StringBuilder();
                StringBuilder sbCurrent_Lgd = new StringBuilder();
                StringBuilder sbCurrent_Int_Rate_NotValid = new StringBuilder();
                StringBuilder sbEIR_NotValid = new StringBuilder();
                StringBuilder sbCurrent_Lgd_NotValid = new StringBuilder();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                    if (x.Eir.ToString() == null || x.Eir.ToString() == "")
                    {
                        sbEIR.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Principal.ToString() == null || x.Principal.ToString() == "")
                    {
                        sbPrincipal.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Interest_Receivable.ToString() == null || x.Interest_Receivable.ToString() == "")
                    {
                        sbInterest_Receivable.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Delinquent_Days.ToString() == null || x.Delinquent_Days.ToString() == "")
                    {
                        sbDelinquent_Days.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Current_Rating_Code == null || x.Current_Rating_Code == "")
                    {
                        sbCurrent_Rating_Code.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Current_Int_Rate.ToString() == null || x.Current_Int_Rate.ToString() == "")
                    {
                        sbCurrent_Int_Rate.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Maturity_Date.ToString() == null || x.Maturity_Date.ToString() == "")
                    {
                        sbLexp_Date.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Current_Lgd.ToString() == null || x.Current_Lgd.ToString() == "")
                    {
                        sbCurrent_Lgd.Append(_parameter + Environment.NewLine);
                    }

                    if ((x.Current_Int_Rate > 0 && x.Current_Int_Rate < 1) == false)
                    {
                        sbCurrent_Int_Rate_NotValid.Append(_parameter + "," + x.Current_Int_Rate.ToString() + Environment.NewLine);
                    }

                    if ((x.Eir > 0 && x.Eir < 1) == false)
                    {
                        sbEIR_NotValid.Append(_parameter + "," + x.Eir.ToString() + Environment.NewLine);
                    }

                    if ((x.Current_Lgd >= 0 && x.Current_Lgd <= 1) == false)
                    {
                        sbCurrent_Lgd_NotValid.Append(_parameter + "," + x.Current_Lgd.ToString() + Environment.NewLine);
                    }
                });

                //var HaveNull = false;
                //var IsNotValid = false;

                if (sbEIR.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "有效利率 EIR 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbEIR.ToString();
                }

                if (sbPrincipal.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "剩餘本金餘額 Principal 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbPrincipal.ToString();
                }

                if (sbInterest_Receivable.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "應收利息 Interest_Receivable 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbInterest_Receivable.ToString();
                }

                if (sbDelinquent_Days.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "違約天數 Delinquent_Days 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbDelinquent_Days.ToString();
                }

                if (sbCurrent_Rating_Code.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "信用評等 Current_Rating_Code 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Rating_Code.ToString();
                }

                if (sbCurrent_Int_Rate.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "合約利率 Current_Int_Rate 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate.ToString();
                }

                if (sbLexp_Date.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "到期日 Lexp_Date 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbLexp_Date.ToString();
                }

                if (sbCurrent_Lgd.Length > 0)
                {
                    //HaveNull = true;
                    _customerStr_End += "違約損失率 Current_Lgd 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Lgd.ToString();
                }

                //if (HaveNull == false)
                //{
                //    _customerStr_End += "B01(IFRS9_Main)內重要欄位皆非空值" + Environment.NewLine;
                //}

                if (sbCurrent_Int_Rate_NotValid.Length > 0)
                {
                    //IsNotValid = true;
                    _customerStr_End += "合约利率 Current_Int_Rate (必須 > 0 且 < 1)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate_NotValid.ToString();
                }

                if (sbEIR_NotValid.Length > 0)
                {
                    //IsNotValid = true;
                    _customerStr_End += "有效利率 EIR (必須 > 0 且 < 1)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbEIR_NotValid.ToString();
                }

                if (sbCurrent_Lgd_NotValid.Length > 0)
                {
                    //IsNotValid = true;
                    _customerStr_End += "違約損失率 Current_Lgd (必須 >= 0 且 <= 1)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Lgd_NotValid.ToString();
                }

                //if (IsNotValid == false)
                //{
                //    _customerStr_End += "減損計算表B01需求欄位已通過合格值檢核" + Environment.NewLine;
                //}

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("彙總資訊");

                sb.AppendLine($@"◆B01(IFRS9_Main)總資料筆數：{data.Count} 筆");

                sb.AppendLine("◆違約損失率 Current_LGD 值分布統計");
                var CurrentLgdCountQuery = from a in data
                                           group a by a.Current_Lgd into grouping
                                           select new
                                           {
                                               grouping.Key,
                                               dataCount = grouping.Count()
                                           };
                foreach (var item in CurrentLgdCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆信用評等 Current_Rating_Code 筆數統計");
                var CurrentRatingCodeCountQuery = from a in data
                                                  group a by a.Current_Rating_Code into grouping
                                                  select new
                                                  {
                                                      grouping.Key,
                                                      dataCount = grouping.Count()
                                                  };
                foreach (var item in CurrentRatingCodeCountQuery)
                {
                    string itemKey = (item.Key.IsNullOrWhiteSpace()) ? "空值" : item.Key;
                    sb.AppendLine($"{itemKey}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆假扣押 Collateral_Legal_Action_Ind 筆數統計");
                var CollateralLegalActionIndCountQuery = from a in data
                                                         group a by a.Collateral_Legal_Action_Ind into grouping
                                                         select new
                                                         {
                                                             grouping.Key,
                                                             dataCount = grouping.Count()
                                                         };
                foreach (var item in CollateralLegalActionIndCountQuery)
                {
                    string itemKey = (item.Key.IsNullOrWhiteSpace()) ? "空值":item.Key;
                    sb.AppendLine($"{itemKey}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆逾期天數 Delinquent_Days 值分布統計");
                var DelinquentDaysCountQuery = from a in data
                                               group a by a.Delinquent_Days into grouping
                                               select new
                                               {
                                                   grouping.Key,
                                                   dataCount = grouping.Count()
                                               };
                foreach (var item in DelinquentDaysCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆是否具客觀減損證據 IAS39_Impaire_Ind 筆數統計");
                var IAS39ImpaireIndCountQuery = from a in data
                                                group a by a.Ias39_Impaire_Ind into grouping
                                                select new
                                                {
                                                    grouping.Key,
                                                    dataCount = grouping.Count()
                                                };
                foreach (var item in IAS39ImpaireIndCountQuery)
                {
                    string itemKey = (item.Key.IsNullOrWhiteSpace()) ? "空值" : item.Key;
                    sb.AppendLine($"{itemKey}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆符合客觀減損證據之條件 IAS39_Impaire_Desc 筆數統計");
                var IAS39ImpaireDescCountQuery = from a in data
                                                 group a by a.Ias39_Impaire_Desc into grouping
                                                 select new
                                                 {
                                                     grouping.Key,
                                                     dataCount = grouping.Count()
                                                 };
                foreach (var item in IAS39ImpaireDescCountQuery)
                {
                    string itemKey = (item.Key.IsNullOrWhiteSpace()) ? "空值" : item.Key;
                    sb.AppendLine($"{itemKey}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆紓困註記 Restructure_Ind 筆數統計");
                var Restructure_IndCountQuery = from a in data
                                                group a by a.Restructure_Ind into grouping
                                                select new
                                                {
                                                    grouping.Key,
                                                    dataCount = grouping.Count()
                                                };
                foreach (var item in Restructure_IndCountQuery)
                {
                    string itemKey = (item.Key.IsNullOrWhiteSpace()) ? "空值" : item.Key;
                    sb.AppendLine($"{itemKey}：{item.dataCount} 筆");
                }

                sb.AppendLine("");

                var _first = data.FirstOrDefault();
                List<Loan_IAS39_Info> A01data = new List<Loan_IAS39_Info>();
                List<Loan_Account_Info> A02data = new List<Loan_Account_Info>();
                List<Loan_IAS39_Info> A01North = new List<Loan_IAS39_Info>();
                List<Loan_IAS39_Info> A01NotNorth = new List<Loan_IAS39_Info>();
                List<Loan_IAS39_Info> A01_5Data = new List<Loan_IAS39_Info>();
                List<Loan_Account_Info> A02_5Data = new List<Loan_Account_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (_first != null)
                    {
                        A01data = db.Loan_IAS39_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        A02data = db.Loan_Account_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        A01North = A01data.Where(x => x.NO34RCV == "北區").ToList();

                        A01NotNorth = A01data.Where(x => x.NO34RCV == "非北區").ToList();
                    }
                }

                var B01North = (from a in data
                                join b in A01North on a.Reference_Nbr equals b.Reference_Nbr
                                select a).ToList();
                var B01NorthGroupData = B01North.GroupBy(x => new { x.Current_Lgd })
                                                .Select(x => x.FirstOrDefault()).ToList();
                string NorthCurrentLgd = B01NorthGroupData.FirstOrDefault().Current_Lgd.ToString();
                string NorthCurrentLgdCount = B01North.Where(x => x.Current_Lgd.ToString() == NorthCurrentLgd).Count().ToString();

                var B01NotNorth = (from a in data
                                   join b in A01NotNorth on a.Reference_Nbr equals b.Reference_Nbr
                                   select a).ToList();
                var B01NotNorthGroupData = B01NotNorth.GroupBy(x => new { x.Current_Lgd })
                                                      .Select(x => x.FirstOrDefault()).ToList();
                string NotNorthCurrentLgd = B01NotNorthGroupData.FirstOrDefault().Current_Lgd.ToString();
                string NotNorthCurrentLgdCount = B01NotNorth.Where(x => x.Current_Lgd.ToString() == NotNorthCurrentLgd).Count().ToString();

                A01_5Data = A01data.Where(x => x.IAS39_Impaire_Ind == "Y" && x.IAS39_Impaire_Desc != "逾期 29天").ToList();
                A02_5Data = A02data.Where(x => x.Delinquent_Days == 100).ToList();
                A02_5Data = (from a in A02_5Data
                             join b in A01_5Data on a.Reference_Nbr equals b.Reference_Nbr
                             select a).ToList();

                sb.AppendLine($"◆A01(Loan_IAS39_Info)：{A01data.Count} 筆     B01(IFRS9_Main)：{data.Count} 筆");
                sb.AppendLine("◆A01(Loan_IAS39_Info) 回收率區隔");
                sb.AppendLine($"北區：{A01North.Count} 筆     非北區：{A01NotNorth.Count} 筆");
                sb.AppendLine("◆B01(IFRS9_Main) 違約損失率");
                sb.AppendLine($"{NorthCurrentLgd}：{NorthCurrentLgdCount} 筆     {NotNorthCurrentLgd}：{NotNorthCurrentLgdCount} 筆");
                sb.AppendLine($"◆B01(IFRS9_Main) 信用評等等級5：{data.Where(x=>x.Current_Rating_Code == "5").Count()} 筆");
                sb.AppendLine($"  應用規則計算：{A02_5Data.Count()} 筆");

                _customerStr_End += sb.ToString();
            }

            return result;
        }
        #endregion

        #region C01TransferCheck_Before
        private List<messageTable> C01TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                messageTable mt = new messageTable()
                {
                    title = @"檢查B01(IFRS9_Main)資料是否為空值",
                    successStr = @"來源資料B01(IFRS9_Main)內重要欄位皆非空值"
                };

                var data = _data.Cast<IFRS9_Main>().ToList();

                StringBuilder sbProduct_Code = new StringBuilder();
                StringBuilder sbCurrent_Rating_Code = new StringBuilder();
                StringBuilder sbCurrent_Int_Rate = new StringBuilder();
                StringBuilder sbEIR = new StringBuilder();
                StringBuilder sbOri_Amount = new StringBuilder();
                StringBuilder sbPrincipal = new StringBuilder();
                StringBuilder sbInterest_Receivable = new StringBuilder();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                    if (x.Product_Code == null || x.Product_Code == "")
                    {
                        sbProduct_Code.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Current_Rating_Code == null || x.Current_Rating_Code == "")
                    {
                        sbCurrent_Rating_Code.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Current_Int_Rate.ToString() == null || x.Current_Int_Rate.ToString() == "")
                    {
                        sbCurrent_Int_Rate.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Eir.ToString() == null || x.Eir.ToString() == "")
                    {
                        sbEIR.Append(_parameter + Environment.NewLine);
                    }

                    //if (x.Ori_Amount.ToString() == null || x.Ori_Amount.ToString() == "")
                    //{
                    //    sbOri_Amount.Append(_parameter + Environment.NewLine);
                    //}

                    if (x.Principal.ToString() == null || x.Principal.ToString() == "")
                    {
                        sbPrincipal.Append(_parameter + Environment.NewLine);
                    }

                    if (x.Interest_Receivable.ToString() == null || x.Interest_Receivable.ToString() == "")
                    {
                        sbInterest_Receivable.Append(_parameter + Environment.NewLine);
                    }
                });

                if (sbProduct_Code.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "產品 Product_Code 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbProduct_Code.ToString();
                }

                if (sbCurrent_Rating_Code.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "風險區隔 Current_Rating_Code 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Rating_Code.ToString();
                }

                if (sbCurrent_Int_Rate.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "合约利率/產品利率 Current_Int_Rate 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate.ToString();
                }

                if (sbEIR.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "有效利率 EIR 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate.ToString();
                }

                //if (sbOri_Amount.Length > 0)
                //{
                //    _checkFlag = true;
                //    _customerStr_End += "原始購買金額 Ori_Amount 有空值，有問題的帳號：" + Environment.NewLine;
                //    _customerStr_End += sbOri_Amount.ToString();
                //}

                if (sbPrincipal.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "金融資產餘額 Principal 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbPrincipal.ToString();
                }

                if (sbInterest_Receivable.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "應收利息 Interest_Receivable 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbInterest_Receivable.ToString();
                }
            }

            return result;
        }
        #endregion

        #region C01dbModelCheck
        private List<messageTable> C01dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                messageTable mt = new messageTable()
                {
                    title = @"彙總資訊",
                    successStr = @"轉檔完成"
                };

                var data = _data.Cast<EL_Data_In>().ToList();

                StringBuilder sbActual_Year_To_Maturity = new StringBuilder();
                StringBuilder sbDuration_Year = new StringBuilder();
                StringBuilder sbRemaining_Month = new StringBuilder();
                StringBuilder sbCurrent_Int_Rate = new StringBuilder();
                StringBuilder sbEIR = new StringBuilder();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                    if (x.Actual_Year_To_Maturity == null || x.Actual_Year_To_Maturity.ToString() == "")
                    {
                        sbActual_Year_To_Maturity.Append(_parameter + Environment.NewLine);
                    }

                    if ((x.Duration_Year > 0) == false)
                    {
                        sbDuration_Year.Append(_parameter + "," + x.Duration_Year.ToString() + Environment.NewLine);
                    }

                    if ((x.Remaining_Month > 0) == false)
                    {
                        sbRemaining_Month.Append(_parameter + "," + x.Remaining_Month.ToString() + Environment.NewLine);
                    }

                    if ((x.Current_Int_Rate > 0 && x.Current_Int_Rate < 1) == false)
                    {
                        sbCurrent_Int_Rate.Append(_parameter + "," + x.Current_Int_Rate.ToString() + Environment.NewLine);
                    }

                    if ((x.EIR > 0 && x.EIR < 1) == false)
                    {
                        sbEIR.Append(_parameter + "," + x.EIR.ToString() + Environment.NewLine);
                    }
                });

                if (sbDuration_Year.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "估計存續期間_年 Duration_Year 為不合格值 (不得空值且須大於0)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbDuration_Year.ToString();
                }

                if (sbRemaining_Month.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "估計存續期間_月 Remaining_Month 為不合格值 (不得空值且須大於0)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbRemaining_Month.ToString();
                }

                if (sbCurrent_Int_Rate.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "合约利率/產品利率 Current_Int_Rate 為不合格值 (必須 > 0 且 < 1)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Int_Rate.ToString();
                }

                if (sbEIR.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "有效利率 EIR 為不合格值 (必須 > 0 且 < 1)，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbEIR.ToString();
                }

                var _first = data.FirstOrDefault();
                List<IFRS9_Main> B01data = new List<IFRS9_Main>();
                List<Loan_IAS39_Info> A01data = new List<Loan_IAS39_Info>();
                List<Loan_IAS39_Info> A01North = new List<Loan_IAS39_Info>();
                List<Loan_IAS39_Info> A01NotNorth = new List<Loan_IAS39_Info>();
                List<EL_Data_Out> C07data = new List<EL_Data_Out>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (_first != null)
                    {
                        B01data = (from a in db.IFRS9_Main.AsNoTracking().ToList()
                                   join b in data on a.Reference_Nbr equals b.Reference_Nbr
                                   select a).ToList();

                        A01data = db.Loan_IAS39_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        A01North = A01data.Where(x => x.NO34RCV == "北區").ToList();

                        A01NotNorth = A01data.Where(x => x.NO34RCV == "非北區").ToList();

                        C07data = (from a in db.EL_Data_Out.AsNoTracking().ToList()
                                   join b in data on a.Reference_Nbr equals b.Reference_Nbr
                                   select a).ToList();
                    }
                }

                var C01North = (from a in data
                                join b in A01North on a.Reference_Nbr equals b.Reference_Nbr
                                select a).ToList();
                var C01NorthGroupData = C01North.GroupBy(x => new { x.Current_LGD })
                                                .Select(x => x.FirstOrDefault()).ToList();
                string NorthCurrentLgd = C01NorthGroupData.FirstOrDefault().Current_LGD.ToString();
                string NorthCurrentLgdCount = C01North.Where(x => x.Current_LGD.ToString() == NorthCurrentLgd).Count().ToString();

                var C01NotNorth = (from a in data
                                   join b in A01NotNorth on a.Reference_Nbr equals b.Reference_Nbr
                                   select a).ToList();
                var C01NotNorthGroupData = C01NotNorth.GroupBy(x => new { x.Current_LGD })
                                                      .Select(x => x.FirstOrDefault()).ToList();
                string NotNorthCurrentLgd = C01NotNorthGroupData.FirstOrDefault().Current_LGD.ToString();
                string NotNorthCurrentLgdCount = C01NotNorth.Where(x => x.Current_LGD.ToString() == NotNorthCurrentLgd).Count().ToString();

                var CurrentRatingCode5Data = data.Where(x => x.Current_Rating_Code == "5");
                var ImpairmentStage3Data = data.Where(x => x.Impairment_Stage == "3");
                var ImpairmentStage2Data = data.Where(x => x.Impairment_Stage == "2");
                var CurrentRatingCodeNot5Data1 = (from a in data.Where(x => x.Current_Rating_Code != "5")
                                                  join b in B01data.Where(x => x.Restructure_Ind == "Y"
                                                                            || x.Collateral_Legal_Action_Ind == "Y")
                                                  on a.Reference_Nbr equals b.Reference_Nbr
                                                  select a).ToList();
                var CurrentRatingCodeNot5Data2 = (from a in data.Where(x => x.Current_Rating_Code != "5")
                                                  join b in B01data.Where(x => x.Restructure_Ind == null 
                                                                            && x.Collateral_Legal_Action_Ind == null
                                                                            && x.Delinquent_Days > 29
                                                                            && x.Delinquent_Days < 90)
                                                  on a.Reference_Nbr equals b.Reference_Nbr
                                                  select a).ToList();

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("彙總資訊");

                sb.AppendLine($@"◆C01(EL_Data_In) 總資料筆數：{data.Count} 筆");

                sb.AppendLine("◆違約損失率 Current_LGD 值分布統計");
                var CurrentLgdCountQuery = from a in data
                                           group a by a.Current_LGD into grouping
                                           select new
                                           {
                                               grouping.Key,
                                               dataCount = grouping.Count()
                                           };
                foreach (var item in CurrentLgdCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆Current_Rating_Code 風險區隔");
                var CurrentRatingCodeCountQuery = from a in data
                                                  group a by a.Current_Rating_Code into grouping
                                                  select new
                                                  {
                                                      grouping.Key,
                                                      dataCount = grouping.Count()
                                                  };
                foreach (var item in CurrentRatingCodeCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆Impairment_Stage 分類資料值統計");
                var ImpairmentStageCountQuery = from a in data
                                                group a by a.Impairment_Stage into grouping
                                                select new
                                                {
                                                    grouping.Key,
                                                    dataCount = grouping.Count()
                                                };
                foreach (var item in ImpairmentStageCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("");

                sb.AppendLine($"◆A01(Loan_IAS39_Info)：{A01data.Count} 筆     C01(EL_Data_In)：{data.Count} 筆");
                sb.AppendLine("◆A01(Loan_IAS39_Info) 回收率區隔");
                sb.AppendLine($"◆北區：{A01North.Count} 筆     非北區：{A01NotNorth.Count} 筆");
                sb.AppendLine("◆C01(EL_Data_In) 違約損失率");
                sb.AppendLine($"{NorthCurrentLgd}：{NorthCurrentLgdCount} 筆     {NotNorthCurrentLgd}：{NotNorthCurrentLgdCount} 筆");

                if (ImpairmentStage3Data.Count() != CurrentRatingCode5Data.Count())
                {
                    sb.AppendLine("◆檢查風險區隔='5' 是否減損階段='3'");
                    sb.AppendLine("  減損階段=3 的筆數 與 風險區隔=5 的筆數不一致");

                    ImpairmentStage3Data.Where(x => x.Current_Rating_Code != "5").ToList()
                    .ForEach(x =>
                     {
                         sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Current_Rating_Code：{x.Current_Rating_Code},Impairment_Stage：{x.Impairment_Stage}");
                     });

                    CurrentRatingCode5Data.Where(x => x.Impairment_Stage != "3").ToList()
                    .ForEach(x =>
                    {
                        sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Current_Rating_Code：{x.Current_Rating_Code},Impairment_Stage：{x.Impairment_Stage}");
                    });
                }

                if (CurrentRatingCodeNot5Data1.Any(x => x.Impairment_Stage != "2"))
                {
                    sb.AppendLine("◆檢查 (Restructure_Ind='Y' 或 Collateral_Legal_Action_Ind='Y') 且 風險區隔 <>'5' 的資料 是否減損階段 ='2'");
                    sb.AppendLine("  減損階段=2 的筆數與規則條件判別不一致");
                    foreach (var x in CurrentRatingCodeNot5Data1.Where(x => x.Impairment_Stage != "2"))
                    {
                        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                        sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01?.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01?.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01?.Delinquent_Days},Current_Rating_Code：{oneB01?.Current_Rating_Code}");
                    }
                }

                if (CurrentRatingCodeNot5Data2.Any(x => x.Impairment_Stage != "2"))
                {
                    sb.AppendLine("◆檢查 (Delinquent_Days > 29 AND Delinquent_Days < 90) 且 風險區隔 <> '5' 的資料 是否減損階段='2'");
                    sb.AppendLine("  減損階段=2 的筆數與規則條件判別不一致");
                    foreach (var x in CurrentRatingCodeNot5Data2.Where(x => x.Impairment_Stage != "2"))
                    {
                        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                        sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01?.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01?.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01?.Delinquent_Days},Current_Rating_Code：{oneB01?.Current_Rating_Code}");
                    }
                }

                //if (ImpairmentStage2Data.Count() != CurrentRatingCodeNot5Data1.Count())
                    //{
                    //    sb.AppendLine("◆檢查 (Restructure_Ind='Y' 或 Collateral_Legal_Action_Ind='Y') 且 風險區隔 <>'5' 的資料 是否減損階段 ='2'");
                    //    sb.AppendLine("  減損階段=2 的筆數與規則條件判別不一致");

                //    ImpairmentStage2Data.ToList()
                    //    .ForEach(x =>
                    //    {
                    //        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                    //        if (oneB01 != null)
                    //        {
                    //            if (
                    //                 (
                    //                  (oneB01.Restructure_Ind == "Y" || oneB01.Collateral_Legal_Action_Ind == "Y") 
                    //                  && oneB01.Current_Rating_Code != "5"
                    //                 ) == false
                    //               )
                    //            {
                    //                sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01.Delinquent_Days},Current_Rating_Code：{oneB01.Current_Rating_Code}");
                    //            }
                    //        }
                    //    });

                //    CurrentRatingCodeNot5Data1.Where(x=>x.Impairment_Stage != "2").ToList()
                    //    .ForEach(x =>
                    //    {
                    //        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                    //        if (oneB01 != null)
                    //        {
                    //            sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01.Delinquent_Days},Current_Rating_Code：{oneB01.Current_Rating_Code}");
                    //        }
                    //    });
                    //}

                //if (ImpairmentStage2Data.Count() != CurrentRatingCodeNot5Data2.Count())
                //{
                //    sb.AppendLine("◆檢查 (Delinquent_Days > 29 AND Delinquent_Days < 90) 且 風險區隔 <> '5' 的資料 是否減損階段='2'");
                //    sb.AppendLine("  減損階段=2 的筆數與規則條件判別不一致");

                //    ImpairmentStage2Data.ToList()
                //    .ForEach(x =>
                //    {
                //        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                //        if (oneB01 != null)
                //        {
                //            if (
                //                 (
                //                  (oneB01.Delinquent_Days > 29 && oneB01.Delinquent_Days < 90)
                //                  && oneB01.Current_Rating_Code != "5"
                //                 ) == false
                //               )
                //            {
                //                sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01.Delinquent_Days},Current_Rating_Code：{oneB01.Current_Rating_Code}");
                //            }
                //        }
                //    });

                //    CurrentRatingCodeNot5Data2.Where(x => x.Impairment_Stage != "2").ToList()
                //    .ForEach(x =>
                //    {
                //        var oneB01 = B01data.Where(y => y.Reference_Nbr == x.Reference_Nbr).FirstOrDefault();
                //        if (oneB01 != null)
                //        {
                //            sb.AppendLine($"Reference_Nbr：{x.Reference_Nbr},Impairment_Stage：{x.Impairment_Stage},Restructure_Ind：{oneB01.Restructure_Ind},Collateral_Legal_Action_Ind：{oneB01.Collateral_Legal_Action_Ind},Delinquent_Days：{oneB01.Delinquent_Days},Current_Rating_Code：{oneB01.Current_Rating_Code}");
                //        }
                //    });
                //}

                sb.AppendLine("");

                var C07PFVData = C07data.GroupBy(x => new { x.PRJID, x.FLOWID, x.Version })
                                        .Select(x => x.FirstOrDefault()).ToList();
                foreach (var item in C07PFVData)
                {
                    sb.AppendLine($@"◆專案名稱：{item.PRJID}");
                    sb.AppendLine($@"  流程名稱：{item.FLOWID}");
                    sb.AppendLine($@"  資料版本：{item.Version}");
                }

                sb.AppendLine($@"◆C07(EL_Data_Out)總資料筆數：{C07data.Count} 筆");

                sb.AppendLine("◆Impairment_Stage 分類資料值統計");
                var C07ImpairmentStageCountQuery = from a in C07data
                                                   group a by a.Impairment_Stage into grouping
                                                   select new
                                                   {
                                                       grouping.Key,
                                                       dataCount = grouping.Count()
                                                   };
                foreach (var item in C07ImpairmentStageCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine($"◆C01(EL_Data_In)：{data.Count()} 筆     C07(EL_Data_Out)：{C07data.Count()} 筆");
                sb.AppendLine($"◆減損階段='1'     C01：{data.Where(x => x.Impairment_Stage == "1").Count()} 筆     C07：{C07data.Where(x => x.Impairment_Stage == "1").Count()} 筆");
                sb.AppendLine($"  減損階段='2'     C01：{data.Where(x => x.Impairment_Stage == "2").Count()} 筆     C07：{C07data.Where(x => x.Impairment_Stage == "2").Count()} 筆");
                sb.AppendLine($"  減損階段='3'     C01：{data.Where(x => x.Impairment_Stage == "3").Count()} 筆     C07：{C07data.Where(x => x.Impairment_Stage == "3").Count()} 筆");

                _customerStr_End += sb.ToString();
            }

            return result;
        }
        #endregion

        #region C02TransferCheck_Before
        private List<messageTable> C02TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                messageTable mt = new messageTable()
                {
                    title = @"檢查A02下面資料是否有空值資料",
                    successStr = @"來源資料A02內重要欄位皆非空值"
                };

                var data = _data.Cast<Loan_Account_Info>().ToList();

                var _first = data.OrderByDescending(x => x.Report_Date).FirstOrDefault();
                List<Loan_Account_Info> A02ThisReportDateData = new List<Loan_Account_Info>();
                if (_first != null)
                {
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        A02ThisReportDateData = db.Loan_Account_Info.AsNoTracking().Where(x => x.Report_Date == _first.Report_Date).ToList();
                    }
                }

                StringBuilder sbCurrent_Rating_Code = new StringBuilder();
                A02ThisReportDateData.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr}";

                    if (x.Current_Rating_Code == null || x.Current_Rating_Code == "")
                    {
                        sbCurrent_Rating_Code.Append(_parameter + Environment.NewLine);
                    }
                });

                if (sbCurrent_Rating_Code.Length > 0)
                {
                    _checkFlag = true;
                    _customerStr_End += "信用評等 Current_Rating_Code 有空值，有問題的帳號：" + Environment.NewLine;
                    _customerStr_End += sbCurrent_Rating_Code.ToString();
                }
            }

            return result;
        }
        #endregion

        #region C02dbModelCheck
        private List<messageTable> C02dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<Loan_Account_Info>().ToList();

                messageTable mt = new messageTable()
                {
                    title = @"彙總資訊",
                    successStr = @"轉檔完成"
                };

                var _first = data.OrderByDescending(x=>x.Report_Date).FirstOrDefault();
                DateTime lastReportDate = DateTime.Now.Date.AddDays(-DateTime.Now.Date.Day);
                List<Loan_Account_Info> A02ThisReportDateData = new List<Loan_Account_Info>();
                List<Loan_Account_Info> A02LastReportDateData = new List<Loan_Account_Info>();
                List<Rating_History> C02Data = new List<Rating_History>();
                List<Rating_History> C02ThisReportDateData = new List<Rating_History>();
                List<Rating_History> C02LastReportDateData = new List<Rating_History>();
                List<Loan_IAS39_Info> A01data = new List<Loan_IAS39_Info>();
                List<EL_Data_In> C01Data = new List<EL_Data_In>();
                List<Rating_History> C02ThisReportDateData_1 = new List<Rating_History>();
                List<Rating_History> C02ThisReportDateData_2 = new List<Rating_History>();
                List<Rating_History> C02ThisReportDateData_3 = new List<Rating_History>();
                List<Rating_History> C02ThisReportDateData_4 = new List<Rating_History>();
                List<Rating_History> C02ThisReportDateData_5 = new List<Rating_History>();
                if (_first != null)
                {
                    lastReportDate = _first.Report_Date.AddDays(-_first.Report_Date.Day);
                    A02ThisReportDateData = data.Where(x => x.Report_Date == _first.Report_Date).ToList();
                    A02LastReportDateData = data.Where(x => x.Report_Date == lastReportDate).ToList();
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        C02Data = (from a in db.Rating_History.AsNoTracking().ToList()
                                   join b in data on new { a.Reference_Nbr,a.Rating_Date } 
                                              equals new { b.Reference_Nbr,b.Rating_Date }
                                   select a).ToList();

                        C02ThisReportDateData = (from a in db.Rating_History.AsNoTracking().ToList()
                                                 join b in A02ThisReportDateData on new { a.Reference_Nbr, a.Rating_Date }
                                                                             equals new { b.Reference_Nbr, b.Rating_Date }
                                                 select a).ToList();

                        C02LastReportDateData = (from a in db.Rating_History.AsNoTracking().ToList()
                                                 join b in A02LastReportDateData on new { a.Reference_Nbr, a.Rating_Date }
                                                                             equals new { b.Reference_Nbr, b.Rating_Date }
                                                 select a).ToList();

                        A01data = db.Loan_IAS39_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        C01Data = db.EL_Data_In.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date).ToList();

                        C02ThisReportDateData_1 = (from a in C02ThisReportDateData
                                                   join b in C01Data.Where(x=>x.Current_Rating_Code == "1") 
                                                   on new { a.Reference_Nbr} equals new { b.Reference_Nbr}
                                                   select a).ToList();

                        C02ThisReportDateData_2 = (from a in C02ThisReportDateData
                                                   join b in C01Data.Where(x => x.Current_Rating_Code == "2")
                                                   on new { a.Reference_Nbr } equals new { b.Reference_Nbr }
                                                   select a).ToList();

                        C02ThisReportDateData_3 = (from a in C02ThisReportDateData
                                                   join b in C01Data.Where(x => x.Current_Rating_Code == "3")
                                                   on new { a.Reference_Nbr } equals new { b.Reference_Nbr }
                                                   select a).ToList();

                        C02ThisReportDateData_4 = (from a in C02ThisReportDateData
                                                   join b in C01Data.Where(x => x.Current_Rating_Code == "4")
                                                   on new { a.Reference_Nbr } equals new { b.Reference_Nbr }
                                                   select a).ToList();

                        C02ThisReportDateData_5 = (from a in C02ThisReportDateData
                                                   join b in C01Data.Where(x => x.Current_Rating_Code == "5")
                                                   on new { a.Reference_Nbr } equals new { b.Reference_Nbr }
                                                   select a).ToList();
                    }
                }

                StringBuilder sb = new StringBuilder();

                sb.AppendLine("◆評等日期 Rating_Date 資料筆數");
                var RatingDateCountQuery = from a in C02ThisReportDateData
                                           group a by a.Rating_Date into grouping
                                           select new
                                           {
                                               grouping.Key,
                                               dataCount = grouping.Count()
                                           };
                foreach (var item in RatingDateCountQuery)
                {
                    sb.AppendLine($"{item.Key}：{item.dataCount} 筆");
                }

                sb.AppendLine("◆信用評等 Current_Rating_Code 筆數統計");
                var LastReportDateCurrentRatingCodeCountQuery = from a in C02LastReportDateData
                                                                group a by a.Current_Rating_Code into grouping
                                                                select new
                                                                {
                                                                    grouping.Key,
                                                                    dataCount = grouping.Count()
                                                                };
                foreach (var item in LastReportDateCurrentRatingCodeCountQuery)
                {
                    sb.AppendLine($"{lastReportDate.ToString("yyyy/MM/dd")}     {item.Key}：{item.dataCount} 筆");
                }

                var ThisReportDateCurrentRatingCodeCountQuery = from a in C02ThisReportDateData
                                                                group a by a.Current_Rating_Code into grouping
                                                                select new
                                                                {
                                                                    grouping.Key,
                                                                    dataCount = grouping.Count()
                                                                };
                foreach (var item in ThisReportDateCurrentRatingCodeCountQuery)
                {
                    sb.AppendLine($"{A02ThisReportDateData.FirstOrDefault().Report_Date.ToString("yyyy/MM/dd")}     {item.Key}：{item.dataCount} 筆");
                }
                sb.AppendLine("");

                sb.AppendLine($"◆A01(Loan_IAS39_Info)：{A01data.Count} 筆     C02(Rating_History)：{C02ThisReportDateData.Count} 筆");

                sb.AppendLine("◆C01(EL_Data_In)與C02(Rating_History)的風險區隔值分布");
                sb.AppendLine($"風險區隔='1'  C01：{C01Data.Where(x => x.Current_Rating_Code == "1").Count()}筆    C02：{C02ThisReportDateData_1.Count()}筆");
                sb.AppendLine($"風險區隔='2'  C01：{C01Data.Where(x => x.Current_Rating_Code == "2").Count()}筆    C02：{C02ThisReportDateData_2.Count()}筆");
                sb.AppendLine($"風險區隔='3'  C01：{C01Data.Where(x => x.Current_Rating_Code == "3").Count()}筆    C02：{C02ThisReportDateData_3.Count()}筆");
                sb.AppendLine($"風險區隔='4'  C01：{C01Data.Where(x => x.Current_Rating_Code == "4").Count()}筆    C02：{C02ThisReportDateData_4.Count()}筆");
                sb.AppendLine($"風險區隔='5'  C01：{C01Data.Where(x => x.Current_Rating_Code == "5").Count()}筆    C02：{C02ThisReportDateData_5.Count()}筆");

                _customerStr_End = sb.ToString();
            }

            return result;
        }
        #endregion
    }
}