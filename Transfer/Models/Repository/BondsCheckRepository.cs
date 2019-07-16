using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using System.Web;
using Transfer.Infrastructure;
using Transfer.Models.Abstract;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    /// <summary>
    /// 減損計算過程資料檢核條件_債券
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class BondsCheckRepository<T> : CheckDataAbstract<T> 
        where T : class
    {
        /// <summary>
        /// 債券 ProductCodes
        /// </summary>
        private List<string> _Product_Codes = new List<string>()
        {
            Product_Code.B_A.GetDescription(),
            Product_Code.B_B.GetDescription(),
            Product_Code.B_P.GetDescription()
        };

        /// <summary>
        /// ProductCode 需取代的文字
        /// </summary>
        private List<FormateTitle> ProductCodeTitles =
                    new List<FormateTitle>() {
                        new FormateTitle() {
                            OldTitle = Product_Code.B_A.GetDescription(),
                            NewTitle = $@"{Product_Code.B_A.GetDescription()}-每月本息均攤"},
                        new FormateTitle() {
                            OldTitle = Product_Code.B_B.GetDescription(),
                            NewTitle = $@"{Product_Code.B_B.GetDescription()}-一次到期還本"},
                        new FormateTitle() {
                            OldTitle = Product_Code.B_P.GetDescription(),
                            NewTitle = $@"{Product_Code.B_P.GetDescription()}-無到期日(永續債)"}
                    };

        /// <summary>
        /// 債券檢核
        /// </summary>
        /// <param name="data"></param>
        /// <param name="_event"></param>
        public BondsCheckRepository(IEnumerable<T> data, Check_Table_Type _event, DateTime? reportDate = null, int? version = null) : base(data, _event, reportDate, version)
        {

        }

        /// <summary>
        /// 設定字典資源
        /// </summary>
        protected override void Set()
        {
            _resources.Add(Check_Table_Type.Bonds_A41_UpLoad, A41ViewModelCheck);
            _resources.Add(Check_Table_Type.Bonds_A58_Before_Check, A58TransferCheck_Before);
            _resources.Add(Check_Table_Type.Bonds_A58_Transfer_Check, A58dbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_A59_Before_Check, A59TransferCheck_Before);
            _resources.Add(Check_Table_Type.Bonds_B01_Before_Check, B01TransferCheck_Before);
            _resources.Add(Check_Table_Type.Bonds_B01_Transfer_Check, B01dbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C01_Before_Check, C01TransferCheck_Before);
            _resources.Add(Check_Table_Type.Bonds_C01_Transfer_Check, C01dbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C01_HK_VN_UpLoad, C01ViewModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C01_HK_Transfer_Check, C01HK_VNdbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C01_VN_Transfer_Check, C01HK_VNdbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C07_Transfer_Check, C07dbModelCheck);
            _resources.Add(Check_Table_Type.Bonds_C10_UpLoad_Check, C10ViewModelCheck);
        }

        #region (一)資料上傳 
        private List<messageTable> A41ViewModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<A41ViewModel>().ToList();

                //(1)檢查是否有  到期日(#16Maturity_Date)小於報導日的情況
                messageTable A41_1 = new messageTable()
                {
                    title = @"A.檢查到期日小於報導日 (A41:Excel)",
                    successStr = @"資料內無到期日小於報導日的情況"
                };
                //(2)檢查下面欄位原始上傳資料是否有空值資料 
                messageTable A41_2 = new messageTable()
                {
                    title = @"B.檢查下面欄位原始上傳資料是否有空值資料 (A41:Excel)",
                    successStr = @"來源資料內重要欄位皆非空值"
                };
                //(3)檢查下面所列經程式處理取得的欄位內容是否有空值
                messageTable A41_3 = new messageTable()
                {
                    title = @"C.檢查下面所列經程式處理取得的欄位內容是否有空值",
                    successStr = @"金融商品分類資料轉換後無空"
                };

                data.ForEach(x =>
                {
                    var _parameter = $@"Bond_Number : {x.Bond_Number} , Lots : {x.Lots} , Portfolio_Name : {x.Portfolio_Name}";

                    //到期日小於報導日
                    setCheckMsg(A41_1, 
                        @"到期日小於報導日",
                        _parameter,
                        checkStringToDateTime(x.Maturity_Date) ||
                        TypeTransfer.stringToDateTime(x.Maturity_Date) <TypeTransfer.stringToDateTime(x.Report_Date));

                    //#13 帳列面額(原幣) Ori_Amount
                    setCheckMsg(A41_2,
                        @"#13 帳列面額(原幣)",
                        _parameter,
                        checkStringToDouble(x.Ori_Amount));

                    //#14 原始利率(票面利率) Current_Int_Rate
                    setCheckMsg(A41_2,
                        @"#14 原始利率(票面利率)",
                        _parameter,
                        checkStringToDouble(x.Current_Int_Rate));

                    //#16 到期日 Maturity_Date
                    setCheckMsg(A41_2,
                        @"#16 到期日 Maturity_Date",
                        _parameter,
                        checkStringToDateTime(x.Maturity_Date));

                    //#23 SMF PRODUCT
                    setCheckMsg(A41_2,
                        @"#23 SMF PRODUCT",
                        _parameter,
                        checkString(x.Product));

                    //#27 攤銷後之成本數(原幣) Principal
                    setCheckMsg(A41_2,
                        @"#27 攤銷後之成本數(原幣) Principal",
                        _parameter,
                        checkStringToDouble(x.Principal));

                    //#29 應收利息(原幣) Interest_Receivable
                    setCheckMsg(A41_2,
                        @"#29 應收利息(原幣) Interest_Receivable",
                        _parameter,
                        checkStringToDouble(x.Interest_Receivable));

                    //#34 原始有效利率 EIR
                    setCheckMsg(A41_2,
                        @"#34 原始有效利率 EIR",
                        _parameter,
                        checkStringToDouble(x.Eir));

                    //#35債券幣別 Currency_Code
                    setCheckMsg(A41_2,
                        @"#35債券幣別 Currency_Code",
                        _parameter,
                        checkString(x.Currency_Code));

                    //#41 報表日匯率 Ex_rate
                    setCheckMsg(A41_2,
                        @"#41 報表日匯率 Ex_rate",
                        _parameter,
                        checkStringToDouble(x.Ex_rate));

                    //#46 成本匯率 Ori_Ex_rate
                    setCheckMsg(A41_2,
                        @"#46 成本匯率 Ori_Ex_rate",
                        _parameter,
                        checkStringToDouble(x.Ori_Ex_rate));

                    //#57 市價(原幣) Market_Value_Ori
                    setCheckMsg(A41_2,
                        @"#57 市價(原幣) Market_Value_Ori",
                        _parameter,
                        checkStringToDouble(x.Market_Value_Ori));

                    //#60 Portfolio英文 Portfolio_Name
                    setCheckMsg(A41_2,
                        @"#60 Portfolio英文 Portfolio_Name",
                        _parameter,
                        checkString(x.Portfolio_Name));

                    //#52 金融商品分類 Asset_Type
                    setCheckMsg(A41_3,
                        @"#52 金融商品分類 Asset_Type",
                        _parameter,
                        checkString(x.Asset_Type));
                });
                result.Add(A41_1);
                result.Add(A41_2);


                var _bondNumber_Count = data.Select(x => x.Bond_Number).Distinct().Count();
                var reportDate = data.First();
                var _Report_Date = TypeTransfer.stringToDateTime(data.First().Report_Date);
                var _year = _Report_Date.Year;
                var _month = _Report_Date.Month;
                var _Origination_Date_Count =
                    data.Select(x => TypeTransfer.stringToDateTimeN(x.Origination_Date))
                    .Where(x => x != null && x.Value.Year == _year && x.Value.Month == _month).Count();
                
                //A41_過濾基本要件評估
                var _Assessment_Check_Count =
                   data.Select(x => x.Assessment_Check)
                   .Where(x => x != null && x.Equals("N")).Count();

                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"D.資料群組統計(內容分類筆數統計) (A41:Excel)");
                sb.AppendLine($@"1.總資料筆數 : {data.Count} 筆");
                sb.AppendLine($@"2.總債券資料數 : {_bondNumber_Count} 筆");
                sb.AppendLine($@"3.當月新購入的資料筆數, 債券數 : {_Origination_Date_Count} 筆");
                sb.AppendLine(@"4.#17現金流類型Principal_Payment_Method_Code  分類資料值統計");
                groupData(data.GroupBy(x => x.Principal_Payment_Method_Code).OrderBy(x => x.Key),
                    sb,
                    new List<FormateTitle>() {
                        new FormateTitle() { OldTitle = "01", NewTitle = "01-每月本息均攤"},
                        new FormateTitle() { OldTitle = "02", NewTitle = "02-一次到期還本"},
                        new FormateTitle() { OldTitle = "04", NewTitle = "04-無到期日(永續債)"}
                    });
                sb.AppendLine(@"5.#42債券擔保順位 Lien_position  分類資料值統計");
                groupData(data.GroupBy(x => x.Lien_position).OrderBy(x => x.Key), sb);
                sb.AppendLine(@"6.#51國內\國外Bond_Aera 分類資料值統計");
                groupData(data.GroupBy(x => x.Bond_Aera).OrderBy(x => x.Key), sb);
                sb.AppendLine(string.Empty);
                sb.AppendLine(@"E.特別債券資料顯示");
                var _ISIN_Changed_Infos = data.Where(x => x.ISIN_Changed_Ind == "Y").ToList();
                sb.AppendLine($@"1.#61 是否為換券  ISIN_Changed_Ind=’Y (By Lot)’資料筆數 : {_ISIN_Changed_Infos.Count} 筆");
                //foreach (var item in _ISIN_Changed_Infos)
                //{
                //    sb.AppendLine($@" Reference_Nbr:{item.Reference_Nbr} , Bond_Number:{item.Bond_Number} , Lots:{item.Lots} , Portfolio_Name:{item.Portfolio_Name} , Bond_Number_Old:{item.Bond_Number_Old} , Lots_Old:{item.Lots_Old} , Portfolio_Name_Old:{item.Portfolio_Name_Old} , Origination_Date_Old:{item.Origination_Date_Old}");
                //}
                //sb.AppendLine($@"2.#61 是否為換券  ISIN_Changed_Ind=’Y (By ISIN)’債券數 : {_ISIN_Changed_Infos.Select(x=>x.Bond_Number).Distinct().Count()} 筆");
                var changInd = 0;
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    foreach (var A44 in db.Bond_ISIN_Changed_Info.AsNoTracking())
                    {
                        if (A44.Change_Date.Year == _year && A44.Change_Date.Month == _month)
                            changInd += 1;
                    }
                }
                sb.AppendLine($@"2. 本月換券資料:共{changInd}筆");

                sb.AppendLine(string.Empty);
                sb.AppendLine($@"F.不進行基本要件評估債券數量:{_Assessment_Check_Count}");
                _customerStr_End = sb.ToString();

            }
            return result;
        }
        #endregion

        #region (二)執行信評轉檔

        /// <summary>
        /// 執行信評轉檔前檢核
        /// </summary>
        /// <returns></returns>
        private List<messageTable> A58TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<Bond_Rating_Info>().ToList();
                var _first = data.First();

                messageTable A57_1 = new messageTable()
                {
                    title = $@"A57 基準日:{_first.Report_Date.ToString("yyyy/MM/dd")} 版本:{_first.Version.ToString()} 評等種類:原始投資信評 是否有rating特殊值 (A57:Bond_Rating_Info)",
                    successStr = @"信評資料如常,未有異常信評內容"
                };
                data.GroupBy(x => new { x.Bond_Number, x.RTG_Bloomberg_Field, x.Rating, x.Rating_Type })
                    .ToList().ForEach(x=> {
                        var _parameter = $@"Bond_Number : {x.Key.Bond_Number} , Rating_Type : {x.Key.Rating_Type} , RTG_Bloomberg_Field : {x.Key.RTG_Bloomberg_Field} , Rating : {x.Key.Rating}";
                        setCheckMsg(A57_1,
                           @"信評內容有未處理到的特殊值",
                           _parameter,
                           true);
                });
                result.Add(A57_1);
                if (A57_1.data != null && A57_1.data.Values.Any())
                    _checkFlag = true;
            }
            return result;
        }

        /// <summary>
        /// 執行信評轉檔統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> A58dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<Bond_Rating_Summary>().ToList();
                var _first = data.First();

                var _RO = Rating_Type.A.GetDescription();
                var _RR = Rating_Type.B.GetDescription();
                var _ratingType_O = new List<Bond_Rating_Summary>();
                var _ratingType_R = new List<Bond_Rating_Summary>();
                data.Where(x => x.Rating_Type == _RO)
                    .GroupBy(x=>x.Reference_Nbr).ToList()
                    .ForEach(x => 
                    {
                        if (!x.Any(y => y.Grade_Adjust != null))
                            _ratingType_O.AddRange(x);
                    });
                data.Where(x => x.Rating_Type == _RR)
                    .GroupBy(x => x.Reference_Nbr).ToList()
                    .ForEach(x =>
                    {
                        if (!x.Any(y => y.Grade_Adjust != null))
                            _ratingType_R.AddRange(x);
                    });

                StringBuilder sb = new StringBuilder();
                //A.資料群組統計(內容分類筆數統計)
                sb.AppendLine(@"A.資料群組統計(內容分類筆數統計) (A58:Bond_Rating_Summary)");
                sb.AppendLine($@"1.完全缺乏原始投資信評的債券筆數統計 : {_ratingType_O.Count} 筆");
                if (_ratingType_O.Any())
                    groupData(_ratingType_O.GroupBy(x => x.Bond_Number).OrderBy(x => x.Key), sb);
                else
                    sb.AppendLine(@"所有債券均有原始投資信評資料");
                sb.AppendLine(string.Empty);
                sb.AppendLine($@"2.完全缺乏報導日最近信評的債券筆數統計 : {_ratingType_R.Count} 筆");
                if (_ratingType_R.Any())
                    groupData(_ratingType_R.GroupBy(x => x.Bond_Number).OrderBy(x => x.Key), sb);
                else
                    sb.AppendLine(@"所有債券均有報導日最近信評資料");
                sb.AppendLine(string.Empty);
                //sb.AppendLine($@"3.進行券種類與評估次分類 內容值的統計");
                //sb.AppendLine(@"債券種類 統計");
                //groupData(data.GroupBy(x => x.Bond_Type).OrderBy(x=>x.Key), sb);
                //sb.AppendLine(@"評估次分類 統計");
                //PS: A58 無評估次分類

                _customerStr_Start = sb.ToString();

                //(1)檢查是否有  到期日(#16Maturity_Date)小於報導日的情況
                messageTable A57_1 = new messageTable()
                {
                    title = @"B.是否有rating特殊值 (A57:Bond_Rating_Info)",
                    successStr = @"信評資料如常,未有異常信評內容"
                };
                var A57Data = new List<IGrouping<string,Bond_Rating_Info>>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    var A57Datas = db.Bond_Rating_Info.AsNoTracking()
                                     .Where(x => x.Report_Date == _first.Report_Date &&
                                                 x.Version == _first.Version);
                    A57Data = A57Datas.Where(x => x.ISIN_Changed_Ind == "Y").GroupBy(x => x.Reference_Nbr).ToList();
                    var A57s = A57Datas.Where(x => x.Rating != null && x.PD_Grade == null)
                                       .GroupBy(x => new {x.Bond_Number, x.RTG_Bloomberg_Field, x.Rating ,x.Rating_Type}).ToList();
                    A57s.ForEach(x =>
                    {
                        var _parameter = $@"Bond_Number : {x.Key.Bond_Number} , Rating_Type : {x.Key.Rating_Type} , RTG_Bloomberg_Field : {x.Key.RTG_Bloomberg_Field} , Rating : {x.Key.Rating}";
                        setCheckMsg(A57_1,
                            @"信評內容有未處理到的特殊值",
                            _parameter,
                            true);
                    });
                }
                result.Add(A57_1);
                var _A41Data = new List<Bond_Account_Info>();
                var A41Data = new List<Bond_Account_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    _A41Data = db.Bond_Account_Info.AsNoTracking()
                                .Where(x => x.Report_Date == _first.Report_Date &&
                                            x.Version == _first.Version).ToList();
                    A41Data = _A41Data.Where(x =>  x.ISIN_Changed_Ind == "Y").ToList();
                }
                var A58Data = data.Where(x => x.ISIN_Changed_Ind == "Y").GroupBy(x => x.Reference_Nbr).ToList();
                StringBuilder sb2 = new StringBuilder();
                sb2.AppendLine(@"C.來源資料與產出資料的比較");
                //sb2.AppendLine(@"1.相同版本A41的是否為換券ISIN_Changed_Ind='Y' 的資料數與A57,A58的ISIN_Changed_Ind='Y'資料數是否一致");
                //if (A41Data.Count == A57Data.Count &&  A41Data.Count == A58Data.Count)
                //{
                //    sb2.AppendLine(@" 相同版本A41與A57,A58換券資料筆數一致");
                //}
                //else
                //{
                //    sb2.AppendLine(@" 相同版本A41與A58(A57)換券資料筆數不一致");
                //    if (A41Data.Count != A57Data.Count)
                //        sb2.AppendLine($@" A41資料筆數: {A41Data.Count}筆  A57資料筆數: {A57Data.Count}筆");
                //    if (A41Data.Count != A58Data.Count)
                //        sb2.AppendLine($@" A41資料筆數: {A41Data.Count}筆  A58資料筆數: {A58Data.Count}筆");
                //}
                //sb2.AppendLine(@"2.相同版本A41的是否為換券ISIN_Changed_Ind='Y' 的購入日(Origination_Date & Origination_Date_Old)與A57,A58的債券購入(認列)日期是否一致");
                //bool _Origination_Date_Flag = true;
                //StringBuilder sb3 = new StringBuilder();
                //foreach (var item in A41Data)
                //{
                //    var _A57Data = A57Data.FirstOrDefault(x => x.Key == item.Reference_Nbr);
                //    if (_A57Data != null)
                //    {
                //        if (_A57Data.Any(x => x.Origination_Date != item.Origination_Date ||
                //                            x.Origination_Date_Old != item.Origination_Date_Old))
                //        {
                //            _Origination_Date_Flag = false;
                //            sb3.AppendLine($@" A41與A57 不一致  Reference_Nbr {item.Reference_Nbr}");
                //        }
                //    }
                //    else
                //    {
                //        _Origination_Date_Flag = false;
                //        sb3.AppendLine($@" A41 Reference_Nbr {item.Reference_Nbr} , A57 null");
                //    }
                //    var _A58Data = A58Data.FirstOrDefault(x => x.Key == item.Reference_Nbr);
                //    if (_A58Data != null)
                //    {
                //        if (_A58Data.Any(x => x.Origination_Date != item.Origination_Date ||
                //                              x.Origination_Date_Old != item.Origination_Date_Old))
                //        {
                //            _Origination_Date_Flag = false;
                //            sb3.AppendLine($@" A41與A58 不一致  Reference_Nbr {item.Reference_Nbr}");
                //        }
                //    }
                //    else
                //    {
                //        _Origination_Date_Flag = false;
                //        sb3.AppendLine($@" A41 Reference_Nbr {item.Reference_Nbr} , A58 null");
                //    }
                //}
                //if (!_Origination_Date_Flag)
                //{
                //    sb2.AppendLine(@" 不一致債券清單");
                //    sb2.AppendLine(sb3.ToString());
                //}
                //else
                //    sb2.AppendLine(@" 相同版本A57,A58換券資訊與A41比對結果一致");
                //sb2.AppendLine(@"3.相同版本A41的債券種類資料數與A57,A58的債券種類資料數是否一致");
                sb2.AppendLine(@"1.相同版本A41的債券種類資料數與A57,A58的債券種類資料數是否一致");
                var _bond_types = _A41Data.GroupBy(x => x.Bond_Type).OrderBy(x => x.Key).ToList();
                StringBuilder sb4 = new StringBuilder();
                sb4 = compare(
                    _bond_types,
                    data.GroupBy(x=>x.Reference_Nbr).Select(x=>x.First())
                    .GroupBy(x=>x.Bond_Type).ToList(),
                    sb4, "A41", "A58");
                if (sb4.Length > 0)
                {
                    sb2.AppendLine(@" 相同版本A41與A58(A57)債券種類筆數不一致");
                    sb2.Append(sb4);
                }
                else
                {
                    sb2.AppendLine(@" 相同版本A41與A58(A57)債券種類筆數一致");
                    groupData(_bond_types, sb2);
                    sb2.AppendLine($@" 總筆數({string.Join("+",_bond_types.Select(x=>x.Key))}) : {(_bond_types.Select(x=>x.Count())).Sum()}筆");
                }
                _customerStr_End = sb2.ToString();
                
            }
            return result;
        }
        #endregion

        #region (三)B01轉檔

        /// <summary>
        /// B01轉檔前檢核
        /// </summary>
        /// <returns></returns>
        private List<messageTable> B01TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                //(1)檢查A41下面資料是否有空值資料
                messageTable A41_1 = new messageTable()
                {
                    title = @"A.檢查A41下面資料是否有空值資料 (A41:Bond_Account_Info)",
                    successStr = @"來源資料A41內重要欄位皆非空值"
                };
                var data = _data.Cast<Bond_Account_Info>().ToList();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version}";
                    //#13 帳列面額(原幣) Ori_Amount
                    setCheckMsg(A41_1,
                        @"#13 帳列面額(原幣) Ori_Amount",
                        _parameter,
                        x.Ori_Amount == null);

                    //#14 原始利率(票面利率)  Current_Int_Rate
                    setCheckMsg(A41_1,
                        @"#14 原始利率(票面利率) Current_Int_Rate",
                        _parameter,
                        x.Current_Int_Rate == null);

                    //#16 到期日 Maturity_Date
                    setCheckMsg(A41_1,
                        @"#16 到期日 Maturity_Date",
                        _parameter,
                        x.Maturity_Date == null);

                    //#27 攤銷後之成本數(原幣) Principal
                    setCheckMsg(A41_1,
                        @"#27 攤銷後之成本數(原幣) Principal",
                        _parameter,
                        x.Principal == null);

                    //#29 應收利息(原幣) Interest_Receivable
                    setCheckMsg(A41_1,
                        @"#29 應收利息(原幣) Interest_Receivable",
                        _parameter,
                        x.Interest_Receivable == null);

                    //#34 原始有效利率 EIR
                    setCheckMsg(A41_1,
                        @"#34 原始有效利率 EIR",
                        _parameter,
                        x.EIR == null);
                });
                result.Add(A41_1);
                if (A41_1.data != null && A41_1.data.Values.Any())
                    _checkFlag = true;
            }
            return result;
        }

        /// <summary>
        /// B01轉檔後統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> B01dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                StringBuilder _sb = new StringBuilder();
                _sb.AppendLine(@"A.檢查A41下面資料是否有空值資料 (A41:Bond_Account_Info)");
                _sb.AppendLine(@"來源資料A41內重要欄位皆非空值");
                _customerStr_Start = _sb.ToString();
                var data = _data.Cast<IFRS9_Main>().ToList();
                //(1)檢查下面的資料是否有空值
                messageTable B01_1 = new messageTable()
                {
                    title = @"B.檢查資料是否有空值 (B01:IFRS9_Main)",
                    successStr = @"欄位皆非空值"
                };
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version}";
                    //#18信用評等Current_Rating_Code
                    setCheckMsg(B01_1,
                        @"#18信用評等Current_Rating_Code",
                        _parameter,
                        checkString(x.Current_Rating_Code));

                    //#28最近外部信評評等Current_External_Rating
                    setCheckMsg(B01_1,
                        @"#28最近外部信評評等Current_External_Rating",
                        _parameter,
                        checkString(x.Current_External_Rating));

                    //#45起始外部評等Original_External_Rating
                    setCheckMsg(B01_1,
                        @"#45起始外部評等Original_External_Rating",
                        _parameter,
                        checkString(x.Original_External_Rating));
                });
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"C.資料群組統計(內容分類筆數統計) (B01:IFRS9_Main)");
                sb.AppendLine($@"1.總資料筆數 : {data.Count} 筆");
                sb.AppendLine(@"2. #15產品Product_Code分類資料值統計");
                groupData(data.GroupBy(x => x.Product_Code).OrderBy(x => x.Key), sb, ProductCodeTitles);
                _customerStr_End = sb.ToString();
                result.Add(B01_1);
            }
            return result;
        }

        #endregion

        #region (四)C01轉檔

        /// <summary>
        /// C01轉檔前檢核
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C01TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                //(1)檢查來源B01下面欄位是否有空值資料
                messageTable B01_1 = new messageTable()
                {
                    title = @"A.檢查來源B01下面欄位是否有空值資料 (B01:IFRS9_Main)",
                    successStr = @"來源資料B01內重要欄位皆非空值"
                };

                var data = _data.Cast<IFRS9_Main>().ToList();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version}";
                    //#3產品Product_Code
                    setCheckMsg(B01_1,
                        @"#3產品Product_Code",
                        _parameter,
                        checkString(x.Product_Code));

                    //#5風險區隔Current_Rating_Code
                    setCheckMsg(B01_1,
                        @"#5風險區隔Current_Rating_Code",
                        _parameter,
                        checkString(x.Current_Rating_Code));

                    //#6曝險額Exposure (B01.Principal + B01.Interest_Receivable)
                    setCheckMsg(B01_1,
                        @"#6曝險額Exposure(B01.Principal + B01.Interest_Receivable)",
                        _parameter,
                        (x.Principal == null) || (x.Interest_Receivable == null));

                    //#11合约利率/產品利率Current_Int_Rate
                    setCheckMsg(B01_1,
                        @"#11合约利率/產品利率Current_Int_Rate",
                        _parameter,
                        x.Current_Int_Rate == null);

                    //#12有效利率EIR
                    setCheckMsg(B01_1,
                        @"#12有效利率EIR",
                        _parameter,
                        x.Eir == null);

                    //#16原始購買金額Ori_Amount
                    setCheckMsg(B01_1,
                        @"#16原始購買金額Ori_Amount",
                        _parameter,
                        x.Ori_Amount == null);

                    //#17金融資產餘額Principal
                    setCheckMsg(B01_1,
                        @"#17金融資產餘額Principal",
                        _parameter,
                        x.Principal == null);

                    //#18應收利息Interest_Receivable
                    setCheckMsg(B01_1,
                        @"#18應收利息Interest_Receivable",
                        _parameter,
                        x.Interest_Receivable == null);
                });
                result.Add(B01_1);
                if (B01_1.data != null && B01_1.data.Values.Any())
                    _checkFlag = true;
            }
            return result;
        }

        /// <summary>
        /// C01轉檔後統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C01dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                StringBuilder _sb = new StringBuilder();
                _sb.AppendLine(@"A.檢查來源B01下面欄位是否有空值資料 (B01:IFRS9_Main)");
                _sb.AppendLine(@"來源資料B01內重要欄位皆非空值");
                _customerStr_Start = _sb.ToString();
                //(1)轉檔完成檢查下面的欄位值是否為合格值
                messageTable C01_1 = new messageTable()
                {
                    title = @"B.檢查C01資料是否為合格值 (C01:EL_Data_In)",
                    successStr = @"欄位皆為合格值"
                };
                var data = _data.Cast<EL_Data_In>().ToList();
                var _first = data.FirstOrDefault();

                List<Bond_Account_Info> A41data = new List<Bond_Account_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (_first != null)
                    {
                        A41data = db.Bond_Account_Info.AsNoTracking()
                                    .Where(x => x.Report_Date == _first.Report_Date &&
                                                x.Version == _first.Version).ToList();
                    }
                }

                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version}";
                //#7合約到期年限Actual_Year_To_Maturity (不得空值)
                setCheckMsg(C01_1,
                        @"#7合約到期年限Actual_Year_To_Maturity (不得空值)",
                        _parameter,
                        x.Actual_Year_To_Maturity == null);
                //#8估計存續期間_年Duration_Year (不得空值且須大於0)
                setCheckMsg(C01_1,
                        @"#8估計存續期間_年Duration_Year (不得空值且須大於0)",
                        _parameter,
                        x.Duration_Year == null || !(x.Duration_Year.Value > 0));

                //#9 估計存續期間_月Remaining_Month(不得空值且須大於0)
                setCheckMsg(C01_1,
                        @"#9 估計存續期間_月Remaining_Month(不得空值且須大於0)",
                        _parameter,
                        x.Remaining_Month == null || !(x.Remaining_Month.Value > 0));

                //#11合约利率/產品利率Current_Int_Rate (必須 > 0 且< 1)
                setCheckMsg(C01_1,
                        @"#11合约利率/產品利率Current_Int_Rate (必須 > 0 且< 1)",
                        _parameter,
                        x.Current_Int_Rate == null ||
                        !(1 > x.Current_Int_Rate.Value && x.Current_Int_Rate.Value > 0));

                //#12有效利率EIR (必須 > 0 且< 1)
                setCheckMsg(C01_1,
                        @"#12有效利率EIR (必須 > 0 且< 1)",
                        _parameter,
                        x.EIR == null ||
                        !(1 > x.EIR.Value && x.EIR.Value > 0));
                });
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"C.來源資料與產出資料的比較");
                sb.AppendLine(@"1.相同版本 A41與C01總資料筆數是否一致");
                if (data.Count == A41data.Count)
                {
                    sb.AppendLine(@"  A41與C01的總資料筆數一致");
                }
                else
                {
                    sb.AppendLine($@"  A41:{A41data.Count}筆 C01:{data.Count}筆");
                }
                //sb.AppendLine(@"2.相同版本 A41與C01債券筆數是否一致");
                var _01 = A41data.Where(x => x.Principal_Payment_Method_Code == "01").Count();
                var _02 = A41data.Where(x => x.Principal_Payment_Method_Code == "02").Count();
                var _04 = A41data.Where(x => x.Principal_Payment_Method_Code == "04").Count();
                var _Bond_A = data.Where(x => x.Product_Code == Product_Code.B_A.GetDescription()).Count();
                var _Bond_B = data.Where(x => x.Product_Code == Product_Code.B_B.GetDescription()).Count();
                var _Bond_P = data.Where(x => x.Product_Code == Product_Code.B_P.GetDescription()).Count();
                sb.AppendLine(@"2.相同版本 A41 #17 Principal_Payment_Method_Code");
                sb.AppendLine(@"  內容值 '01'(每月本息均攤) 的筆數是否與 C01 #3 Product_Code內容值 'Bond_A'(每月本息均攤) 的筆數一致;");
                sb.AppendLine(@"  內容值 '02'(一次到期還本) 的筆數是否與 C01 #3 Product_Code內容值 'Bond_B'(一次到期還本) 的筆數一致;");
                sb.AppendLine(@"  內容值 '04'(無到期日-永續債) 的筆數是否與 C01 #3 Product_Code內容值 'Bond_P'(無到期日-永續債) 的筆數一致;");
                if (_01 == _Bond_A && _02 == _Bond_B && _04 == _Bond_P)
                {
                    sb.AppendLine(@"  A41的還本方式 與 C01的產品別一致");
                }
                else
                {
                    sb.AppendLine(@"  A41的還本方式 與 C01的產品別不一致");
                    if (_01 != _Bond_A)
                        sb.AppendLine($@"  產品別-Bond_A(每月本息均攤) : {_Bond_A}筆 還本方式-01(每月本息均攤) : {_01}筆");
                    if (_02 != _Bond_B)
                        sb.AppendLine($@"  產品別-Bond_B(一次到期還本) : {_Bond_B}筆 還本方式-02(一次到期還本) : {_02}筆");
                    if (_04 != _Bond_P)
                        sb.AppendLine($@"  產品別-Bond_P(無到期日-永續債) : {_Bond_P}筆 還本方式-04(無到期日-永續債) : {_04}筆");
                }
                var _A41Lien_position = A41data.GroupBy(x => x.Lien_position).ToList();
                var _C01Lien_position = data.GroupBy(x => x.Lien_position).ToList();
                StringBuilder sbLien_position = new StringBuilder();
                sb.AppendLine(@"3.相同版本 A41與C01債券擔保順位分類資料值統計是否一致");
                sbLien_position = compare(
                    _A41Lien_position,
                    _C01Lien_position,
                    sbLien_position, "A41", "C01");
                if (sbLien_position.Length > 0)
                {
                    sb.AppendLine(@" 相同版本A41與C01債券擔保順位分類筆數不一致");
                    sb.AppendLine(sbLien_position.ToString());
                }
                else
                {
                    sb.AppendLine(@" 相同版本A41與C01債券擔保順位分類筆數一致");
                    groupData(_A41Lien_position.OrderBy(x => x.Key), sb);
                }
                _customerStr_End = sb.ToString();
                result.Add(C01_1);
            }
            return result;
        }

        /// <summary>
        /// C01轉檔 (香港,越南) 上傳Excel
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C01ViewModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<C01ViewModel>().ToList();

                //(1)檢查下面欄位原始上傳資料是否有空值資料 
                messageTable C01_1 = new messageTable()
                {
                    title = @"A.檢查上傳資料是否有空值資料 (C01:Excel)",
                    successStr = @"來源資料內重要欄位皆非空值"
                };

                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version} ";

                    //#3產品Product_Code
                    setCheckMsg(C01_1,
                        @"#3產品Product_Code",
                        _parameter,
                        checkString(x.Product_Code));

                    //#5風險區隔Current_Rating_Code
                    setCheckMsg(C01_1,
                        @"#5風險區隔Current_Rating_Code",
                        _parameter,
                        checkString(x.Current_Rating_Code));

                    //#6曝險額Exposure
                    setCheckMsg(C01_1,
                        @"#6曝險額Exposure",
                        _parameter,
                        checkStringToDouble(x.Exposure));

                    //#7合約到期年限Actual_Year_To_Maturity (不得空值)
                    setCheckMsg(C01_1,
                        @"#7合約到期年限Actual_Year_To_Maturity (不得空值)",
                        _parameter,
                        checkStringToDouble(x.Actual_Year_To_Maturity));

                    //#8估計存續期間_年Duration_Year (不得空值且須大於0)
                    setCheckMsg(C01_1,
                        @"#8估計存續期間_年Duration_Year (不得空值且須大於0)",
                        _parameter,
                        checkStringToDouble(x.Duration_Year) ||
                        !(TypeTransfer.stringToDouble(x.Duration_Year) > 0));

                    //#9 估計存續期間_月Remaining_Month(不得空值且須大於0)
                    setCheckMsg(C01_1,
                        @"/#9 估計存續期間_月Remaining_Month(不得空值且須大於0)",
                        _parameter,
                        checkStringToDouble(x.Remaining_Month) ||
                        !(TypeTransfer.stringToDouble(x.Remaining_Month) > 0));

                    //#11合约利率/產品利率Current_Int_Rate
                    setCheckMsg(C01_1,
                        @"#11合约利率/產品利率Current_Int_Rate",
                        _parameter,
                        checkStringToDouble(x.Current_Int_Rate));

                    //#12有效利率EIR
                    setCheckMsg(C01_1,
                        @"#12有效利率EIR",
                        _parameter,
                        checkStringToDouble(x.EIR));

                    //#16原始購買金額Ori_Amount
                    setCheckMsg(C01_1,
                        @"#16原始購買金額Ori_Amount",
                        _parameter,
                        checkStringToDouble(x.Ori_Amount));

                    //#17金融資產餘額Principal
                    setCheckMsg(C01_1,
                        @"/#17金融資產餘額Principal",
                        _parameter,
                        checkStringToDouble(x.Principal));

                    //#18應收利息Interest_Receivable
                    setCheckMsg(C01_1,
                        @"#18應收利息Interest_Receivable",
                        _parameter,
                        checkStringToDouble(x.Interest_Receivable));
                });
                result.Add(C01_1);
            }
            return result;
        }

        /// <summary>
        /// C01 香港越南轉檔後統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C01HK_VNdbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<EL_Data_In>().ToList();

                StringBuilder _sb = new StringBuilder();
                _sb.AppendLine(@"A.檢查上傳資料是否有空值資料 (C01:Excel)");
                _sb.AppendLine(@"來源資料內重要欄位皆非空值");
                _customerStr_Start = _sb.ToString();
                List<string> Product_Codes = new List<string>();
                if ((_event & Check_Table_Type.Bonds_C01_HK_Transfer_Check) == Check_Table_Type.Bonds_C01_HK_Transfer_Check)
                    Product_Codes = new List<string>()
                    {
                        Product_Code.HK_A.GetDescription(),
                        Product_Code.HK_B.GetDescription(),
                        Product_Code.HK_P.GetDescription()
                    };
                if ((_event & Check_Table_Type.Bonds_C01_VN_Transfer_Check) == Check_Table_Type.Bonds_C01_VN_Transfer_Check)
                    Product_Codes = new List<string>()
                    {
                        Product_Code.VN_A.GetDescription(),
                        Product_Code.VN_B.GetDescription(),
                        Product_Code.VN_P.GetDescription()
                    };
                //(1)檢查下面欄位原始上傳資料是否有空值資料 
                messageTable C01_1 = new messageTable()
                {
                    title = @"B.檢查轉檔資料是否為合格值 (C01:Grade_Moody_Info)",
                    successStr = @"來源資料內重要欄位皆通過合格值檢核"
                };
                var _A51 = new List<Grade_Moody_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    _A51 = db.Grade_Moody_Info.AsNoTracking().Where(x => x.Status == "1").ToList();
                }
                var _Grade_Adjust = _A51
                    .Where(x=>x.Grade_Adjust != null)
                    .Select(x => x.Grade_Adjust.Value.ToString()).ToList();
                data.ForEach(x =>
                {
                    var _parameter = $@"Reference_Nbr : {x.Reference_Nbr} , Version : {x.Version} ";

                    //#3產品Product_Code  
                    //(檢查是否符合規範的產品別
                    //例如香港固定為Bond_HK_A, Bond_HK_B, Bond_HK_P;
                    //越南固定為Bond_VN_A, Bond_VN_B, Bond_VN_P;
                    setCheckMsg(C01_1,
                            @"#3產品Product_Code",
                            _parameter,
                            !Product_Codes.Contains(x.Product_Code));

                    //#5風險區隔Current_Rating_Code
                    //                檢查內容值是否為最新年度A51 - 信評主標尺對應檔_Moody之 Grade_Adjust
                    // 從A51.Grade_Adjust 找出最大值,即可知道資料合理值應落於1~最大值範圍內,)
                    setCheckMsg(C01_1,
                            @"#5風險區隔Current_Rating_Code",
                            _parameter,
                            !_Grade_Adjust.Contains(x.Current_Rating_Code));

                    //#6曝險額Exposure
                    //   檢查Exposure 是否等於(Principal + Interest_Receivable)
                    setCheckMsg(C01_1,
                            @"#6曝險額Exposure 是否等於(Principal + Interest_Receivable)",
                            _parameter,
                            x.Exposure == null ||
                            x.Principal == null ||
                            x.Interest_Receivable == null ||
                            (x.Exposure.Value != x.Principal.Value + x.Interest_Receivable.Value));

                    //#7合約到期年限Actual_Year_To_Maturity (不得空值)
                    setCheckMsg(C01_1,
                            @"#7合約到期年限Actual_Year_To_Maturity (不得空值)",
                            _parameter,
                            x.Actual_Year_To_Maturity == null);

                    //#8估計存續期間_年Duration_Year (不得空值且須大於0)
                    setCheckMsg(C01_1,
                            @"#8估計存續期間_年Duration_Year (不得空值且須大於0)",
                            _parameter,
                            x.Duration_Year == null ||
                            !(x.Duration_Year.Value > 0));

                    //#9 估計存續期間_月Remaining_Month(不得空值且須大於0)
                    setCheckMsg(C01_1,
                            @"#9 估計存續期間_月Remaining_Month(不得空值且須大於0)",
                            _parameter,
                            x.Remaining_Month == null ||
                            !(x.Remaining_Month.Value > 0));

                    //#11合约利率/產品利率Current_Int_Rate (必須 > 0 且< 1)
                    setCheckMsg(C01_1,
                            @"#11合约利率/產品利率Current_Int_Rate (必須 > 0 且< 1)",
                            _parameter,
                            x.Current_Int_Rate == null ||
                            !(1 > x.Current_Int_Rate.Value && x.Current_Int_Rate.Value > 0));

                    //#12有效利率EIR (必須 > 0 且< 1)
                    setCheckMsg(C01_1,
                            @"#12有效利率EIR (必須 > 0 且< 1)",
                            _parameter,
                            x.EIR == null ||
                            !(1 > x.EIR.Value && x.EIR.Value > 0));

                });
                result.Add(C01_1);
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"C.資料群組統計(內容分類筆數統計) (C01:EL_Data_In)");
                sb.AppendLine($@"1.總資料筆數 : {data.Count}");
                //sb.AppendLine($@"2.總債券資料數(Distinct Bond_Number) : {data.Count}");
                //C01 無Bond_Number
                sb.AppendLine($@"2.#3產品Product_Code分類資料值統計");
                groupData(data.GroupBy(x => x.Product_Code).OrderBy(x => x.Key), sb, ProductCodeTitles);
                sb.AppendLine(@"3.#15債券擔保順位 Lien_position分類資料值統計");
                groupData(data.GroupBy(x => x.Lien_position).OrderBy(x => x.Key), sb);
                _customerStr_End = sb.ToString();
            }
            return result;
        }
        #endregion

        #region (五)減損計算結果
        /// <summary>
        /// C07轉檔後統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C07dbModelCheck()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<EL_Data_Out>().ToList();
                var _first = data.First();
                DateTime _reportData = DateTime.MinValue;
                DateTime.TryParse(_first.Report_Date, out _reportData);
                var C09s = new List<EL_Data_In_Update>();
                var C01s = new List<EL_Data_In>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    C09s = db.EL_Data_In_Update.AsNoTracking()
                             .Where(x => x.Report_Date == _first.Report_Date &&
                                         x.Version == _first.Version &&
                                         x.FlowID == _first.FLOWID &&
                                         x.PrjID == _first.PRJID).ToList();
                    C01s = db.EL_Data_In.AsNoTracking()
                             .Where(x => x.Report_Date == _reportData &&
                                         x.Version == _first.Version &&
                                         _Product_Codes.Contains(x.Product_Code)).ToList();
                }
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"A.套用流程資訊 (C07:EL_Data_Out)");
                sb.AppendLine($@"1.專案名稱 : {_first.PRJID}");
                sb.AppendLine($@"2.流程名稱 : {_first.FLOWID}");
                sb.AppendLine($@"3.資料版本 : {_first.Version}");
                sb.AppendLine(string.Empty);
                sb.AppendLine(@"B.資料群組統計(內容分類筆數統計) (C07:EL_Data_Out)");
                sb.AppendLine($@"1.總資料筆數 : {data.Count}");
                //sb.AppendLine($@"2.總債券資料數(Distinct Bond_Number) : {data.Count}");
                //C07 無Bond_Number
                sb.AppendLine($@"2.#3產品Product_Code分類資料值統計");
                groupData(data.GroupBy(x => x.Product_Code).OrderBy(x => x.Key), sb, ProductCodeTitles);
                sb.AppendLine(@"3.#15債券擔保順位 Lien_position分類資料值統計 (C09)");
                groupData(C09s.GroupBy(x => x.Lien_position).OrderBy(x => x.Key), sb);
                sb.AppendLine(@"4.C09-減損計算輸入資料修正 之LGD值");
                groupData(C09s.GroupBy(x => x.Current_LGD).OrderBy(x => x.Key), sb);
                sb.AppendLine(string.Empty);
                sb.AppendLine(@"C.來源資料與產出資料的比較");
                sb.AppendLine(@"1.相同版本 C01與C07總資料筆數是否一致");
                if (data.Count != C01s.Count)
                {
                    sb.AppendLine(@"相同版本C01與C07資料總數不一致");
                    sb.AppendLine($@" C01: {C01s.Count}筆  C07: {data.Count}筆");
                }                
                else
                    sb.AppendLine(@"相同版本C01與C07資料總數一致");

                var _BA = Product_Code.B_A.GetDescription(); //Bond_A
                var _BB = Product_Code.B_B.GetDescription(); //Bond_B
                var _BP = Product_Code.B_P.GetDescription(); //Bond_P

                var C07_Bond_A = data.Where(x => x.Product_Code == _BA).Count();
                var C07_Bond_B = data.Where(x => x.Product_Code == _BB).Count();
                var C07_Bond_P = data.Where(x => x.Product_Code == _BP).Count();
                var C01_Bond_A = C01s.Where(x => x.Product_Code == _BA).Count();
                var C01_Bond_B = C01s.Where(x => x.Product_Code == _BB).Count();
                var C01_Bond_P = C01s.Where(x => x.Product_Code == _BP).Count();
                sb.AppendLine(@"2.相同版本 C01,C07 產品別(Product_Code)分類統計值是否一致");
                if (C01_Bond_A == C07_Bond_A && C01_Bond_B == C07_Bond_B && C01_Bond_P == C07_Bond_P)
                {
                    sb.AppendLine(@"  相同版本 C01 與 C07 產品別一致");
                }
                else
                {
                    sb.AppendLine(@"  相同版本 C01 與 C07 產品別不一致");
                    if (C01_Bond_A != C07_Bond_A)
                        sb.AppendLine($@" C01 產品別-Bond_A : {C01_Bond_A}筆  ,C07 產品別-Bond_A : {C07_Bond_A}筆 ");
                    if (C01_Bond_B != C07_Bond_B)
                        sb.AppendLine($@" C01 產品別-Bond_B : {C01_Bond_B}筆  ,C07 產品別-Bond_B : {C07_Bond_B}筆 ");
                    if (C01_Bond_P != C07_Bond_P)
                        sb.AppendLine($@" C01 產品別-Bond_P : {C01_Bond_P}筆  ,C07 產品別-Bond_P : {C07_Bond_P}筆 ");
                }
                _customerStr_End = sb.ToString();
            }
            return result;
        }
        #endregion

        #region (六)執行信評補登前檢核
        private List<messageTable> A59TransferCheck_Before()
        {
            List<messageTable> result = new List<messageTable>();
            if (_data.Any())
            {
                var data = _data.Cast<Bond_Rating_Info>().ToList();

                messageTable A59_1 = new messageTable()
                {
                    title = @"「轉檔內容有信評特殊值，資料未寫入，請修改後重新上傳。」
A.檢查A59上傳資料是否有異常評等資料 (A57:Rating)",
                    successStr = @"來源資料A59評等資料皆無異常"
                };
                foreach (var item in data.GroupBy(x => new { x.Bond_Number, x.RTG_Bloomberg_Field, x.Rating, x.Rating_Type }))
                {
                    var _parameter = $@"Bond_Number : {item.Key.Bond_Number} , Rating_Type : {item.Key.Rating_Type} , RTG_Bloomberg_Field : {item.Key.RTG_Bloomberg_Field} , Rating : {item.Key.Rating}";
                    //#13 帳列面額(原幣) Ori_Amount
                    setCheckMsg(A59_1,
                        @"信評內容有未處理到的特殊值",
                        _parameter,
                        true);
                }
                result.Add(A59_1);
                if(A59_1.data != null && A59_1.data.Values.Any())
                    _checkFlag = true;
            }
            return result;
        }
        #endregion

        #region privateFunction
        /// <summary>
        /// 比對兩邊Group資料是否一致
        /// </summary>
        /// <typeparam name="_Data1">比對資料類別</typeparam>
        /// <typeparam name="_Data2">比對資料類別</typeparam>
        /// <param name="Data1">比對資料1</param>
        /// <param name="Data2">比對資料2</param>
        /// <param name="sb">不一致時加入訊息位置</param>
        /// <param name="Data1TableName">比對資料1類別名稱</param>
        /// <param name="Data2TableName">比對資料2類別名稱</param>
        /// <param name="compareKey">兩邊比對的key值 預設依樣</param>
        /// <param name="title_1">必對不一致時取代訊息1</param>
        /// <param name="title_2">必對不一致時取代訊息2</param>
        /// <returns></returns>
        private StringBuilder compare<_Data1, _Data2>(
            List<IGrouping<string, _Data1>> Data1, 
            List<IGrouping<string, _Data2>> Data2,
            StringBuilder sb,
            string Data1TableName,
            string Data2TableName,
            Dictionary<string,string> compareKey = null,
            List<FormateTitle> title_1 = null,
            List<FormateTitle> title_2 = null
            )
        {
            bool differenceFlag = false;
            StringBuilder _result = new StringBuilder();
            if (sb == null)
                sb = new StringBuilder();
            List<string> _keys = new List<string>(); 
            foreach (var item in Data1)
            {
                var _compareKey = string.Empty;
                _compareKey = item.Key;
                if (compareKey != null && item.Key != null && compareKey.ContainsKey(item.Key))
                {
                    _compareKey = compareKey[item.Key];
                }
                var _Comparedata = Data2.FirstOrDefault(x => x.Key == _compareKey);
                if (_Comparedata != null)
                {
                    _keys.Add(_Comparedata.Key);
                    if (item.Count() != _Comparedata.Count())
                    {
                        differenceFlag = true;
                    }
                    _result.AppendLine($@" {Data1TableName}-{formateTitle(title_1,item.Key)} : 資料數 {item.Count()} 筆 , {Data2TableName}-{formateTitle(title_2, _Comparedata.Key)} : 資料數 {_Comparedata.Count()} 筆 ");
                }
                else
                {
                    differenceFlag = true;
                    _result.AppendLine($@" {Data1TableName}-{formateTitle(title_1, item.Key)} : 資料數 {item.Count()} 筆 ");
                }
            }
            foreach (var item in Data2.Where(x => !_keys.Contains(x.Key)))
            {
                differenceFlag = true;
                _result.AppendLine($@" {Data2TableName}-{formateTitle(title_2, item.Key)} : 資料數 {item.Count()} 筆 ");
            }
            if (differenceFlag)
                sb = _result;
            return sb;
        }
        #endregion


        #region (七)C10上傳
        /// <summary>
        /// C10ViewModel轉檔後統計
        /// </summary>
        /// <returns></returns>
        private List<messageTable> C10ViewModelCheck()
        {
            List<messageTable> result = new List<messageTable>();

            if (_data.Any())
            {
                var data = _data.Cast<C10ViewModel>().ToList();

                //(1)檢查是否有  到期日(#16Maturity_Date)小於報導日的情況
                messageTable C10ViewModel_1 = new messageTable()
                {
                    title = @"A.檢查到期日小於報導日 (C10:Excel)",
                    successStr = @"資料內無到期日小於報導日的情況"
                };
                //(2)檢查下面欄位原始上傳資料是否有空值資料 
                messageTable C10ViewModel_2 = new messageTable()
                {
                    title = @"B.檢查下面欄位原始上傳資料是否有空值資料 (C10:Excel)",
                    successStr = @"來源資料內重要欄位皆非空值"
                };

                //設一個字典 債券編號&Lots&Portfolio、 狀態
                Dictionary<string, C10DataType> C10Dictionary = new Dictionary<string, C10DataType>();

                foreach (var SingleC10Data in data)
                {
                    C10DataType SingleC10Type = C10CheckData(SingleC10Data);
                    string C10key = SingleC10Data.Bond_Number + SingleC10Data.Lots + SingleC10Data.Portfolio_Name;
                    if (C10Dictionary.ContainsKey(C10key))
                    {
                        if (SingleC10Type == C10DataType.Amort && C10Dictionary[C10key] == C10DataType.Interest)
                        {
                            C10Dictionary[C10key] = C10DataType.Amort_Interest;
                        }
                        else if (SingleC10Type == C10DataType.Interest && C10Dictionary[C10key] == C10DataType.Amort)
                        {
                            C10Dictionary[C10key] = C10DataType.Amort_Interest;
                        }
                    }
                    else
                        C10Dictionary.Add(C10key, C10CheckData(SingleC10Data));
                }

                data.ForEach(x =>
                {
                    string C10key = x.Bond_Number + x.Lots + x.Portfolio_Name;
                    var _parameter = $@"Bond_Number : {x.Bond_Number} , Lots : {x.Lots} ,  : Portfolio_Name : {x.Portfolio_Name}";

                    //到期日小於報導日
                    setCheckMsg(C10ViewModel_1,
                        @"到期日小於報導日",
                        _parameter,
                        checkStringToDateTime(x.Maturity_Date) ||
                        TypeTransfer.stringToDateTime(x.Maturity_Date) < TypeTransfer.stringToDateTime(x.Report_Date));

                    if (C10Dictionary[C10key] == C10DataType.Amort_Interest && C10CheckData(x) == C10DataType.Interest)
                    {
                    }
                    else
                    {
                        //#7 攤銷後之成本數(原幣) Amort_Amt_import 
                        setCheckMsg(C10ViewModel_2,
                            @"#7 金融資產餘額攤銷後之成本數(原幣) Amort_Amt_import",
                            _parameter,
                            checkStringToDouble(x.Amort_Amt_import)); //Amount_AMT_import

                        //#8 金融資產餘額(台幣)攤銷後之成本數(台幣) Amort_Amt_import_TW
                        setCheckMsg(C10ViewModel_2,
                            @"#8 金融資產餘額(台幣)攤銷後之成本數(台幣) Amort_Amt_import_TW",
                            _parameter,
                            checkStringToDouble(x.Amort_Amt_import_TW)); //Amort_Amt_import_TW
                    }
                    if (C10Dictionary[C10key] == C10DataType.Amort_Interest && C10CheckData(x) == C10DataType.Amort)
                    { }
                    else
                    {
                        //#9 應收利息(原幣) Interest_Receivable_import
                        setCheckMsg(C10ViewModel_2,
                            @"#9 應收利息(原幣) Interest_Receivable_import",
                            _parameter,
                            checkStringToDouble(x.Interest_Receivable_import)); //Interest_Receivable_import

                        //#10 應收利息(台幣) Interest_Receivable_import_TW
                        setCheckMsg(C10ViewModel_2,
                            @"#10 應收利息(台幣) Interest_Receivable_import_TW",
                            _parameter,
                            checkStringToDouble(x.Interest_Receivable_import_TW)); // Interest_Receivable_import_TW
                    }
                    //#34 最近一次評等PD PD_import
                    setCheckMsg(C10ViewModel_2,
                        @"#34 最近一次評等PD  PD_import",
                        _parameter,
                        checkStringToDouble(x.PD_import)); // PD_import

                    //#35 LGD LGD_import LGD
                    setCheckMsg(C10ViewModel_2,
                        @"#35 LGD LGD_import",
                        _parameter,
                        checkStringToDouble(x.LGD_import)); // LGD

                    //#47 EL_本金(原幣) EL_import_Principle
                    setCheckMsg(C10ViewModel_2,
                        @"#47 EL_本金(原幣)  EL_import_Principle",
                        _parameter,
                        checkStringToDouble(x.EL_import_Principle)); // EL_import_Principle

                    //#48 EL_利息(原幣) EL_import_Int
                    setCheckMsg(C10ViewModel_2,
                        @"#48 EL_利息(原幣)  EL_import_Int",
                        _parameter,
                        checkStringToDouble(x.EL_import_Int)); // EL_import_Int
                });
                result.Add(C10ViewModel_1);
                result.Add(C10ViewModel_2);

                var _bondNumber_Count = data.Select(x => x.Bond_Number).Distinct().Count();
                var _Report_Date = TypeTransfer.stringToDateTime(data.First().Report_Date);
                var _year = _Report_Date.Year;
                var _month = _Report_Date.Month;
                var _Origination_Date_Count =
                    data.Select(x => TypeTransfer.stringToDateTimeN(x.Origination_Date))
                    .Where(x => x != null && x.Value.Year == _year && x.Value.Month == _month).Count();


                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"C.補上傳要件評估資料比對(ISIN&Lots&Portfolio)");

                //A41_抓出要件評估項目
                List<Bond_Account_Info> A41data_Assessment_Check = new List<Bond_Account_Info>();
                //取出DB內同報導日之A41要件評估項目
                try
                {
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        if (_Report_Date != null)
                        {
                            var version = db.Bond_Account_Info.AsNoTracking()
                                         .Where(e => e.Report_Date == _Report_Date)
                                         .Select(e => e.Version).Max(); //取最大version

                            A41data_Assessment_Check = db.Bond_Account_Info.AsNoTracking()
                                        .Where(p => p.Report_Date == _Report_Date &&
                                        p.Assessment_Check == "N" &&
                                        p.Version == version
                                       ).ToList();
                        }
                    }

                    //A41未補齊之檢核///
                    if (A41data_Assessment_Check.Count() > 0)
                    {
                        //A41_data為比對完上傳資料後仍然未補齊之項目
                        var A41_data = A41data_Assessment_Check.Where(
                                       a => !data.Exists(t =>
                                       a.Bond_Number != null &&
                                       a.Lots != null &&
                                       a.Portfolio_Name != null &&
                                       (a.Bond_Number == t.Bond_Number) &&
                                       (a.Lots == t.Lots) &&
                                       (a.Portfolio_Name == t.Portfolio_Name)
                                       )).ToList();

                        sb.AppendLine($@"需要補上傳之資料總數:{A41data_Assessment_Check.Count()}");
                        //sb.AppendLine($@"本次上傳資料總數(Group By ISIN&Lots&Portfolio)：{A41data_Assessment_Check.Count()- A41_data.Count() }");
                        if (A41_data.Count() > 0)
                        {
                            sb.AppendLine($@"基本要件評估債券ISIN&Lots&Portfolio尚未補齊");
                            sb.AppendLine($@"尚有{A41_data.Count()}筆資料未上傳，以下為尚未補齊之資料");
                            if (A41_data.Count() > 0)
                            {
                                foreach (var item in A41_data)
                                {
                                    sb.AppendLine($@"Bond_Number:{item.Bond_Number} , Lots:{item.Lots} , Portfolio_Name:{item.Portfolio}");
                                }
                            }
                        }
                        else
                        {
                            sb.AppendLine($@"基本要件評估債券ISIN&Lots&Portfolio已補齊");
                        }
                    }
                    else
                    {
                        sb.AppendLine($@"沒有需要補上傳之債券項目");
                    }
                    sb.AppendLine(string.Empty);

                    //上傳檔案多補的檢核
                    var C10_data = data.Where(
                                   a => !A41data_Assessment_Check.Exists(t => (a.Bond_Number == t.Bond_Number) && (
                                   a.Lots == t.Lots) && (a.Portfolio_Name == t.Portfolio_Name))
                                   ).ToList();

                    if (C10_data.Count() > 0)
                    {
                        sb.AppendLine($@"本次未能對應之上傳資料數量: {C10_data.Count()}");
                        foreach (var item in C10_data)
                        {
                            sb.AppendLine($@"Bond_Number:{item.Bond_Number} , Lots:{item.Lots} , Portfolio_Name:{item.Portfolio}");
                        }
                    }
                }
                catch (DbUpdateException ex)
                {

                }
                catch (Exception ex)
                {

                }
                sb.AppendLine(string.Empty);
                sb.AppendLine(@"D.資料群組統計(內容分類筆數統計) (C10)");
                sb.AppendLine($@"1.總資料筆數 : {data.Count} 筆");
                sb.AppendLine($@"2.總債券資料數 : {_bondNumber_Count} 筆");
                sb.AppendLine($@"3.總資料筆數(Group By ISIN & Lots & Portfolio)： {C10Dictionary.Count}筆");

                _customerStr_End = sb.ToString();
            }
            return result;
        }
        #endregion

        #region C10本金利息判斷
        private C10DataType C10CheckData(C10ViewModel C10)
        {
            C10DataType c10;
            bool C10_Amort = false;
            bool C10_Interest = false;

            //檢查資料本金欄位是否為空值，若填寫不齊全則報錯
            if (!C10.Amort_Amt_import.IsNullOrZero() || !C10.Amort_Amt_import_TW.IsNullOrZero() || !C10.EL_import_Principle.IsNullOrZero())
            {
                C10_Amort = true;
            }

            //檢查資料利息欄位是否為空值，若填寫不齊全則報錯
            if (!C10.Interest_Receivable_import.IsNullOrZero() || !C10.Interest_Receivable_import_TW.IsNullOrZero() || !C10.EL_import_Int.IsNullOrZero())
            {
                C10_Interest = true;
            }

            if (C10_Amort && C10_Interest)
                c10 = C10DataType.Amort_Interest;
            else if (C10_Interest && (C10_Amort == false))
                c10 = C10DataType.Interest;
            else if (C10_Amort && (C10_Interest == false))
                c10 = C10DataType.Amort;
            else
                c10 = C10DataType.Error_Data;

            return c10;
        }
        #endregion
    }
}