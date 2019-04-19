﻿using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class C0Repository : IC0Repository
    {
        #region 其他
        protected Common common
        {
            get;
            private set;
        }
        public C0Repository()
        {
            this.common = new Common();
        }
        #endregion 其他

        #region getC01Version
        public List<string> getC01Version(string reportDate, string productCode)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.EL_Data_In.Any())
                {
                    var query = from q in db.EL_Data_In.AsNoTracking()
                                select q;

                    if (reportDate.IsNullOrWhiteSpace() == false)
                    {
                        DateTime tempDate = DateTime.Parse(reportDate);
                        query = query.Where(x => x.Report_Date == tempDate);
                    }

                    if (productCode.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Product_Code.Contains(productCode));
                    }

                    List<string> data = query.AsEnumerable().OrderBy(x => x.Version)
                                                            .Select(x => x.Version.ToString()).Distinct()
                                                            .ToList();
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region GetC01LogData
        public List<string> GetC01LogData(string tableTypes, string debt)
        {
            List<string> result = new List<string>();
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.IFRS9_Log.Any())
                    {
                        string[] arrayTableType = tableTypes.Split(',');
                        string tableType1 = arrayTableType[0];
                        string tableType2 = arrayTableType[1];

                        var items = db.Transfer_CheckTable.AsNoTracking()
                                      .Where(x => x.File_Name == tableType1
                                                  || x.File_Name == tableType2).ToList();
                        if (items.Any())
                        {
                            result.AddRange(
                                items.OrderByDescending(x=>x.Create_date).ThenByDescending(x => x.Create_time)
                                .Select(x =>
                                {
                                    return string.Format("{0}    轉檔日期：{1}    轉檔時間：{2}    基準日：{3}    版本：{4}    結果：{5}",
                                        x.File_Name,
                                        x.Create_date,
                                        x.Create_time,
                                        x.ReportDate.ToString("yyyy/MM/dd"),
                                        x.Version.ToString(),
                                        "Y".Equals(x.TransferType) ? "成功" : "失敗"
                                   );
                                })
                            );
                        }
                    }
                }
            }
            catch
            {
            }

            return result;
        }
        #endregion

        #region Get C01 Data

        /// <summary>
        /// get C01 data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<C01ViewModel>> getC01(string reportDate, string productCode, string version)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.EL_Data_In.Any())
                {
                    var query = from q in db.EL_Data_In.AsNoTracking()
                                select q;

                    if (reportDate.IsNullOrWhiteSpace() == false)
                    {
                        DateTime tempDate = DateTime.Parse(reportDate);
                        query = query.Where(x => x.Report_Date == tempDate);
                    }

                    if (productCode.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Product_Code.Contains(productCode));
                    }

                    if (version.IsNullOrWhiteSpace() == false)
                    {
                        query = query.Where(x => x.Version.ToString() == version);
                    }

                    return new Tuple<bool, List<C01ViewModel>>(query.Any(),
                        query.AsEnumerable().OrderBy(x => x.Report_Date).ThenBy(x => x.Reference_Nbr).ThenBy(x => x.Product_Code)
                        .Select(x => { return DbToC01ViewModel(x); }).ToList());
                }
            }

            return new Tuple<bool, List<C01ViewModel>>(false, new List<C01ViewModel>());
        }

        #endregion

        #region Db 組成 C01ViewModel
        private C01ViewModel DbToC01ViewModel(EL_Data_In data)
        {
            return new C01ViewModel()
            {
                Report_Date = data.Report_Date.ToString("yyyy/MM/dd"),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Product_Code = data.Product_Code,
                Reference_Nbr = data.Reference_Nbr,
                Current_Rating_Code = data.Current_Rating_Code,
                Exposure = data.Exposure.ToString(),
                Actual_Year_To_Maturity = data.Actual_Year_To_Maturity.ToString(),
                Duration_Year = data.Duration_Year.ToString(),
                Remaining_Month = data.Remaining_Month.ToString(),
                Current_LGD = data.Current_LGD.ToString(),
                Current_Int_Rate = data.Current_Int_Rate.ToString(),
                EIR = data.EIR.ToString(),
                Impairment_Stage = data.Impairment_Stage,
                Version = data.Version.ToString(),
                Lien_position = data.Lien_position,
                Ori_Amount = data.Ori_Amount.ToString(),
                Principal = data.Principal.ToString(),
                Interest_Receivable = data.Interest_Receivable.ToString(),
                Payment_Frequency = data.Payment_Frequency.ToString()
            };
        }
        #endregion

        #region getC01Excel
        public List<C01ViewModel> getC01Excel(string pathType, string path, string version)
        {
            List<C01ViewModel> dataModel = new List<C01ViewModel>();
            try
            {
                IWorkbook wb = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new HSSFWorkbook(stream);
                        }
                        break;

                    case "xlsx":
                        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
                        {
                            wb = new XSSFWorkbook(stream);
                        }
                        break;
                }

                ISheet sheet = wb.GetSheetAt(0);
                DataTable dt = sheet.ISheetToDataTable(true);
                if (dt.Rows.Count > 0) //判斷有無資料
                {
                    dataModel = dt.AsEnumerable()
                        .Select((x, y) =>
                        {
                            return getC01ViewModelInExcel(x,version);
                        }
                    ).ToList();
                }
            }
            catch (Exception ex)
            { }

            return dataModel;
        }

        #endregion

        #region Excel 組成 C01ViewModel

        private C01ViewModel getC01ViewModelInExcel(DataRow item,string version)
        {
            string reportDate = TypeTransfer.objToString(item[0]);
            reportDate = reportDate.Substring(0, 4) + "/" + reportDate.Substring(4, 2) + "/" + reportDate.Substring(6, 2);

            string productCode = TypeTransfer.objToString(item[2]);
            if (productCode.IndexOf('(') > 0)
            {
                productCode = productCode.Substring(0, productCode.IndexOf('('));
            }

            string exposure = TypeTransfer.objToString(item[5]);
            string actualYearToMaturity = TypeTransfer.objToString(item[6]);
            string durationYear = TypeTransfer.objToString(item[7]);
            string remainingMonth = TypeTransfer.objToString(item[8]);
            string currentLGD = TypeTransfer.objToString(item[9]);
            string currentIntRate = TypeTransfer.objToString(item[10]);
            string eir = TypeTransfer.objToString(item[11]);
            string oriAmount = TypeTransfer.objToString(item[15]);
            string principal = TypeTransfer.objToString(item[16]);
            string interestReceivable = TypeTransfer.objToString(item[17]);
            string paymentFrequency = TypeTransfer.objToString(item[18]);

            C01ViewModel data = new C01ViewModel();

            data.Report_Date = reportDate;
            data.Processing_Date = DateTime.Now.ToString("yyyy/MM/dd");
            data.Product_Code = productCode;
            data.Reference_Nbr = TypeTransfer.objToString(item[3]);
            data.Current_Rating_Code = TypeTransfer.objToString(item[4]);

            data.Exposure = null;
            if (exposure.IsNullOrWhiteSpace() == false)
            {
                data.Exposure = exposure;
            }

            data.Actual_Year_To_Maturity = null;
            if (actualYearToMaturity.IsNullOrWhiteSpace() == false)
            {
                data.Actual_Year_To_Maturity = actualYearToMaturity;
            }

            data.Duration_Year = null;
            if (durationYear.IsNullOrWhiteSpace() == false)
            {
                data.Duration_Year = durationYear;
            }

            data.Remaining_Month = null;
            if (remainingMonth.IsNullOrWhiteSpace() == false)
            {
                data.Remaining_Month = remainingMonth;
            }

            data.Current_LGD = null;
            if (currentLGD.IsNullOrWhiteSpace() == false)
            {
                data.Current_LGD = currentLGD;
            }

            data.Current_Int_Rate = null;
            if (currentIntRate.IsNullOrWhiteSpace() == false)
            {
                data.Current_Int_Rate = currentIntRate;
            }

            data.EIR = null;
            if (eir.IsNullOrWhiteSpace() == false)
            {
                data.EIR = eir;
            }

            data.Impairment_Stage = TypeTransfer.objToString(item[12]);

            data.Version = "1";
            if (version.IsNullOrWhiteSpace() == false)
            {
                data.Version = version;
            }

            data.Lien_position = TypeTransfer.objToString(item[14]);

            data.Ori_Amount = null;
            if (oriAmount.IsNullOrWhiteSpace() == false)
            {
                data.Ori_Amount = oriAmount;
            }

            data.Principal = null;
            if (principal.IsNullOrWhiteSpace() == false)
            {
                data.Principal = principal;
            }

            data.Interest_Receivable = null;
            if (interestReceivable.IsNullOrWhiteSpace() == false)
            {
                data.Interest_Receivable = interestReceivable;
            }

            data.Payment_Frequency = null;
            if (paymentFrequency.IsNullOrWhiteSpace() == false)
            {
                data.Payment_Frequency = paymentFrequency;
            }

            return data;
        }

        #endregion

        #region Save C01
        public MSGReturnModel saveC01(string country, string version, List<C01ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            string fileName = "C01_" + country;
            DateTime reportDate = DateTime.Parse(dataModel[0].Report_Date);
            int verInt = int.Parse(version);

            DateTime startTime = DateTime.Now;

            try
            {

                if (!dataModel.Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();
                    return result;
                }

                for (int i = 0; i < dataModel.Count; i++)
                {
                    if (dataModel[0].Report_Date != dataModel[i].Report_Date)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = "第 " + (i + 1).ToString() + " 筆資料錯誤：Report_Date 不一樣";
                        return result;
                    }
                }

                IEnumerable<EL_Data_In> edi = dataModel.Select(x =>
                                                                new EL_Data_In()
                                                                {
                                                                    Report_Date = DateTime.Parse(x.Report_Date),
                                                                    Processing_Date = DateTime.Parse(x.Processing_Date),
                                                                    Product_Code = x.Product_Code,
                                                                    Reference_Nbr = x.Reference_Nbr,
                                                                    Current_Rating_Code = x.Current_Rating_Code,
                                                                    Exposure = TypeTransfer.stringToDoubleN(x.Exposure),
                                                                    Actual_Year_To_Maturity = TypeTransfer.stringToDoubleN(x.Actual_Year_To_Maturity),
                                                                    Duration_Year = TypeTransfer.stringToDoubleN(x.Duration_Year),
                                                                    Remaining_Month = TypeTransfer.stringToIntN(x.Remaining_Month),
                                                                    Current_LGD = TypeTransfer.stringToDoubleN(x.Current_LGD),
                                                                    Current_Int_Rate = TypeTransfer.stringToDoubleN(x.Current_Int_Rate),
                                                                    EIR = TypeTransfer.stringToDoubleN(x.EIR),
                                                                    Impairment_Stage = x.Impairment_Stage,
                                                                    Version = int.Parse(x.Version),
                                                                    Lien_position = x.Lien_position,
                                                                    Ori_Amount = TypeTransfer.stringToDoubleN(x.Ori_Amount),
                                                                    Principal = TypeTransfer.stringToDoubleN(x.Principal),
                                                                    Interest_Receivable = TypeTransfer.stringToDoubleN(x.Interest_Receivable),
                                                                    Payment_Frequency = TypeTransfer.stringToIntN(x.Payment_Frequency)
                                                                }).AsEnumerable();

                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    if (db.Transfer_CheckTable.Any(x => x.File_Name == fileName 
                                                       && x.ReportDate == reportDate 
                                                       && x.Version == verInt 
                                                       && x.TransferType == "Y") == true)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = string.Format("{0}、{1}、版本：{2}，在 Transfer_CheckTable 已有轉檔成功的紀錄",
                                                           fileName,reportDate.ToString("yyyy/MM/dd"),version);
                        return result;
                    }

                    IEnumerable<EL_Data_In> dbEDI = db.EL_Data_In.AsEnumerable();
                    var existData = (
                                         from a in dbEDI
                                         join b in edi
                                         on new { a.Report_Date, a.Version, a.Reference_Nbr }
                                         equals new { b.Report_Date, b.Version, b.Reference_Nbr }
                                         select a
                                     ).ToList();

                    if (existData.Any())
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = string.Format("基準日：{0}，案件編號：{1}，版本：{2}，資料重複。", 
                                                           existData[0].Report_Date.ToString("yyyy/MM/dd"),
                                                           existData[1].Reference_Nbr,
                                                           existData[2].Version.ToString());
                        return result;
                    }

                    db.EL_Data_In.AddRange(edi);
                    db.SaveChanges();

                    result.RETURN_FLAG = common.saveTransferCheck(
                                                   fileName,
                                                   true,
                                                   reportDate,
                                                   verInt,
                                                   startTime,
                                                   DateTime.Now
                                         );

                    //add check by mark --2018 / 01 / 24
                    if (country == "HK")
                    {
                        new BondsCheckRepository<EL_Data_In>(edi, Check_Table_Type.Bonds_C01_HK_Transfer_Check);
                    }
                    else if (country == "VN")
                    {
                        new BondsCheckRepository<EL_Data_In>(edi, Check_Table_Type.Bonds_C01_VN_Transfer_Check);
                    }                 
                }
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = ex.Message;

                common.saveTransferCheck(
                        fileName,
                        false,
                        reportDate,
                        verInt,
                        startTime,
                        DateTime.Now,
                        ex.exceptionMessage()
                    );
            }

            return result;
        }
        #endregion

        #region getProduct
        public List<SelectOption> getProduct(GroupProductCode type)
        {
            List<SelectOption> options = new List<SelectOption>();
            string set = string.Empty;
            if (type == GroupProductCode.B)
            {
                set = GroupProductCode.B.GetDescription();
            }
            if (type == GroupProductCode.M)
            {
                set = GroupProductCode.M.GetDescription();
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                options.AddRange(
                db.Group_Product.AsNoTracking().Where(x =>
                x.Group_Product_Code.StartsWith(set))
                .GroupJoin(db.Group_Product_Code_Mapping.AsNoTracking(),
                   x => x.Group_Product_Code,
                   y => y.Group_Product_Code,
                   (x, y) =>
                   new temp()
                   {
                       Name = x.Group_Product_Name,
                       Code = x.Group_Product_Code,
                       Product_Code = y.Select(z => z.Product_Code)
                   }
                ).AsEnumerable().Select(x =>
                new SelectOption()
                {
                    Text = string.Format("{0} ({1})", x.Name, x.Code),
                    Value = string.Join(",", x.Product_Code)
                }));
            }
            return options;
        }
        #endregion

        #region temp
        protected class temp
        {
            public string Name { get; set; }
            public string Code { get; set; }
            public IEnumerable<string> Product_Code { get; set; }
        }
        #endregion

        #region getC07Version
        public List<string> getC07Version(string debtType, string productCode, string reportDate)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.EL_Data_Out.Any())
                {
                    var query = from q in db.EL_Data_Out.AsNoTracking()
                                select q;

                    if (productCode != "")
                    {
                        query = query.Where(x => productCode.Contains(x.Product_Code));
                    }

                    if (reportDate != "")
                    {
                        reportDate = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                        string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy/MM/dd");
                        query = query.Where(x => x.Report_Date == reportDate || x.Report_Date == reportDate2);
                    }

                    List<string> data = query.AsEnumerable().OrderBy(x=>x.Version)
                                                            .Select(x => x.Version.ToString()).Distinct()
                                                            .ToList();
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region Get C07 Data

        /// <summary>
        /// get C07 data
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<C07ViewModel>> getC07(string debtType, string groupProductCode, string productCode, string reportDate, string version)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.EL_Data_Out.Any())
                {
                    List<EL_Data_Out> query = db.EL_Data_Out.AsNoTracking().ToList();

                    if (productCode != "")
                    {
                        query = query.Where(x => productCode.Contains(x.Product_Code)).ToList();
                    }

                    if (reportDate != "")
                    {
                        string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                        query = query.Where(x => x.Report_Date == reportDate || x.Report_Date == reportDate2).ToList();
                    }

                    if (version != "")
                    {
                        query = query.Where(x => x.Version.ToString() == version).ToList();
                    }

                    List<Flow_Info> flowInfoList = db.Flow_Info.AsNoTracking()
                                                     .Where(x => x.Group_Product_Code == groupProductCode)
                                                     .Where(x => x.Apply_Off_Date == null)
                                                     .ToList();
                    query = (
                                from a in query
                                join b in flowInfoList
                                on new { a.PRJID, a.FLOWID }
                                equals new { b.PRJID, b.FLOWID }
                                select a
                            ).ToList();

                    List<C07ViewModel> C07s = query.OrderBy(x => x.Report_Date)
                                                   .ThenBy(x => x.Reference_Nbr)
                                                   .ThenBy(x => x.Product_Code)
                                                   .Select(x => { return DbToC07ViewModel(x); }).ToList();

                    if (debtType == "4")
                    {
                        DateTime dtReportDate = DateTime.Parse(reportDate);
                        List<Bond_Account_Info> A41s = db.Bond_Account_Info.Where(x => x.Report_Date == dtReportDate).ToList();

                        for (int i=0;i < C07s.Count;i++)
                        {
                            var referenceNbr = C07s[i].Reference_Nbr;
                            var A41 = A41s.Where(x => x.Reference_Nbr == referenceNbr).FirstOrDefault();
                            if (A41 != null)
                            {
                                C07s[i].Currency_Code = A41.Currency_Code;
                                C07s[i].Ex_rate = A41.Ex_rate.ToString();
                                C07s[i].Lifetime_EL_TWD = (double.Parse(C07s[i].Lifetime_EL) * A41.Ex_rate).ToString();
                                C07s[i].Y1_EL_TWD = (double.Parse(C07s[i].Y1_EL) * A41.Ex_rate).ToString();
                                C07s[i].EL_TWD = (double.Parse(C07s[i].EL) * A41.Ex_rate).ToString();
                            }
                        }
                    }

                    return new Tuple<bool, List<C07ViewModel>>(C07s.Any(), C07s);
                }
            }

            return new Tuple<bool, List<C07ViewModel>>(false, new List<C07ViewModel>());
        }

        #endregion Get C07 Data

        #region Db 組成 C07ViewModel

        /// <summary>
        /// Db 組成 C07ViewModel
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private C07ViewModel DbToC07ViewModel(EL_Data_Out data)
        {
            return new C07ViewModel()
            {
                Report_Date = data.Report_Date,
                Processing_Date = data.Processing_Date,
                Product_Code = data.Product_Code,
                Reference_Nbr = data.Reference_Nbr,
                PD = TypeTransfer.doubleNToString(data.PD),
                Lifetime_EL = TypeTransfer.doubleNToString(data.Lifetime_EL),
                Y1_EL = TypeTransfer.doubleNToString(data.Y1_EL),
                EL = TypeTransfer.doubleNToString(data.EL),
                Impairment_Stage = data.Impairment_Stage,
                Version = data.Version.ToString(),
                PRJID = data.PRJID,
                FLOWID = data.FLOWID
            };
        }

        #endregion Db 組成 C07ViewModel

        #region 下載 Excel

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(C07Mortgage.C07Bond)</param>
        /// <param name="path">下載位置</param>
        /// <param name="dbDatas"></param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel(string type, string path, List<C07ViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(type, Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = getC07ModelFromDb(type,dbDatas).Item1;

                if (Excel_DownloadName.C07Mortgage.ToString().Equals(type))
                {
                    result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C07Mortgage);
                    result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                }

                if (Excel_DownloadName.C07Bond.ToString().Equals(type))
                {
                    result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C07Bond);
                    result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                }

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription(type);
                }
            }

            return result;
        }

        #endregion 下載 Excel

        #region DB EL_Data_Out 組成 DataTable

        /// <summary>
        ///  DB EL_Data_Out 組成 DataTable
        /// </summary>
        /// <param name="dbDatas"></param>
        /// <returns></returns>
        private Tuple<DataTable> getC07ModelFromDb(string type,List<C07ViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                #region 組出資料

                dt.Columns.Add("評估基準日/報導日", typeof(object));
                dt.Columns.Add("資料處理日期", typeof(object));
                dt.Columns.Add("產品", typeof(object));
                dt.Columns.Add("案件編號/帳號", typeof(object));
                dt.Columns.Add("第一年違約率", typeof(object));
                if (type == "C07Bond")
                {
                    dt.Columns.Add("債券幣別", typeof(object));
                    dt.Columns.Add("基準日匯率", typeof(object));
                }
                dt.Columns.Add("存續期間預期信用損失(原幣)", typeof(object));
                dt.Columns.Add("一年期預期信用損失(原幣)", typeof(object));
                dt.Columns.Add("最終預期信用損失(原幣)", typeof(object));
                if (type == "C07Bond")
                {
                    dt.Columns.Add("存續期間預期信用損失(台幣)", typeof(object));
                    dt.Columns.Add("一年期預期信用損失(台幣)", typeof(object));
                    dt.Columns.Add("最終預期信用損失(台幣)", typeof(object));
                }
                dt.Columns.Add("減損階段", typeof(object));
                dt.Columns.Add("資料版本", typeof(object));
                dt.Columns.Add("專案名稱", typeof(object));
                dt.Columns.Add("流程名稱", typeof(object));

                foreach (C07ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();
                    nrow["評估基準日/報導日"] = item.Report_Date;
                    nrow["資料處理日期"] = item.Processing_Date;
                    nrow["產品"] = item.Product_Code;
                    nrow["案件編號/帳號"] = item.Reference_Nbr;
                    nrow["第一年違約率"] = item.PD;
                    if (type == "C07Bond")
                    {
                        nrow["債券幣別"] = item.Currency_Code;
                        nrow["基準日匯率"] = item.Ex_rate;
                    }
                    nrow["存續期間預期信用損失(原幣)"] = item.Lifetime_EL;
                    nrow["一年期預期信用損失(原幣)"] = item.Y1_EL;
                    nrow["最終預期信用損失(原幣)"] = item.EL;
                    if (type == "C07Bond")
                    {
                        nrow["存續期間預期信用損失(台幣)"] = item.Lifetime_EL_TWD;
                        nrow["一年期預期信用損失(台幣)"] = item.Y1_EL_TWD;
                        nrow["最終預期信用損失(台幣)"] = item.EL_TWD;
                    }
                    nrow["減損階段"] = item.Impairment_Stage;
                    nrow["資料版本"] = item.Version;
                    nrow["專案名稱"] = item.PRJID;
                    nrow["流程名稱"] = item.FLOWID;
                    dt.Rows.Add(nrow);
                }

                #endregion 組出資料
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }

        #endregion DB EL_Data_Out 組成 DataTable

        #region getA41AdvancedAssessment_Sub_Kind
        public List<string> getA41AdvancedAssessment_Sub_Kind()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Bond_Account_Info.Any())
                {
                    List<string> data = db.Bond_Account_Info.AsNoTracking()
                                            .Where(x => x.Assessment_Sub_Kind != null)
                                            .Where(x => x.Assessment_Sub_Kind.ToString() != "")
                                            .Select(x => x.Assessment_Sub_Kind)
                                            .Distinct()
                                            .ToList();                              
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region getC07Advanced
        public Tuple<string, List<C07AdvancedViewModel>> getC07Advanced(string groupProductCode, string productCode, string reportDate, string version, string assessmentSubKind, string watchIND)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                string _reportDate1 = string.Empty;
                if (db.EL_Data_Out.Any())
                {
                    List<EL_Data_Out> C07Data = db.EL_Data_Out.AsNoTracking().ToList();

                    if (productCode.IsNullOrWhiteSpace() == false)
                    {
                        C07Data = C07Data.Where(x => productCode.Contains(x.Product_Code)).ToList();
                    }

                    if (reportDate.IsNullOrWhiteSpace() == false)
                    {
                        string reportDate1 = DateTime.Parse(reportDate).ToString("yyyy-MM-dd");
                        _reportDate1 = reportDate1;
                        string reportDate2 = DateTime.Parse(reportDate).ToString("yyyy/MM/dd");
                        C07Data = C07Data.Where(x => x.Report_Date == reportDate1 || x.Report_Date == reportDate2).ToList();
                    }

                    if (version.IsNullOrWhiteSpace() == false)
                    {
                        C07Data = C07Data.Where(x => x.Version.ToString() == version).ToList();
                    }

                    DateTime dtReportDate = DateTime.Parse(reportDate);

                    List<Bond_Account_Info> A41Data = db.Bond_Account_Info.AsNoTracking()
                                                        .Where(x => x.Report_Date == dtReportDate
                                                               && x.Version.ToString() == version)
                                                        .ToList();

                    if (assessmentSubKind.IsNullOrWhiteSpace() == false)
                    {
                        A41Data = A41Data.Where(x => x.Assessment_Sub_Kind == assessmentSubKind).ToList();
                    }

                    List<Bond_Basic_Assessment> D62Data = db.Bond_Basic_Assessment.AsNoTracking()
                                                            .Where(x => x.Report_Date == dtReportDate
                                                                && x.Version.ToString() == version)
                                                            .ToList();

                    //if (watchIND.IsNullOrWhiteSpace() == false)
                    //{
                    //    D62Data = D62Data.Where(x => x.Watch_IND == watchIND).ToList();
                    //}

                    List<Group_Product_Code_Mapping> D05Data = db.Group_Product_Code_Mapping.AsNoTracking()
                                                                 .Where(x => x.Group_Product_Code == groupProductCode
                                                                          && productCode.Contains(x.Product_Code))
                                                                 .ToList();
                    List<Flow_Info> D01Data = db.Flow_Info.AsNoTracking()
                                                .Where(x => x.Group_Product_Code == groupProductCode 
                                                         && x.Apply_Off_Date.ToString() == "")
                                                .ToList();
                    //D01Data = (
                    //              from a in D01Data
                    //              join b in D05Data
                    //              on new { a.Group_Product_Code }
                    //              equals new { b.Group_Product_Code }
                    //              select a
                    //          ).ToList();

                    C07Data = (
                                  from a in C07Data
                                  join b in D01Data
                                  on new { a.PRJID, a.FLOWID }
                                  equals new { b.PRJID, b.FLOWID }
                                  select a
                              ).ToList();

                    if (C07Data.Any() == false)
                    {
                        return new Tuple<string, List<C07AdvancedViewModel>>("C07、D01 查無相關聯的資料", new List<C07AdvancedViewModel>());
                    }

                    C07Data = (
                                  from a in C07Data
                                  join b in A41Data
                                  on new { a.Reference_Nbr }
                                  equals new { b.Reference_Nbr }
                                  select a
                              ).ToList();

                    if (C07Data.Any() == false)
                    {
                        return new Tuple<string, List<C07AdvancedViewModel>>("C07、A41 查無相關聯的資料", new List<C07AdvancedViewModel>());
                    }

                    C07Data = (
                                  from a in C07Data
                                  join b in D62Data
                                  on new { a.Reference_Nbr }
                                  equals new { b.Reference_Nbr }
                                  into ps
                                  from b in ps.DefaultIfEmpty()
                                  select a
                              ).ToList();

                    if (C07Data.Any() == false)
                    {
                        return new Tuple<string, List<C07AdvancedViewModel>>("C07、D62 查無相關聯的資料", new List<C07AdvancedViewModel>());
                    }

                    var _first = C07Data.First();
                    var _PRJID = _first.PRJID;
                    var _FLOWID = _first.FLOWID;

                    var productCodes = new List<string>();
                    if (!productCode.IsNullOrWhiteSpace())
                        productCodes = productCode.Split(',').ToList();
                    var C09s = db.EL_Data_In_Update.AsNoTracking()
                        .Where(x => x.Report_Date == _reportDate1 &&
                                    x.PrjID == _PRJID &&
                                    x.FlowID == _FLOWID)
                        .Where(x => productCodes.Contains(x.Product_Code), productCodes.Any())
                        .ToList();

                    List<Group_Product> GroupProductData = db.Group_Product.ToList();
                    List<C07AdvancedViewModel> c07AdvancedList = new List<C07AdvancedViewModel>();
                    for (int i=0;i< C07Data.Count;i++)
                    {
                        string Product_Code = C07Data[i].Product_Code;
                        string Group_Product_Code = "";
                        string Group_Product_Name = "";
                        string Reference_Nbr = C07Data[i].Reference_Nbr;
                        string Assessment_Sub_Kind = "";
                        string Bond_Number = "";
                        string Lots = "";
                        string Portfolio = "";
                        string Ex_rate = "";
                        string Basic_Pass = "";
                        string Accumulation_Loss_This_Month = "";
                        string Chg_In_Spread_This_Month = "";
                        string Watch_IND = "";
                        string Exposure_EL = "";
                        string Exposure_Ex = "";
                        string PD = C07Data[i].PD.ToString();

                        Bond_Basic_Assessment oneD62 = D62Data.Where(x => x.Reference_Nbr == Reference_Nbr)
                                      .FirstOrDefault();
                        if (oneD62 != null)
                        {
                            Basic_Pass = oneD62.Basic_Pass;
                            Accumulation_Loss_This_Month = oneD62.Accumulation_Loss_This_Month?.ToString();
                            Chg_In_Spread_This_Month = oneD62.Chg_In_Spread_This_Month?.ToString();
                            Watch_IND = oneD62.Watch_IND;
                        }

                        Group_Product_Code_Mapping oneD05 = D05Data.Where(x => x.Product_Code == Product_Code)
                                                            .FirstOrDefault();
                        if (oneD05 != null)
                        {
                            Group_Product_Code = oneD05.Group_Product_Code;

                            Group_Product oneGroupProduct = GroupProductData.Where(x => x.Group_Product_Code == Group_Product_Code)
                                                                            .FirstOrDefault();
                            if (oneGroupProduct != null)
                            {
                                Group_Product_Name = oneGroupProduct.Group_Product_Name;
                            }
                        }

                        Bond_Account_Info oneA41 = A41Data.Where(x => x.Reference_Nbr == Reference_Nbr)
                                                          .FirstOrDefault();
                        if (oneA41 != null)
                        {
                            Assessment_Sub_Kind = oneA41.Assessment_Sub_Kind;
                            Bond_Number = oneA41.Bond_Number;
                            Lots = oneA41.Lots;
                            Portfolio = oneA41.Portfolio;
                            Ex_rate = oneA41.Ex_rate.ToString();
                            Exposure_EL =
                                (TypeTransfer.doubleNToDecimal(oneA41?.Principal) + TypeTransfer.doubleNToDecimal(oneA41.Interest_Receivable)).Normalize().ToString();
                            Exposure_Ex =
                                (TypeTransfer.doubleNToDecimal(oneA41?.Amort_Amt_Tw) + TypeTransfer.doubleNToDecimal(oneA41.Interest_Receivable_Tw)).Normalize().ToString();
                        }



                        var C09 = C09s.FirstOrDefault(x => x.Reference_Nbr == Reference_Nbr &&
                        x.Product_Code == Product_Code);

                        C07AdvancedViewModel c07Advanced = new C07AdvancedViewModel();
                        c07Advanced.Group_Product_Code = Group_Product_Code;
                        c07Advanced.Group_Product_Name = Group_Product_Name;
                        c07Advanced.Product_Code = Product_Code;
                        c07Advanced.Assessment_Sub_Kind = Assessment_Sub_Kind;
                        c07Advanced.Report_Date = C07Data[i].Report_Date;
                        c07Advanced.Version = C07Data[i].Version.ToString();
                        c07Advanced.Reference_Nbr = Reference_Nbr;
                        c07Advanced.Bond_Number = Bond_Number;
                        c07Advanced.Lots = Lots;
                        c07Advanced.Portfolio = Portfolio;
                        c07Advanced.PD = PD;
                        c07Advanced.LGD = TypeTransfer.doubleNToString(C09?.Current_LGD);
                        c07Advanced.Exposure_EL = Exposure_EL;
                        c07Advanced.Exposure_Ex = Exposure_Ex;
                        c07Advanced.Y1_EL = C07Data[i].Y1_EL.ToString();
                        c07Advanced.Lifetime_EL = C07Data[i].Lifetime_EL.ToString();
                        c07Advanced.Y1_EL_Ex = (C07Data[i].Y1_EL * TypeTransfer.doubleNToDouble(oneA41.Ex_rate)).ToString();
                        c07Advanced.Lifetime_EL_Ex = (C07Data[i].Lifetime_EL * TypeTransfer.doubleNToDouble(oneA41.Ex_rate)).ToString();
                        c07Advanced.Impairment_Stage = C07Data[i].Impairment_Stage;
                        c07Advanced.Ex_rate = Ex_rate;
                        if (!c07Advanced.Portfolio.IsNullOrWhiteSpace() &&
                            c07Advanced.Portfolio.IndexOf("FVPL") > -1)
                        {
                            c07Advanced.Basic_Pass = "不適用";
                            c07Advanced.Accumulation_Loss_This_Month = "不適用";
                            c07Advanced.Chg_In_Spread_This_Month = "不適用";
                            c07Advanced.Watch_IND = "不適用";
                        }
                        else
                        {
                            c07Advanced.Basic_Pass = Basic_Pass;
                            c07Advanced.Accumulation_Loss_This_Month = Accumulation_Loss_This_Month;
                            c07Advanced.Chg_In_Spread_This_Month = Chg_In_Spread_This_Month;
                            c07Advanced.Watch_IND = Watch_IND;
                        }
                        if (!watchIND.IsNullOrWhiteSpace())
                        {
                            if(c07Advanced.Watch_IND == watchIND)
                                c07AdvancedList.Add(c07Advanced);
                        }
                        else
                            c07AdvancedList.Add(c07Advanced);
                    }

                    return new Tuple<string, List<C07AdvancedViewModel>>("", c07AdvancedList);
                }
            }

            return new Tuple<string, List<C07AdvancedViewModel>>("查無資料", new List<C07AdvancedViewModel>());
        }

        #endregion

        #region DownloadC07AdvancedExcel
        public MSGReturnModel DownloadC07AdvancedExcel(string type, string path, List<C07AdvancedViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(type, Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = dbDatas.ToDataTable();

                if (Excel_DownloadName.C07Advanced.ToString().Equals(type))
                {
                    result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C07Advanced,new C07AdvancedViewModel().GetFormateTitles());
                    result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                }

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription(type);
                }
            }

            return result;
        }

        #endregion

        #region getC07AdvancedSum
        public Tuple<bool, List<C07AdvancedSumViewModel>> getC07AdvancedSum(string reportDate, string version, string groupProductCode, string groupProductName, string referenceNbr,string assessmentSubKind, string watchIND,string productCode)
        {
            List<C07AdvancedSumViewModel> c07AdvancedSumList = new List<C07AdvancedSumViewModel>();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<Bond_Account_Info> A41s = new List<Bond_Account_Info>();
                //List<EL_Data_In> C01s = new List<EL_Data_In>();
                List<EL_Data_Out> C07s = new List<EL_Data_Out>();
                var _reportDate = DateTime.Parse(reportDate);
                int _version = Int32.Parse(version);
                if (referenceNbr == "All")
                {
                    A41s = db.Bond_Account_Info.AsNoTracking().Where(x => x.Report_Date == _reportDate && x.Version == _version)
                        .Where(x=>x.Assessment_Sub_Kind == assessmentSubKind, !assessmentSubKind.IsNullOrWhiteSpace())
                        .ToList();
                    //C01s = db.EL_Data_In.AsNoTracking().Where(x => x.Report_Date == _reportDate && x.Version == _version).ToList();
                    C07s = db.EL_Data_Out.AsNoTracking().Where(x => x.Report_Date == reportDate && x.Version == _version).ToList();

                    if (!watchIND.IsNullOrWhiteSpace())
                    {
                        var refs = getC07Advanced(groupProductCode, productCode, reportDate, version, assessmentSubKind, watchIND).Item2.Select(x=>x.Reference_Nbr).ToList();
                        A41s = A41s.Where(x => refs.Contains(x.Reference_Nbr)).ToList();
                    }
                }
                else
                {
                    string[] arrayReferenceNbr = referenceNbr.Split(',');
                    string Reference_Nbr = "";

                    for (int i = 0; i < arrayReferenceNbr.Length; i++)
                    {
                        Reference_Nbr += "'" + arrayReferenceNbr[i] + "',";
                    }

                    Reference_Nbr = Reference_Nbr.Substring(0, Reference_Nbr.Length - 1);
                    A41s = db.Bond_Account_Info.AsNoTracking().Where(x => arrayReferenceNbr.Contains(x.Reference_Nbr)).ToList();
                    //C01s = db.EL_Data_In.AsNoTracking().Where(x => arrayReferenceNbr.Contains(x.Reference_Nbr)).ToList();
                    C07s = db.EL_Data_Out.AsNoTracking().Where(x => arrayReferenceNbr.Contains(x.Reference_Nbr)).ToList();
                }


                A41s
                    //.Join(C01s,
                    //A41 => A41.Reference_Nbr,
                    //C01 => C01.Reference_Nbr,
                    //(A41, C01) => new
                    //{
                    //    Reference_Nbr = A41.Reference_Nbr,
                    //    Assessment_Sub_Kind = A41.Assessment_Sub_Kind,
                    //    Ex_rate = TypeTransfer.doubleNToDecimal(A41.Ex_rate),
                    //    Exposure_EX = TypeTransfer.doubleNToDecimal(A41.Amort_Amt_Tw) + TypeTransfer.doubleNToDecimal(A41.Interest_Receivable_tw)
                    //})
                    .Join(C07s,
                    A41 => A41.Reference_Nbr,
                    C07 => C07.Reference_Nbr,
                    (A41, C07) => new 
                    {
                        Assessment_Sub_Kind = A41.Assessment_Sub_Kind,
                        Exposure_EX = TypeTransfer.doubleNToDecimal(A41.Amort_Amt_Tw) + TypeTransfer.doubleNToDecimal(A41.Interest_Receivable_Tw),
                        Y1_EL_EX = TypeTransfer.doubleNToDecimal(A41.Ex_rate) * TypeTransfer.doubleToDecimal(C07.Y1_EL),
                        Lifetime_EL_EX = TypeTransfer.doubleNToDecimal(A41.Ex_rate) * TypeTransfer.doubleToDecimal(C07.Lifetime_EL)
                    })
                    .GroupBy(x => x.Assessment_Sub_Kind).ToList()
                    .ForEach(x => {
                        c07AdvancedSumList.Add(
                            new C07AdvancedSumViewModel()
                            {
                                Report_Date = reportDate,
                                Version = version,
                                Group_Product_Code = groupProductCode,
                                Group_Product_Name = groupProductName,
                                Assessment_Sub_Kind = x.Key,
                                Exposure_EX = x.Sum(y=>y.Exposure_EX).Normalize().ToString(),
                                Y1_EL_EX = x.Sum(y => y.Y1_EL_EX).Normalize().ToString(),
                                Lifetime_EL_EX = x.Sum(y => y.Lifetime_EL_EX).Normalize().ToString()
                            });
                    });

                var _TotalExposure_EX = c07AdvancedSumList.Sum(x => TypeTransfer.stringToDecimal(x.Exposure_EX)).Normalize().ToString();
                var _TotalY1_EL_EX = c07AdvancedSumList.Sum(x => TypeTransfer.stringToDecimal(x.Y1_EL_EX)).Normalize().ToString();
                var _TotalLifetime_EL_EX = c07AdvancedSumList.Sum(x => TypeTransfer.stringToDecimal(x.Lifetime_EL_EX)).Normalize().ToString();

                c07AdvancedSumList.Add(new C07AdvancedSumViewModel()
                {
                    Report_Date = reportDate,
                    Version = version,
                    Group_Product_Code = groupProductCode,
                    Group_Product_Name = groupProductName,
                    Assessment_Sub_Kind = "總計",
                    Exposure_EX = _TotalExposure_EX,
                    Y1_EL_EX = _TotalY1_EL_EX,
                    Lifetime_EL_EX = _TotalLifetime_EL_EX
                });
            }

            return new Tuple<bool, List<C07AdvancedSumViewModel>>(c07AdvancedSumList.Any(), c07AdvancedSumList);
        }

        #endregion

        #region DownloadC07AdvancedSumExcel
        public MSGReturnModel DownloadC07AdvancedSumExcel(string type, string path, List<C07AdvancedSumViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                                 .GetDescription(type, Message_Type.not_Find_Any.GetDescription());

            if (dbDatas.Any())
            {
                DataTable datas = getC07AdvancedSumModelFromDb(dbDatas).Item1;

                if (Excel_DownloadName.C07AdvancedSum.ToString().Equals(type))
                {
                    //result.DESCRIPTION = FileRelated.DataTableToExcel(dbDatas.ToDataTable(), path, Excel_DownloadName.C07AdvancedSum, new C07AdvancedSumViewModel().GetFormateTitles());
                    result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.C07AdvancedSum);
                    result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
                }

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription(type);
                }
            }

            return result;
        }

        #endregion

        #region getC07AdvancedSumModelFromDb
        private Tuple<DataTable> getC07AdvancedSumModelFromDb(List<C07AdvancedSumViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                #region 組出資料

                dt.Columns.Add("報導日", typeof(object));
                dt.Columns.Add("版本", typeof(object));
                dt.Columns.Add("產品群代碼", typeof(object));
                dt.Columns.Add("產品群名稱", typeof(object));
                dt.Columns.Add("評估次分類", typeof(object));
                dt.Columns.Add("曝險額(台幣)", typeof(object));
                dt.Columns.Add("一年期預期信用損失(台幣)", typeof(object));
                dt.Columns.Add("存續期間預期信用損失(台幣)", typeof(object));

                foreach (C07AdvancedSumViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();
                    nrow["報導日"] = item.Report_Date;
                    nrow["版本"] = item.Version;
                    nrow["產品群代碼"] = item.Group_Product_Code;
                    nrow["產品群名稱"] = item.Group_Product_Name;
                    nrow["評估次分類"] = item.Assessment_Sub_Kind;
                    nrow["曝險額(台幣)"] = item.Exposure_EX;
                    nrow["一年期預期信用損失(台幣)"] = item.Y1_EL_EX;
                    nrow["存續期間預期信用損失(台幣)"] = item.Lifetime_EL_EX;
                    dt.Rows.Add(nrow);
                }

                #endregion 組出資料
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }

        #endregion

        #region Get Data

        #region getC04
        /// <summary>
        /// Get C04 View Data 
        /// </summary>
        /// <param name="ds">起始季別</param>
        /// <param name="de">截止季別</param>
        /// <param name="lastFlag">僅顯示最近更新資料</param>
        /// <returns></returns>
        public List<C04ViewModel> GetC04(string ds, string de,bool lastFlag = false)
        {
            List<C04ViewModel> result = new List<C04ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var C04ViewPros = new C04ViewModel().GetType().GetProperties().ToList(); 
                var C04Pros = new Econ_F_YYYYMMDD().GetType().GetProperties().ToList();
                var _MAxProcessing_Date = string.Empty;
                if (lastFlag)
                    _MAxProcessing_Date = db.Econ_F_YYYYMMDD.AsNoTracking().Max(x => x.Processing_Date);
                foreach (Econ_F_YYYYMMDD item in db.Econ_F_YYYYMMDD.AsNoTracking()
                    .Where(x => x.Processing_Date == _MAxProcessing_Date, !_MAxProcessing_Date.IsNullOrWhiteSpace())
                    .Where(x => x.Year_Quartly.CompareTo(ds) >= 0, !ds.IsNullOrWhiteSpace())
                    .Where(x => x.Year_Quartly.CompareTo(de) <= 0, !de.IsNullOrWhiteSpace())
                    .OrderBy(x => x.Year_Quartly))
                {
                    var C04View = new C04ViewModel();
                    C04Pros.ForEach(y =>
                    {
                        var C04ViewPro = C04ViewPros.FirstOrDefault(z => z.Name.ToUpper() == y.Name.ToUpper());
                        if (C04ViewPro != null)
                        {
                            C04ViewPro.SetValue(C04View, TypeTransfer.objToString(y.GetValue(item)));
                        }
                    });
                    C04View.Product_Code = "債券";
                    C04View.Data_ID = "1";
                    result.Add(C04View);
                }
            }
            return result;
        }
        #endregion

        #region Get C04 數列資料的時間
        /// <summary>
        /// Get C04 數列資料的時間
        /// </summary>
        /// <returns></returns>
        public List<string> GetC04YearQuartly()
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Econ_F_YYYYMMDD.AsNoTracking()
                    .Where(x => x.Year_Quartly != null)
                    .Select(x => x.Year_Quartly).OrderBy(x => x).ToList();
            }
            return result;
        }
        #endregion

        #region getC04Pro
        public List<C04ProViewModel> GetC04Pro()
        {
            List<C04ProViewModel> result = new List<C04ProViewModel>();
            List<string> nosearchs = new List<string>() {
                "Processing_Date","Product_Code","Data_ID","Year_Quartly","PD_Quartly"
            };
            foreach (var pro in new C04ViewModel().GetType().GetProperties())
            {
                if (!nosearchs.Contains(pro.Name))
                {
                    var atts = pro.GetCustomAttributes(typeof(DescriptionAttribute), false);
                    result.Add(
                        new C04ProViewModel()
                        {
                            Table_Property = pro.Name,
                            Property_Name = atts == null ? string.Empty : ((DescriptionAttribute)atts[0]).Description
                        });
                }
            }
            return result;
        }
        #endregion

        #region TransferC04
        /// <summary>
        /// 
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public List<ExpandoObject> C04Transfer(string from, string to, List<C04ProViewModel> data)
        {
            List<ExpandoObject> result = new List<ExpandoObject>();
            var _from = TypeTransfer.stringToIntN(from);
            var _to = TypeTransfer.stringToIntN(to);
            List<C04ViewModel> _allData = GetC04(null,null,false);

            _allData.ForEach(x =>
            {
                dynamic d = new ExpandoObject();
                IDictionary<string, object> item = d as IDictionary<string, object>;
                item.Add("Processing_Date", x.Processing_Date);
                item.Add("Product_Code", "債券");
                item.Add("Data_ID", "1");
                item.Add("Year_Quartly", x.Year_Quartly);
                item.Add("PD_Quartly", x.PD_Quartly);
                foreach (var _data in data)
                {
                    if (_data.Original_Factor == "Y")
                    {
                        item.Add(_data.Table_Property, GetPeriods(_allData, x.Year_Quartly, _data.Table_Property, 0));              
                    }
                    if (_data.Derivative == "Y")
                    {
                        if (_from != null && _to != null && _from.Value <= _to.Value)
                        {
                            var _fv = _from.Value;
                            var _tv = _to.Value;
                            for (var i = _fv; _tv >= i; i++)
                            {
                                if(i != 0)
                                   item.Add($"{_data.Table_Property}_L{i}", GetPeriods(_allData, x.Year_Quartly, _data.Table_Property, -i));
                            }
                        }
                        else if (_from != null)
                        {
                            if(_from != 0)
                                item.Add($"{_data.Table_Property}_L{_from.Value}", GetPeriods(_allData, x.Year_Quartly, _data.Table_Property, -_from.Value));
                        }
                        else if (_to != null)
                        {
                            if (_to != 0)
                                item.Add($"{_data.Table_Property}_L{_to.Value}", GetPeriods(_allData, x.Year_Quartly, _data.Table_Property, -_to.Value));
                        }
                    }
                }
                result.Add(d);
            });
            return result;
        } 
        #endregion

        #endregion

        #region Save Db

        #endregion

        #region Excel
        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <param name="type">(C04_1)</param>
        /// <param name="path">下載位置</param>
        /// <param name="cache">cache 資料</param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel<T>(string type, string path, List<T> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(type, Message_Type.not_Find_Any.GetDescription());
            if (Excel_DownloadName.C04_1.ToString().Equals(type))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<C04ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.C04_1,new C04ViewModel().GetFormateTitles());
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (Excel_DownloadName.C04_Transfer.ToString().Equals(type))
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<System.Dynamic.ExpandoObject>().ToList().ToDataTable(), path, Excel_DownloadName.C04_Transfer);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }
        #endregion

        #region Private function
        private string GetPeriods(
            List<C04ViewModel> _allData, 
            string Year_Quartly,
            string Table_Property, 
            int Number_Period)
        {
            string result = string.Empty;
            var C04Pros = new C04ViewModel().GetType().GetProperties();
            var C04Pro = C04Pros.FirstOrDefault(x => x.Name == Table_Property);

            if (C04Pro != null)
            {
                var _YearQuartly = GetYearQuartly(Year_Quartly, Number_Period);
                var _data = _allData.FirstOrDefault(x => x.Year_Quartly == _YearQuartly);
                if (_data != null)
                {
                    result = (C04Pro.GetValue(_data))?.ToString();
                }
            }

            return result;
        }

        private string GetYearQuartly(string yearQuartly, int Number_Period)
        {
            string result = string.Empty;
            var _Seasons = new string[] { "Q1", "Q2", "Q3", "Q4" };
            var length = _Seasons.Length;
            if (yearQuartly.Length == 6)
            {
                int _year = 0;
                Int32.TryParse(yearQuartly.Substring(0, 4), out _year);
                var season = yearQuartly.Substring(4);
                var _yearNumber = Number_Period / 4;
                var _index = Array.IndexOf(_Seasons, season);
                var _number = Number_Period % 4;
                var index = _index + _number;
                if (index < 0)
                {
                    index = index + length;
                    _year -= 1;
                }
                if (index >= length)
                {
                    index = index - length;
                    _year += 1;
                }
                result = (_year + _yearNumber) + _Seasons[index];
            }
            return result;
        }
        #endregion
    }
}