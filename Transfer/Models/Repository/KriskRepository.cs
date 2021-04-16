using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Text;
using Transfer.Models.Interface;
using Transfer.Utility;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class KriskRepository : IKriskRepository
    {
        private string C07reportDateFormate = "yyyy-MM-dd";

        #region Get Data

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

        public Tuple<string, string> getProductInfo(string productCode)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var data = db.Flow_Info.AsNoTracking().Where(x =>
                x.Group_Product_Code == productCode &&
                x.Apply_Off_Date == null).OrderByDescending(x=>x.Apply_Off_Date)
                .FirstOrDefault();
                if (data != null)
                {
                    return new Tuple<string, string>(data.PRJID, data.FLOWID);
                }
            }
            return new Tuple<string, string>("", "");
        }

        /// <summary>
        /// Get C01 Version
        /// </summary>
        /// <param name="product"></param>
        /// <param name="date"></param>
        /// <returns></returns>
        public List<string> GetC01Version (string product,DateTime date)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                List<string> result = new List<string>();
                var product_Code = db.Group_Product_Code_Mapping.AsNoTracking()
                    .Where(x => x.Group_Product_Code == product).Select(x => x.Product_Code).AsEnumerable();
                if (!product_Code.Any())
                    return result;
                var version = db.EL_Data_In.AsNoTracking()
                    .Where(x => x.Report_Date == date && product_Code.Contains(x.Product_Code))
                    .Select(x => x.Version).Distinct().OrderBy(x => x)
                    .AsEnumerable().Select(x=> x.ToString()).ToList();
                return version;
            }
        }

        public bool checkC07(DateTime reportDate,string PRJID,string FLOWID)
        {
            bool result = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _reportDate = reportDate.ToString(C07reportDateFormate);
                var _Loan_1 = Product_Code.M.GetDescription();
                result = db.EL_Data_Out.AsNoTracking()
                             .Where(x => x.Report_Date == _reportDate &&
                                x.Version == 1 &&
                                x.Product_Code == _Loan_1 &&
                                x.PRJID == PRJID &&
                                x.FLOWID == FLOWID).Any();
            }
            return result;
        }

        public void getBondC07CheckData(DateTime dt , int version)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var _dt = dt.ToString("yyyy-MM-dd");
                var Product_Codes = new List<string>()
                {
                    Product_Code.B_A.GetDescription(),
                    Product_Code.B_B.GetDescription(),
                    Product_Code.B_P.GetDescription()
                };
                var C07s = db.EL_Data_Out.AsNoTracking().
                    Where(x => x.Report_Date == _dt &&
                               x.Version == version &&
                               Product_Codes.Contains(x.Product_Code));
                new BondsCheckRepository<EL_Data_Out>(C07s, Check_Table_Type.Bonds_C07_Transfer_Check);
            }
        }

        #endregion Get Data

        #region Save Data

        /// <summary>
        /// 房貸執行 Krsik 前先刪除
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public Tuple<string,List<EL_Data_Out>> deleteC07(Flow_Apply_Status data)
        {
            string resilt = string.Empty;
            List<EL_Data_Out> cacheC07 = new List<EL_Data_Out>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var reportDate = data.Report_Date.ToString(C07reportDateFormate);
                var _Loan_1 = Product_Code.M.GetDescription();
                cacheC07 = db.EL_Data_Out.AsNoTracking()
                             .Where(x => x.Report_Date == reportDate &&
                                x.Version == 1 &&
                                x.Product_Code == _Loan_1 &&
                                x.PRJID == data.PRJID &&
                                x.FLOWID == data.FLOWID).ToList();

                string sql = $@"
delete EL_Data_Out
where Report_Date = '{reportDate}'
and Product_Code = '{_Loan_1}'
and Version = 1
and PRJID = '{data.PRJID}'
and FLOWID = '{data.FLOWID}'
";
                try
                {
                    db.Database.ExecuteSqlCommand(sql);
                }
                catch (Exception ex)
                {
                    resilt = ex.exceptionMessage();
                }              
            }
            return new Tuple<string, List<EL_Data_Out>>(resilt, cacheC07);
        }

        public void backC07(List<EL_Data_Out> datas)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                StringBuilder sb = new StringBuilder();
                datas.ForEach(
                    x => {
                        sb.Append(
                            $@"
INSERT INTO [EL_Data_Out]
           ([Report_Date]
           ,[Processing_Date]
           ,[Product_Code]
           ,[Reference_Nbr]
           ,[PD]
           ,[Lifetime_EL]
           ,[Y1_EL]
           ,[EL]
           ,[Impairment_Stage]
           ,[Version]
           ,[PRJID]
           ,[FLOWID])
     VALUES
           ('{x.Report_Date}'
           ,'{x.Processing_Date}'
           ,'{x.Product_Code}'
           ,'{x.Reference_Nbr}'
           ,{x.PD}
           ,{x.Lifetime_EL}
           ,{x.Y1_EL}
           ,{x.EL}
           ,'{x.Impairment_Stage}'
           ,{x.Version}
           ,'{x.PRJID}'
           ,'{x.FLOWID}') ; ");
                    });
                try
                {
                    if(sb.Length > 0)
                    db.Database.ExecuteSqlCommand(sb.ToString());
                }
                catch { }
            }
        }

        public void saveD04(Flow_Apply_Status data)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                db.Flow_Apply_Status.Add(data);
                try
                {
                    db.SaveChanges();
                }
                catch (Exception ex) {

                }
            }
        }
        #endregion
        
        #region 執行減損計算前檢核
        public MSGReturnModel CheckInfo(string date, string version)
        {
            DateTime dateReportDate = TypeTransfer.stringToDateTime(date);
            int intVersion = TypeTransfer.stringToInt(version);
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = true;
            try
            {
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    #region 檢核該版本是否已進入風控覆核流程
                    var versioninfo = db.Version_Info.Where(x => x.Report_Date == dateReportDate && x.Version == intVersion).FirstOrDefault();
                    if (versioninfo != null && versioninfo.Risk_Control_Status != 0)
                    {
                        result.RETURN_FLAG = false;
                        result.DESCRIPTION = Message_Type.kriskError.GetDescription(null, Message_Type.ReviewVersionInfo.GetDescription());
                        return result;
                    }
                    #endregion
                    #region C07轉檔前檢核數字一致(190628換券應收未收金額修正)
                    List<Bond_ISIN_Changed_IntRevise> A44_2List = new List<Bond_ISIN_Changed_IntRevise>().Where(x => x.Report_Date == dateReportDate && x.Version == intVersion).ToList();
                    foreach (var item in A44_2List)
                    {
                        var A41item = new List<Bond_Account_Info>()
                            .Where(x => x.Report_Date == dateReportDate && x.Version == intVersion && x.Bond_Number_Old == item.Bond_Number_Old && x.Portfolio_Name_Old == item.Portfolio_Name_Old && x.Lots == "1").FirstOrDefault();
                        var B01item = new List<IFRS9_Main>().Where(x => x.Report_Date == dateReportDate && x.Version == intVersion && x.Reference_Nbr == A41item.Reference_Nbr).FirstOrDefault();
                        var C01item = new List<EL_Data_In>().Where(x => x.Report_Date == dateReportDate && x.Version == intVersion && x.Reference_Nbr == A41item.Reference_Nbr).FirstOrDefault();

                        if (!(A41item.Interest_Receivable == B01item.Interest_Receivable) || !(A41item.Interest_Receivable == C01item.Interest_Receivable))//A41與B01、C01不一致狀況
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = Message_Type.kriskError.GetDescription(null, Message_Type.Not_Match_IntReceivable.GetDescription());                            
                            break;
                        }
                    }
                    return result;
                    #endregion
                }
            }
            catch (Exception ex)
            {
                result.DESCRIPTION = ex.Message.ToString();
                result.RETURN_FLAG = false;
                return result;
            }
        }
        #endregion



        protected class temp
        {
            public string Name { get; set; }
            public string Code { get; set; }
            public IEnumerable<string> Product_Code { get; set; }
        }
    }
}