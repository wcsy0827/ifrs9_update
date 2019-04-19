using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Transactions;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A8Repository : IA8Repository
    {
        #region 其他
        public A8Repository()
        {
            this.common = new Common();
            this._UserInfo = new Common.User();
        }

        protected Common common
        {
            get;
            private set;
        }

        protected Common.User _UserInfo
        {
            get;
            private set;
        }
        #endregion 其他

        #region Get Data

        #region get Moody_Monthly_PD_Info(A81)

        /// <summary>
        /// get Moody_Monthly_PD_Info(A81)
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<A81ViewModel>> GetA81()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Monthly_PD_Info.Any())
                {
                    List<A81ViewModel> data = (from item in db.Moody_Monthly_PD_Info.AsNoTracking()
                                               .AsEnumerable()
                                               select new A81ViewModel() //轉型 Datetime
                                               {
                                                   Trailing_12m_Ending =
                                                   item.Trailing_12m_Ending.HasValue ?
                                                   item.Trailing_12m_Ending.Value.ToString("yyyy/MM/dd") : string.Empty,
                                                   Actual_Allcorp = TypeTransfer.doubleNToString(item.Actual_Allcorp),
                                                   Actual_SG = TypeTransfer.doubleNToString(item.Actual_SG),
                                                   Baseline_forecast_Allcorp = TypeTransfer.doubleNToString(item.Baseline_forecast_Allcorp),
                                                   Baseline_forecast_SG = TypeTransfer.doubleNToString(item.Baseline_forecast_SG),
                                                   Pessimistic_Forecast_Allcorp = TypeTransfer.doubleNToString(item.Pessimistic_Forecast_Allcorp),
                                                   Pessimistic_Forecast_SG = TypeTransfer.doubleNToString(item.Pessimistic_Forecast_SG),
                                                   Data_Year = item.Data_Year
                                               }).ToList();
                    return new Tuple<bool, List<A81ViewModel>>(true, data);
                }
            }

            return new Tuple<bool, List<A81ViewModel>>(false, new List<A81ViewModel>());
        }

        #endregion get Moody_Monthly_PD_Info(A81)

        #region get Moody_Quartly_PD_Info(A82)

        /// <summary>
        /// get Moody_Quartly_PD_Info(A82)
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<Moody_Quartly_PD_Info>> GetA82()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Quartly_PD_Info.Any())
                {
                    return new Tuple<bool, List<Moody_Quartly_PD_Info>>
                        (true, db.Moody_Quartly_PD_Info.AsNoTracking().ToList());
                }
            }

            return new Tuple<bool, List<Moody_Quartly_PD_Info>>(false, new List<Moody_Quartly_PD_Info>());
        }

        #endregion get Moody_Quartly_PD_Info(A82)

        #region get Moody_Predit_PD_Info(A83)

        /// <summary>
        /// get Moody_Predit_PD_Info(A83)
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<Moody_Predit_PD_Info>> GetA83()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_Predit_PD_Info.Any())
                {
                    return new Tuple<bool, List<Moody_Predit_PD_Info>>
                        (true, db.Moody_Predit_PD_Info.AsNoTracking().ToList());
                }
            }
            return new Tuple<bool, List<Moody_Predit_PD_Info>>(false, new List<Moody_Predit_PD_Info>());
        }

        #endregion get Moody_Predit_PD_Info(A83)

        #endregion Get Data

        #region Save Db

        /// <summary>
        /// save A81.A82.A83
        /// </summary>
        /// <param name="type">(A81 or A82 or A83)</param>
        /// <param name="dataModel">Exhibit10Model</param>
        /// <returns></returns>
        public MSGReturnModel SaveA8(string type, List<Exhibit10Model> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            try
            {
                bool _flag = true;
                string _message = string.Empty;

                List<string> A8Type = new List<string>() {
                    Table_Type.A81.ToString(),
                    Table_Type.A82.ToString(),
                    Table_Type.A83.ToString()};
                if (!A8Type.Contains(type))
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    return result;
                }

                //using (IFRS9DBEntities db = new IFRS9DBEntities())
                //{
                #region save Moody_Monthly_PD_Info(A81)

                if (Table_Type.A81.ToString().Equals(type))
                {
                    int id = 1;
                    var A81s = new List<Moody_Monthly_PD_Info>();
                    foreach (var item in dataModel)
                    {
                        DateTime? dt = TypeTransfer.stringToDateTimeN(item.Trailing);
                        A81s.Add(
                            new Moody_Monthly_PD_Info()
                            {
                                Id = id,
                                Trailing_12m_Ending = dt,
                                Actual_Allcorp =
                                TypeTransfer.stringToDoubleN(item.Actual_Allcorp),
                                Baseline_forecast_Allcorp =
                                TypeTransfer.stringToDoubleN(item.Baseline_forecast_Allcorp),
                                Pessimistic_Forecast_Allcorp =
                                TypeTransfer.stringToDoubleN(item.Pessimistic_Forecast_Allcorp),
                                Actual_SG = TypeTransfer.stringToDoubleN(item.Actual_SG),
                                Baseline_forecast_SG = TypeTransfer.stringToDoubleN(item.Baseline_forecast_SG),
                                Pessimistic_Forecast_SG =
                                TypeTransfer.stringToDoubleN(item.Pessimistic_Forecast_SG),
                                Data_Year = (dt == null) ? string.Empty : ((DateTime)dt).Year.ToString()
                            });
                        id += 1;
                    }
                    if (A81s.Any())
                    {
                        using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 10, 0)))
                        {
                            IFRS9DBEntities db = null;
                            try
                            {
                                db = new IFRS9DBEntities();
                                db.Configuration.AutoDetectChangesEnabled = false;
                                if (db.Moody_Monthly_PD_Info.Any())
                                    db.Moody_Monthly_PD_Info.RemoveRange(
                                        db.Moody_Monthly_PD_Info); //資料全刪除
                                db.SaveChanges();
                                db.Dispose();
                                db = new IFRS9DBEntities();
                                db.Configuration.AutoDetectChangesEnabled = false;
                                int count = 0;
                                foreach (var A81 in A81s)
                                {
                                    ++count;
                                    db = Common.AddToContext(db, A81, count, 100, true);
                                }
                                db.SaveChanges();
                            }
                            catch(Exception ex)
                            {
                                _flag = false;
                                _message = ex.exceptionMessage();
                            }
                            finally
                            {
                                if (_flag)
                                    scope.Complete();
                                db.Dispose();
                            }                      
                        }
                    }
                }

                    #endregion save Moody_Monthly_PD_Info(A81)

                    #region save Moody_Quartly_PD_Info(A82)

                if (Table_Type.A82.ToString().Equals(type))
                {

                    int id = 1;
                    List<Moody_Quartly_PD_Info> A82s = new List<Moody_Quartly_PD_Info>();
                    List<Moody_Quartly_PD_Info> allData = new List<Moody_Quartly_PD_Info>();
                    List<int> months = new List<int>() { 3, 6, 9, 12 }; //只搜尋3.6.9.12 月份                        
                    foreach (var item in dataModel
                        .Where(x => !string.IsNullOrWhiteSpace(x.Actual_Allcorp) //要有Actual_Allcorp (排除今年)
                        && months.Contains(DateTime.Parse(x.Trailing).Month)) //只搜尋3.6.9.12 月份
                        .OrderByDescending(x => x.Trailing)) //排序=>日期大到小
                    {
                        DateTime dt = DateTime.Parse(item.Trailing);
                        string quartly = dt.Year.ToString();
                        switch (dt.Month) //判斷季別
                        {
                            case 3:
                                quartly += "Q1";
                                break;

                            case 6:
                                quartly += "Q2";
                                break;

                            case 9:
                                quartly += "Q3";
                                break;

                            case 12:
                                quartly += "Q4";
                                break;
                        }
                        double? actualAllcorp = null;
                        if (!string.IsNullOrWhiteSpace(item.Actual_Allcorp))
                            actualAllcorp = double.Parse(item.Actual_Allcorp);
                        A82s.Add(new Moody_Quartly_PD_Info()
                        {
                            Id = id,
                            Data_Year = dt.Year.ToString(),
                            Year_Quartly = quartly,
                            PD = actualAllcorp
                        });
                        id += 1;
                    }
                    if (A82s.Any())
                    {
                        using (TransactionScope scope = new TransactionScope(TransactionScopeOption.Required, new TimeSpan(0, 10, 0)))
                        {
                            IFRS9DBEntities db = null;
                            try
                            {
                                db = new IFRS9DBEntities();
                                db.Configuration.AutoDetectChangesEnabled = false;
                                if (db.Moody_Quartly_PD_Info.Any())
                                    db.Moody_Quartly_PD_Info.RemoveRange(
                                         db.Moody_Quartly_PD_Info);
                                db.SaveChanges();
                                db.Dispose();
                                db = new IFRS9DBEntities();
                                db.Configuration.AutoDetectChangesEnabled = false;
                                int count = 0;
                                foreach (var A82 in A82s)
                                {
                                    ++count;
                                    db = Common.AddToContext(db, A82, count, 100, true);
                                }
                                db.SaveChanges();
                                using (IFRS9DBEntities db2 = new IFRS9DBEntities())
                                {
                                    string sql = string.Empty;
                                    var C04s = db2.Econ_F_YYYYMMDD.ToList();
                                    int _count = 0;
                                    List<SqlParameter> sps = new List<SqlParameter>();
                                    foreach (var A82 in A82s)
                                    {
                                        ++_count;
                                        var C04 = C04s.FirstOrDefault(x => x.Year_Quartly == A82.Year_Quartly);
                                        if (C04 != null)
                                        {
                                            sql += $@"
update Econ_F_YYYYMMDD
set PD_Quartly = @PD_Quartly{_count},
    LastUpdate_User = 'System',
    LastUpdate_Date = @LastUpdate_Date,
    LastUpdate_Time = @LastUpdate_Time
where Year_Quartly = @Year_Quartly{_count};
";
                                            sps.Add(new SqlParameter($"PD_Quartly{_count}", A82.PD));
                                            sps.Add(new SqlParameter($"Year_Quartly{_count}", A82.Year_Quartly));
                                        }
                                    }
                                    sps.Add(new SqlParameter("LastUpdate_Date", _UserInfo._date));
                                    sps.Add(new SqlParameter("LastUpdate_Time", _UserInfo._time));
                                    if(sql.Length > 0)
                                    db.Database.ExecuteSqlCommand(sql, sps.ToArray());
                                }
                            }
                            catch(Exception ex)
                            {
                                _flag = false;
                                _message = ex.exceptionMessage();
                            }
                            finally
                            {
                                if(_flag)
                                    scope.Complete();
                                db.Dispose();
                            }                         
                        }
                    }
                }

                    #endregion save Moody_Quartly_PD_Info(A82)

                    #region save Moody_Predit_PD_Info(A83)

                if (Table_Type.A83.ToString().Equals(type))
                {
                    using (IFRS9DBEntities db = new IFRS9DBEntities())
                    {
                        if (db.Moody_Predit_PD_Info.Any())
                            db.Moody_Predit_PD_Info.RemoveRange(db.Moody_Predit_PD_Info);
                        List<Exhibit10Model> models = (from q in dataModel
                                                       where !string.IsNullOrWhiteSpace(q.Actual_Allcorp) && //排除掉今年
                                                       12.Equals(DateTime.Parse(q.Trailing).Month) //只取12月
                                                       select q).ToList();
                        string maxYear = models.Max(x => DateTime.Parse(x.Trailing)).Year.ToString(); //抓取最大年
                        string minYear = models.Min(x => DateTime.Parse(x.Trailing)).Year.ToString(); //抓取最小年

                        double? PD = null;
                        double PDValue = models.Sum(x => double.Parse(x.Actual_Allcorp)) / models.Count; //計算 PD
                        if (PDValue > 0)
                            PD = PDValue;
                        db.Moody_Predit_PD_Info.Add(new Moody_Predit_PD_Info()
                        {
                            Id = 1,
                            Data_Year = maxYear,
                            Period = minYear + "-" + maxYear,
                            PD_TYPE = PD_Type.Past_Year_AVG.ToString(),
                            PD = PD
                        });
                        var dtn = DateTime.Now.Year;
                        Exhibit10Model model =
                            dataModel.Where(x => dtn.Equals(DateTime.Parse(x.Trailing).Year)
                            && 12.Equals(DateTime.Parse(x.Trailing).Month)).FirstOrDefault(); //抓今年又是12月的資料
                        string baselineForecastAllcorp = string.Empty;
                        if (model != null)
                            baselineForecastAllcorp = model.Baseline_forecast_Allcorp;
                        PD = null;
                        if (!string.IsNullOrWhiteSpace(baselineForecastAllcorp))
                            PD = double.Parse(baselineForecastAllcorp);

                        db.Moody_Predit_PD_Info.Add(new Moody_Predit_PD_Info()
                        {
                            Id = 2,
                            Data_Year = maxYear,
                            Period = dtn.ToString(),
                            PD_TYPE = PD_Type.Forcast.ToString(),
                            PD = PD
                        });
                        db.SaveChanges();
                    }

                }

                #endregion save Moody_Predit_PD_Info(A83)
                if (_flag)
                {
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription(type);
                }
                else
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = _message;
                }
            }
            catch (DbUpdateException ex)
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type
                                     .save_Fail.GetDescription(type,
                                     $"message: {ex.Message}" +
                                     $", inner message {ex.InnerException?.InnerException?.Message}");
            }
            return result;
        }

        #endregion Save Db

        #region Excel 部分

        /// <summary>
        /// 把Excel 資料轉換成 Exhibit10Model
        /// </summary>
        /// <param name="pathType">string</param>
        /// <param name="stream">Stream</param>
        /// <returns>Exhibit10Models</returns>
        public List<Exhibit10Model> getExcel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<Exhibit10Model> dataModel = new List<Exhibit10Model>();
            try
            {
                IExcelDataReader reader = null;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        break;

                    case "xlsx":
                        reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        break;
                }
                reader.IsFirstRowAsColumnNames = true;
                resultData = reader.AsDataSet();
                reader.Close();

                if (resultData.Tables[0].Rows.Count > 2) //判斷有無資料
                {
                    dataModel = (from q in resultData.Tables[0].AsEnumerable()
                                 select getExhibit10Models(q)).Skip(1).ToList();
                    //skip(1) 為排除Excel Title列那行(參數可調)
                }
            }
            catch
            { }
            return dataModel;
        }

        #endregion Excel 部分

        #region Private Function

        #region datarow 組成 Exhibit10Model

        /// <summary>
        /// datarow 組成 Exhibit10Model
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>Exhibit10Model</returns>
        private Exhibit10Model getExhibit10Models(DataRow item)
        {
            DateTime minDate = DateTime.MinValue;
            if (item[0] != null)
                DateTime.TryParse(item[0].ToString(), out minDate);
            return new Exhibit10Model()
            {
                Trailing = (item[0] != null) && (minDate != DateTime.MinValue) ?
                minDate.ToString("yyyy/MM/dd") : string.Empty,
                Actual_Allcorp = TypeTransfer.objToString(item[1]),
                Baseline_forecast_Allcorp = TypeTransfer.objToString(item[2]),
                Pessimistic_Forecast_Allcorp = TypeTransfer.objToString(item[3]),
                Actual_SG = TypeTransfer.objToString(item[4]),
                Baseline_forecast_SG = TypeTransfer.objToString(item[5]),
                Pessimistic_Forecast_SG = TypeTransfer.objToString(item[6])
            };
        }

        #endregion datarow 組成 Exhibit10Model

        #endregion Private Function
    }
}