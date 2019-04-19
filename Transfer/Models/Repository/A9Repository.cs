using Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using System.Text;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A9Repository : IA9Repository
    {     
        #region 其他
        public A9Repository()
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

        #region getA94All
        public Tuple<bool, List<A94ViewModel>> getA94All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Gov_Info_Ticker.Any())
                {
                    return new Tuple<bool, List<A94ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Gov_Info_Ticker.AsNoTracking()
                                                    .AsEnumerable()
                                        select DbToA94Model(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<A94ViewModel>>(true, new List<A94ViewModel>());
        }
        #endregion

        #region DbToA94Model
        private A94ViewModel DbToA94Model(Gov_Info_Ticker data)
        {
            return new A94ViewModel()
            {
                Country = data.Country,
                IGS_Index_Map = data.IGS_Index_Map,
                Short_term_Debt_Map = data.Short_term_Debt_Map,
                Foreign_Exchange_Map = data.Foreign_Exchange_Map,
                GDP_Yearly_Map = data.GDP_Yearly_Map
            };
        }
        #endregion

        #region getA94
        public Tuple<bool, List<A94ViewModel>> getA94(A94ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Gov_Info_Ticker.Any())
                {
                    var query = db.Gov_Info_Ticker.AsNoTracking()
                                  .Where(x => x.Country == dataModel.Country, dataModel.Country.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<A94ViewModel>>(query.Any(),
                        query.AsEnumerable().Select(x => { return DbToA94Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<A94ViewModel>>(false, new List<A94ViewModel>());
        }
        #endregion

        #region saveA94
        public MSGReturnModel saveA94(string actionType, A94ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Gov_Info_Ticker editData = new Gov_Info_Ticker();

                    if (actionType == "Add")
                    {
                        if (db.Gov_Info_Ticker.AsNoTracking()
                              .Where(x => x.Country == dataModel.Country)
                              .FirstOrDefault() != null)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = "資料重複：您輸入的 國家 已存在";
                            return result;
                        }

                        editData.Country = dataModel.Country;
                    }
                    else if (actionType == "Modify")
                    {
                        editData = db.Gov_Info_Ticker
                                     .Where(x => x.Country == dataModel.Country)
                                     .FirstOrDefault();
                    }

                    editData.IGS_Index_Map = dataModel.IGS_Index_Map;
                    editData.Short_term_Debt_Map = dataModel.Short_term_Debt_Map;
                    editData.Foreign_Exchange_Map = dataModel.Foreign_Exchange_Map;
                    editData.GDP_Yearly_Map = dataModel.GDP_Yearly_Map;

                    if (actionType == "Add")
                    {
                        db.Gov_Info_Ticker.Add(editData);
                    }

                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteA94
        public MSGReturnModel deleteA94(string country)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Gov_Info_Ticker
                                  .Where(x => x.Country == country);

                    db.Gov_Info_Ticker.RemoveRange(query);

                    db.SaveChanges(); //Save

                    result.RETURN_FLAG = true;
                }
                catch (Exception ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region GetData

        /// <summary>
        /// Get A95_1 產業別對應Ticker補登檔 
        /// </summary>
        /// <param name="bondNumber"></param>
        /// <returns></returns>
        public List<A95_1ViewModel> getA95_1(string bondNumber)
        {
            List<A95_1ViewModel> result = new List<A95_1ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = db.Assessment_Sub_Kind_Ticker.AsNoTracking()
                    .Where(x => x.Bond_Number == bondNumber, !bondNumber.IsNullOrWhiteSpace())
                    .AsEnumerable()
                    .Select(x=>new A95_1ViewModel() {
                        Bond_Number = x.Bond_Number,
                        Bloomberg_Ticker = x.Bloomberg_Ticker,
                        Security_Des = x.Security_Des,
                        Processing_Date = TypeTransfer.dateTimeNTimeSpanNToString(x.LastUpdate_Date,x.LastUpdate_Time),
                        Processing_User = x.LastUpdate_User
                    })
                    .ToList();
            }
            return result;
        }

        /// <summary>
        /// get A96 資料
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="version"></param>
        /// <returns></returns>
        public List<A96ViewModel> getA96(DateTime dt, int version)
        {
            List<A96ViewModel> result = new List<A96ViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = common.getViewModel(db.Bond_Spread_Info.AsNoTracking()
                    .Where(x => x.Report_Date == dt &&
                    x.Version == version)
                    .AsEnumerable(), Table_Type.A96).OfType<A96ViewModel>().ToList();
            }
            return result;
        }

        /// <summary>
        /// get A96 最後交易日 資料
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public List<A96TradeViewModel> getA96Trade(DateTime? dt)
        {
            List<A96TradeViewModel> result = new List<A96TradeViewModel>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result = common.getViewModel(db.Bond_Spread_Trade_Info.AsNoTracking()
                    .Where(x => x.Report_Date == dt, dt != null).AsEnumerable(),
                    Table_Type.A96_Trade).OfType<A96TradeViewModel>().ToList();
            }
            return result;
        }

        /// <summary>
        /// get A96 version
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public List<string> getA96Version(DateTime dt)
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                result.AddRange(db.Bond_Spread_Info.AsNoTracking()
                    .Where(x => x.Report_Date == dt)
                    .Select(x => x.Version.ToString()).Distinct());                   
            }
            return result;
        }

        #endregion

        #region Save Db

        #region insert A95_1 (產業別對應Ticker補登檔)
        /// <summary>
        /// insert A95_1 (產業別對應Ticker補登檔)
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel insertA95_1(List<A95_1ViewModel> datas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var _tableName = Table_Type.A95_1.tableNameGetDescription();
            if (!datas.Any())
            {
                result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription(_tableName);
                return result;
            }
            if (datas.Select(x => x.Bond_Number).Distinct().Count() != datas.Count)
            {
                result.DESCRIPTION = "Bond_Number 唯一值重複,請重新上傳!";
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(@"delete from Assessment_Sub_Kind_Ticker ; ");
                datas.ForEach(x => {
                    sb.AppendLine(
                      $@"  INSERT INTO [Assessment_Sub_Kind_Ticker]
           ([Bond_Number]
           ,[Security_Des]
           ,[Bloomberg_Ticker]
           ,[Create_User]
           ,[Create_Date]
           ,[Create_Time]
           ,[LastUpdate_User]
           ,[LastUpdate_Date]
           ,[LastUpdate_Time])
     VALUES
           ({x.Bond_Number.stringToStrSql()}
           ,{x.Security_Des.stringToStrSql()}
           ,{x.Bloomberg_Ticker.stringToStrSql()}
           ,{_UserInfo._user.stringToStrSql()}
           ,{_UserInfo._date.dateTimeToStrSql()}
           ,{_UserInfo._time.timeSpanToStrSql()}
           ,{_UserInfo._user.stringToStrSql()}
           ,{_UserInfo._date.dateTimeToStrSql()}
           ,{_UserInfo._time.timeSpanToStrSql()}) ;"
                        );
                });
                try
                {
                    db.Database.ExecuteSqlCommand(sb.ToString());
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription(_tableName);
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = Message_Type.save_Fail.GetDescription(_tableName, ex.exceptionMessage());
                }
            }
            return result;
        }
        #endregion

        #region save A95_1 (產業別對應Ticker補登檔)
        /// <summary>
        /// save A95_1 (產業別對應Ticker補登檔)
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public MSGReturnModel saveA95_1(A95_1ViewModel data, Action_Type type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            var _tableName = Table_Type.A95_1.tableNameGetDescription();
            if (data == null || data.Bond_Number.IsNullOrWhiteSpace())
            {
                result.DESCRIPTION = Message_Type.parameter_Error.GetDescription(_tableName);
                return result;
            }
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if ((type & Action_Type.Edit) == Action_Type.Edit)
                {
                    var _edit = db.Assessment_Sub_Kind_Ticker.FirstOrDefault(x => x.Bond_Number == data.Bond_Number);
                    if (_edit == null)
                    {
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription(_tableName);
                        return result;
                    }
                    _edit.Bloomberg_Ticker = data.Bloomberg_Ticker;
                    _edit.Security_Des = data.Security_Des;
                    _edit.LastUpdate_User = _UserInfo._user;
                    _edit.LastUpdate_Date = _UserInfo._date;
                    _edit.LastUpdate_Time = _UserInfo._time;
                    result.DESCRIPTION = Message_Type.update_Success.GetDescription(_tableName);
                }
                else if ((type & Action_Type.Dele) == Action_Type.Dele)
                {
                    var _dele = db.Assessment_Sub_Kind_Ticker.FirstOrDefault(x => x.Bond_Number == data.Bond_Number);
                    if (_dele == null)
                    {
                        result.DESCRIPTION = Message_Type.already_Change.GetDescription(_tableName);
                        return result;
                    }
                    db.Assessment_Sub_Kind_Ticker.Remove(_dele);
                    result.DESCRIPTION = Message_Type.delete_Success.GetDescription(_tableName);
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;                   
                }
                catch (DbUpdateException ex)
                {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }
            return result;
        }
        #endregion

        /// <summary>
        /// save A96 信用利差資料
        /// </summary>
        /// <param name="datas"></param>
        /// <returns></returns>
        public MSGReturnModel saveA96(List<A96ViewModel> datas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.not_Find_Any.GetDescription();

            if (datas.Any())
            {
                var first = datas.First();
                DateTime reportDate = DateTime.MinValue;
                int version = 0;
                if (!DateTime.TryParse(first.Report_Date, out reportDate) || !Int32.TryParse(first.Version, out version))
                {
                    result.DESCRIPTION = Message_Type.parameter_Error.GetDescription();
                    return result;
                }
                List<Bond_Spread_Info> A96s = new List<Bond_Spread_Info>();
                using (IFRS9DBEntities db = new IFRS9DBEntities())
                {
                    A96s =
                    db.Bond_Spread_Info.AsNoTracking()
                        .Where(x => x.Report_Date == reportDate && x.Version == version).ToList();
                    var sql = string.Empty;
                    StringBuilder sb = new StringBuilder();
                    datas.Where(x=>x.Reference_Nbr != null).ToList().ForEach(x =>
                {
                    var A96 = A96s.FirstOrDefault(y => y.Reference_Nbr == x.Reference_Nbr);
                    if(A96 != null)
                    {

                        if (A96.Mid_Yield != TypeTransfer.stringToDoubleN(x.Mid_Yield) ||
                            A96.Spread_Current != TypeTransfer.stringToDoubleN(x.Spread_Current) ||
                            A96.Spread_When_Trade != TypeTransfer.stringToDoubleN(x.Spread_When_Trade) ||
                            A96.Treasury_Current != TypeTransfer.stringToDoubleN(x.Treasury_Current) ||
                            A96.Treasury_When_Trade != TypeTransfer.stringToDoubleN(x.Treasury_When_Trade) ||
                            A96.All_in_Chg != TypeTransfer.stringToDoubleN(x.All_in_Chg) ||
                            A96.Chg_In_Spread != TypeTransfer.stringToDoubleN(x.Chg_In_Spread) ||
                            A96.Chg_In_Treasury != TypeTransfer.stringToDoubleN(x.Chg_In_Treasury) ||
                            TypeTransfer.stringToTrim(A96.BNCHMRK_TSY_ISSUE_ID) != TypeTransfer.stringToTrim(x.BNCHMRK_TSY_ISSUE_ID) ||
                            TypeTransfer.stringToTrim(A96.ID_CUSIP) != TypeTransfer.stringToTrim(x.ID_CUSIP)
                        )
                        {
                            sb.AppendLine($@"
update Bond_Spread_Info
   SET Processing_Date = {_UserInfo._date.dateTimeToStrSql()},
       Mid_Yield = {x.Mid_Yield.stringToDblSql()},
	   BNCHMRK_TSY_ISSUE_ID = {x.BNCHMRK_TSY_ISSUE_ID.stringToStrSql()},
	   ID_CUSIP = {x.ID_CUSIP.stringToStrSql()},
	   Spread_Current = {x.Spread_Current.stringToDblSql()},
	   Spread_When_Trade = {x.Spread_When_Trade.stringToDblSql()},
	   Treasury_Current = {x.Treasury_Current.stringToDblSql()},
	   Treasury_When_Trade = {x.Treasury_When_Trade.stringToDblSql()},
	   All_in_Chg = {x.All_in_Chg.stringToDblSql()},
	   Chg_In_Spread = {x.Chg_In_Spread.stringToDblSql()},
	   Chg_In_Treasury = {x.Chg_In_Treasury.stringToDblSql()},
       Memo = {x.Memo.stringToStrSql()},
       LastUpdate_User = {_UserInfo._user.stringToStrSql()},
       LastUpdate_Date = {_UserInfo._date.dateTimeToStrSql()},
       LastUpdate_Time = {_UserInfo._time.timeSpanToStrSql()}
	where Reference_Nbr = {x.Reference_Nbr.stringToStrSql()} ;
");
                        }
                    }
                });
                    if (sb.Length > 0)
                    {
                        try
                        {
                            db.Database.ExecuteSqlCommand(sb.ToString());
                            result.RETURN_FLAG = true;
                            result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                        }
                        catch (Exception ex)
                        {
                            result.DESCRIPTION = ex.exceptionMessage();
                        }
                    }
                    else
                    {
                        result.DESCRIPTION = Message_Type.not_Find_Update_Data.GetDescription();
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 新增,刪除,修改 A96 最後交易日
        /// </summary>
        /// <param name="data"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public MSGReturnModel saveA96Trade(A96TradeViewModel data, Action_Type type)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var reportDate = TypeTransfer.stringToDateTime(data.Report_Date);
                if (type == Action_Type.Add)
                {
                    db.Bond_Spread_Trade_Info.Add(
                        new Bond_Spread_Trade_Info()
                        {
                            Report_Date = reportDate,
                            Last_Date = TypeTransfer.stringToDateTime(data.Last_Date),
                            Create_User = _UserInfo._user,
                            Create_Date = _UserInfo._date,
                            Create_Time = _UserInfo._time,
                            LastUpdate_User = _UserInfo._user,
                            LastUpdate_Date = _UserInfo._date,
                            LastUpdate_Time = _UserInfo._time
                        });
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                else if (type == Action_Type.Dele)
                {
                    db.Bond_Spread_Trade_Info.Remove(db.Bond_Spread_Trade_Info.First(x => x.Report_Date == reportDate));
                    result.DESCRIPTION = Message_Type.delete_Success.GetDescription();
                }
                else if (type == Action_Type.Edit)
                {
                    var _trade = db.Bond_Spread_Trade_Info.FirstOrDefault(x => x.Report_Date == reportDate);
                    if (_trade != null)
                    {
                        _trade.Last_Date = TypeTransfer.stringToDateTime(data.Last_Date);
                        _trade.LastUpdate_User = _UserInfo._user;
                        _trade.LastUpdate_Date = _UserInfo._date;
                        _trade.LastUpdate_Time = _UserInfo._time;
                    }
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;              
                }
                catch (Exception ex) {
                    result.DESCRIPTION = ex.exceptionMessage();
                }
            }

            return result;
        }

        #endregion

        #region Excel 部分

        #region Excel 資料轉成 A95_1ViewModel
        /// <summary>
        /// Excel 資料轉成 A95_1ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<A95_1ViewModel>> getA95_1Excel(string pathType, Stream stream)
        {
            List<A95_1ViewModel> dataModel = new List<A95_1ViewModel>();
            string message = string.Empty;
            DataSet resultData = new DataSet();
            try
            {
                IWorkbook wb = null;
                stream.Position = 0;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        wb = new HSSFWorkbook(stream);
                        break;

                    case "xlsx":
                        wb = new XSSFWorkbook(stream);
                        break;
                }
                ISheet sheet = wb.GetSheetAt(0);
                DataTable dt = sheet.ISheetToDataTable(true);
                if (dt.Rows.Count > 0) //判斷有無資料
                {
                    dataModel = dt.AsEnumerable()
                    .Select((x, y) => new A95_1ViewModel()
                    {
                        Bond_Number = TypeTransfer.objToString(x[0]),
                        Security_Des = TypeTransfer.objToString(x[1]),
                        Bloomberg_Ticker = TypeTransfer.objToString(x[2])
                    }).ToList();
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<A95_1ViewModel>>(message, dataModel);
        }
        #endregion

        #region Excel 資料轉成 A96ViewModel
        /// <summary>
        /// Excel 資料轉成 A96ViewModel
        /// </summary>
        /// <param name="pathType"></param>
        /// <param name="stream"></param>
        /// <returns></returns>
        public Tuple<string, List<A96ViewModel>> getA96Excel(string pathType, Stream stream)
        {
            List<A96ViewModel> dataModel = new List<A96ViewModel>();
            string message = string.Empty;
            //DataSet resultData = new DataSet();
            try
            {
                IWorkbook wb = null;
                stream.Position = 0;
                switch (pathType) //判斷型別
                {
                    case "xls":
                        wb = new HSSFWorkbook(stream);
                        break;

                    case "xlsx":
                        wb = new XSSFWorkbook(stream);
                        break;
                }
                ISheet sheet = wb.GetSheetAt(0);
                DataTable dt = sheet.ISheetToDataTable(true);
                //resultData = dt.DataSet;
                if (dt.Rows.Count > 0) //判斷有無資料
                {
                    dataModel = common.getViewModel(dt, Table_Type.A96)
                                      .Cast<A96ViewModel>().ToList();
                    foreach (var item in dataModel)
                    {
                        item.Reference_Nbr = item.Reference_Nbr.PadLeft(10, '0');
                    }
                }
            }
            catch (Exception ex)
            {
                message = ex.exceptionMessage();
            }
            if (!dataModel.Any())
                message = Message_Type.not_Find_Any.GetDescription();
            return new Tuple<string, List<A96ViewModel>>(message, dataModel);
        }
        #endregion

        /// <summary>
        /// 下載 Excel
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <param name="path"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        public MSGReturnModel DownLoadExcel<T>(Excel_DownloadName type, string path, List<T> data)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail
                .GetDescription(null, Message_Type.not_Find_Any.GetDescription());
            if (type == Excel_DownloadName.A95_1)
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<A95_1ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.A95_1, new A95_1ViewModel().GetFormateTitles());
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            if (type == Excel_DownloadName.A96)
            {
                result.DESCRIPTION = FileRelated.DataTableToExcel(data.Cast<A96ViewModel>().ToList().ToDataTable(), path, Excel_DownloadName.A96);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);
            }
            return result;
        }

        #endregion

        #region Private Function


        #endregion
    }
}