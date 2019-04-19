using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.IO;
using System.Linq;
using Transfer.Controllers;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;
using static Transfer.Enum.Ref;

namespace Transfer.Models.Repository
{
    public class A6Repository : IA6Repository 
    {
        #region 其他

        private Common common = new Common();

        #endregion 其他

        #region Get Data

        #region get A62 Search Year

        /// <summary>
        ///
        /// </summary>
        /// <returns></returns>
        public List<string> GetA62SearchYear(string Status = "All")
        {
            List<string> result = new List<string>();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_LGD_Info.Any())
                {
                    result.Add("All(全部)");
                    result.AddRange(db.Moody_LGD_Info.AsNoTracking()
                        .Where(x=>x.Status == Status, Status != "All")
                        .GroupBy(x=> new { x.Data_Year,x.Status})                      
                        .AsEnumerable()                        
                        .Select(x => $"{x.Key.Data_Year}({A52Status(x.Key.Status)})").OrderByDescending(x => x));
                }
            }
            return result;
        }

        #endregion get A62 Search Year

        #region get Moody_LGD_Info(A62)

        /// <summary>
        /// get Moody_LGD_Info(A62)
        /// </summary>
        /// <returns></returns>
        public Tuple<bool, List<A62ViewModel>> GetA62(string dataYear)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_LGD_Info.Any())
                {
                    var query = from q in db.Moody_LGD_Info.AsNoTracking()
                                                .Where(x => x.Data_Year == dataYear, dataYear != "All")
                                select q;
                    return new Tuple<bool,
                        List<A62ViewModel>>(query.Any(), 
                        query.AsEnumerable().Select(x => { return DbToA62ViewModel(x); }).ToList());
                }
            }
            return new Tuple<bool, List<A62ViewModel>>(false, new List<A62ViewModel>());
        }

        #endregion get Moody_LGD_Info(A62)

        #region DbToA62ViewModel
        private A62ViewModel DbToA62ViewModel(Moody_LGD_Info data)
        {
            string statusName = "";
            string auditorName = "";

            StatusList statusList = new StatusList();
            List<SelectOption> selectOptionStatus = statusList.statusOption;
            for (int i = 0; i < selectOptionStatus.Count; i++)
            {
                if (data.Status.IsNullOrWhiteSpace() && selectOptionStatus[i].Value.IsNullOrWhiteSpace())
                {
                    statusName = selectOptionStatus[i].Text;
                    break;
                }
                else
                {
                    if (selectOptionStatus[i].Value == data.Status)
                    {
                        statusName = selectOptionStatus[i].Text;
                        break;
                     }
                }
            }

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var UserData = db.IFRS9_User.Where(x => x.User_Account == data.Auditor).FirstOrDefault();
                if (UserData != null)
                {
                    auditorName = UserData.User_Name;
                }
            }

            return new A62ViewModel()
            {
                Data_Year = data.Data_Year,
                Lien_Position = data.Lien_Position,
                Recovery_Rate = data.Recovery_Rate.ToString(),
                LGD = data.LGD.ToString(),
                Status = data.Status,
                Status_Name = statusName,
                Auditor_Reply = data.Auditor_Reply,
                Auditor = data.Auditor,
                Auditor_Name = auditorName,
                Audit_date = TypeTransfer.dateTimeNToString(data.Audit_date)
            };
        }
        #endregion

        #endregion Get Data

        #region DownLoadA62Excel
        public MSGReturnModel DownLoadA62Excel(string path, List<A62ViewModel> dbDatas)
        {
            MSGReturnModel result = new MSGReturnModel();
            result.RETURN_FLAG = false;
            result.DESCRIPTION = Message_Type.download_Fail.GetDescription();

            if (dbDatas.Any())
            {
                DataTable datas = getA62ModelFromDb(dbDatas).Item1;

                result.DESCRIPTION = FileRelated.DataTableToExcel(datas, path, Excel_DownloadName.A62);
                result.RETURN_FLAG = string.IsNullOrWhiteSpace(result.DESCRIPTION);

                if (result.RETURN_FLAG)
                {
                    result.DESCRIPTION = Message_Type.download_Success.GetDescription();
                }
            }

            return result;
        }
        #endregion

        #region getA62ModelFromDb
        private Tuple<DataTable> getA62ModelFromDb(List<A62ViewModel> dbDatas)
        {
            DataTable dt = new DataTable();

            try
            {
                dt.Columns.Add("年度", typeof(object));
                dt.Columns.Add("擔保順位", typeof(object));
                dt.Columns.Add("回收率", typeof(object));
                dt.Columns.Add("違約損失率", typeof(object));
                dt.Columns.Add("資料狀態", typeof(object));
                dt.Columns.Add("複核者意見", typeof(object));
                dt.Columns.Add("複核者", typeof(object));
                dt.Columns.Add("複核時間", typeof(object));

                foreach (A62ViewModel item in dbDatas)
                {
                    var nrow = dt.NewRow();

                    nrow["年度"] = item.Data_Year;
                    nrow["擔保順位"] = item.Lien_Position;
                    nrow["回收率"] = item.Recovery_Rate;
                    nrow["違約損失率"] = item.LGD;
                    nrow["資料狀態"] = item.Status_Name;
                    nrow["複核者意見"] = item.Auditor_Reply;
                    nrow["複核者"] = item.Auditor_Name;
                    nrow["複核時間"] = item.Audit_date;

                    dt.Rows.Add(nrow);
                }
            }
            catch
            {
            }

            return new Tuple<DataTable>(dt);
        }
        #endregion

        #region Save Db

        /// <summary>
        /// save A62
        /// </summary>
        /// <param name="dataModel">Exhibit7Model</param>
        /// <returns></returns>
        public MSGReturnModel saveA62(List<Exhibit7Model> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error
                    .GetDescription(Table_Type.A62.ToString());
                return result;
            }
            string dataYear = dataModel.First().Data_Year;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_LGD_Info
                    .Any(x => dataYear.Equals(x.Data_Year)))
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type.already_Save
                        .GetDescription(Table_Type.A62.ToString());
                    return result;
                }

                #region save Moody_LGD_Info(A62)

                db.Moody_LGD_Info.AddRange(
                    dataModel.Select(x => new Moody_LGD_Info()
                    {
                        Data_Year = x.Data_Year,
                        Lien_Position = x.Lien_Position,
                        Recovery_Rate = double.Parse(x.Recovery_Rate),
                        LGD = double.Parse(x.LGD)
                    }));

                #endregion save Moody_LGD_Info(A62)

                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = Message_Type.save_Success.GetDescription();
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = Message_Type
                                         .save_Fail.GetDescription(Table_Type.A62.ToString(),
                                         $"message: {ex.Message}" +
                                         $", inner message {ex.InnerException?.InnerException?.Message}");
                }
            }

            return result;
        }

        #endregion Save Db

        #region Excel 部分

        /// <summary>
        /// 把Excel 資料轉換成 Exhibit7Model
        /// </summary>
        /// <param name="pathType">string</param>
        /// <param name="stream">Stream</param>
        /// <returns>Exhibit7Model</returns>
        public List<Exhibit7Model> getExcel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<Exhibit7Model> dataModel = new List<Exhibit7Model>();
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

                if (resultData.Tables[0].Rows.Count > 5) //判斷有無資料
                {
                    string Data_Year = resultData.Tables[0].Rows[2][3].ToString();
                    dataModel = (from q in resultData.Tables[0].AsEnumerable()
                                 .Skip(3).Take(9)
                                 select getExhibit7Models(q, Data_Year)).ToList();
                    //skip(4) 為排除Excel 前4行(參數可調)
                    if (dataModel.Count() == 9)
                    {
                        //add Sr. Secured Bond &&  (Recovery_Rate & LGD = (1st Lien Bond + 2nd Lien Bond )/2)
                        dataModel.Add(
                            new Exhibit7Model()
                            {
                                Data_Year = Data_Year,
                                Lien_Position = "Sr. Secured Bond",
                                Recovery_Rate =
                                ((TypeTransfer.stringToDouble(dataModel[3].Recovery_Rate)
                                +
                                TypeTransfer.stringToDouble(dataModel[4].Recovery_Rate)) / 2).ToString(),
                                LGD =
                                ((TypeTransfer.stringToDouble(dataModel[3].LGD)
                                +
                                TypeTransfer.stringToDouble(dataModel[4].LGD)) / 2).ToString(),
                            });
                    }
                }
            }
            catch
            { }
            return dataModel;
        }

        #endregion Excel 部分

        #region Private Function

        #region datarow 組成 Exhibit7Model

        /// <summary>
        /// datarow 組成 Exhibit7Model
        /// </summary>
        /// <param name="item">DataRow</param>
        /// <returns>Exhibit7Model</returns>
        private Exhibit7Model getExhibit7Models(DataRow item, string Data_Year)
        {
            return new Exhibit7Model()
            {
                Data_Year = Data_Year,
                Lien_Position = TypeTransfer.objToString(item[0]),
                Recovery_Rate = TypeTransfer.objToString(item[3]),
                LGD = (1 - TypeTransfer.objToDouble(item[3])).ToString()
                //Recovery_Rate = string.Format("{0}%", (Recovery_Rate * 100).ToString()),
                //LGD = string.Format("{0}%", (LGD * 100).ToString())
            };
        }

        #endregion datarow 組成 Exhibit7Model

        public string A52Status(string status)
        {
            if (status == "1")
                return Audit_Type.Enable.GetDescription();
            if (status == "2")
                return Audit_Type.TempDisabled.GetDescription();
            if (status == "3")
                return Audit_Type.Disabled.GetDescription();
            return Audit_Type.None.GetDescription();
        }

        #endregion Private Function

        #region getA63Excel
        public List<A63ViewModel> getA63Excel(string pathType, Stream stream)
        {
            DataSet resultData = new DataSet();
            List<A63ViewModel> dataModel = new List<A63ViewModel>();
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

                if (resultData.Tables[0].Rows.Count > 4) //判斷有無資料
                {
                    string Data_Year = resultData.Tables[0].Rows[1][3].ToString();
                    dataModel = (from q in resultData.Tables[0].AsEnumerable()
                                                     .Skip(2).Take(4)
                                 select getA63Models(q, Data_Year)).ToList();
                }
            }
            catch
            {

            }

            return dataModel;
        }
        #endregion

        #region getA63Models
        private A63ViewModel getA63Models(DataRow item, string Data_Year)
        {
            return new A63ViewModel()
            {
                Data_Year = Data_Year,
                Lien_Position = TypeTransfer.objToString(item[0]),
                Recovery_Rate = TypeTransfer.objToString(item[3]),
                LGD = (1 - TypeTransfer.objToDouble(item[3])).ToString()
            };
        }
        #endregion

        #region saveA63
        public MSGReturnModel saveA63(List<A63ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error
                                     .GetDescription(Table_Type.A63.ToString());
                return result;
            }

            string dataYear = dataModel.First().Data_Year;
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Moody_LGD_Info.AsNoTracking().Where(x=>x.Data_Year == dataYear && x.Status == "1").Any())
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = $"{dataYear} 的資料狀態為 啟用，如要重新轉檔，請複核者先將原本的資料設為 暫不啟用。";
                    return result;
                }

                db.Moody_LGD_Info.RemoveRange(db.Moody_LGD_Info
                                                .Where(x => x.Data_Year == dataYear));

                #region save Moody_LGD_Info(A62)

                db.Moody_LGD_Info.AddRange(
                    dataModel.Select(x => new Moody_LGD_Info()
                    {
                        Data_Year = x.Data_Year,
                        Lien_Position = x.Lien_Position,
                        Recovery_Rate = double.Parse(x.Recovery_Rate),
                        LGD = double.Parse(x.LGD)
                    }));

                #endregion save Moody_LGD_Info(A62)

                try
                {
                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "年度違約損失率已更新，請通知主管複核";
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

        #region A62Audit
        public MSGReturnModel A62Audit(List<A62ViewModel> dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            if (!dataModel.Any())
            {
                result.RETURN_FLAG = false;
                result.DESCRIPTION = Message_Type.parameter_Error
                                     .GetDescription(Table_Type.A62.ToString());
                return result;
            }

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (dataModel[0].Status == "1")
                    {
                        List<Moody_LGD_Info> Datas = db.Moody_LGD_Info.Where(x => x.Status == "1").ToList();

                        for (int i = 0; i < Datas.Count; i++)
                        {
                            Datas[i].Status = "3";
                            Datas[i].Auditor = AccountController.CurrentUserInfo.Name;
                            Datas[i].Audit_date = DateTime.Now.Date;
                        }
                    }

                    for (int i = 0; i < dataModel.Count; i++)
                    {
                        var dataYear = dataModel[i].Data_Year;
                        var lienPosition = dataModel[i].Lien_Position;

                        Moody_LGD_Info data = db.Moody_LGD_Info.Where(x=>x.Data_Year == dataYear
                                                                      && x.Lien_Position == lienPosition)
                                                               .FirstOrDefault();
                        data.Status = dataModel[i].Status;
                        data.Auditor_Reply = dataModel[i].Auditor_Reply;
                        data.Auditor = AccountController.CurrentUserInfo.Name;
                        data.Audit_date = DateTime.Now.Date;
                    }

                    db.SaveChanges();
                    result.RETURN_FLAG = true;
                    result.DESCRIPTION = "複核完成";
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
    }
}