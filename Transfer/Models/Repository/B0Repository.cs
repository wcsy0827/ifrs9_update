using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Repository
{
    public class B0Repository :  IB0Repository
    {
        #region 其他

        public B0Repository()
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

        #region getB06PRJID
        public List<string> getB06PRJID(string productCode)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Foward_Looking_Parm.Any())
                {
                    var query = from q in db.IFRS9_Foward_Looking_Parm.AsNoTracking()
                                            .Where(x=>x.Product_Code == productCode)
                                select q;

                    List<string> data = query.AsEnumerable().OrderBy(x => x.PRJID)
                                                            .Select(x => x.PRJID).Distinct()
                                                            .ToList();
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region getB06FLOWID
        public List<string> getB06FLOWID(string productCode,string PRJID)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Foward_Looking_Parm.Any())
                {
                    var query = from q in db.IFRS9_Foward_Looking_Parm.AsNoTracking()
                                            .Where(x => x.Product_Code == productCode)
                                select q;

                    query = query.Where(x => x.PRJID == PRJID, PRJID.IsNullOrWhiteSpace() == false);

                    List<string> data = query.AsEnumerable().OrderBy(x => x.FLOWID)
                                                            .Select(x => x.FLOWID).Distinct()
                                                            .ToList();
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region getB06CompID
        public List<string> getB06CompID(string productCode,string PRJID, string FLOWID)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Foward_Looking_Parm.Any())
                {
                    var query = from q in db.IFRS9_Foward_Looking_Parm.AsNoTracking()
                                            .Where(x => x.Product_Code == productCode)
                                select q;

                    query = query.Where(x => x.PRJID == PRJID, PRJID.IsNullOrWhiteSpace() == false)
                                 .Where(x => x.FLOWID == FLOWID, FLOWID.IsNullOrWhiteSpace() == false);

                    List<string> data = query.AsEnumerable().OrderBy(x => x.CompID)
                                                            .Select(x => x.CompID).Distinct()
                                                            .ToList();
                    return data;
                }
            }

            return new List<string>();
        }
        #endregion

        #region getB06All
        public Tuple<bool, List<B06ViewModel>> getB06All(B06ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Foward_Looking_Parm.Any())
                {
                    return new Tuple<bool, List<B06ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.IFRS9_Foward_Looking_Parm.AsNoTracking().Where(x=>x.Product_Code == dataModel.Product_Code)
                                                    .AsEnumerable()
                                        select DbToB06Model(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<B06ViewModel>>(true, new List<B06ViewModel>());
        }
        #endregion

        #region DbToB06Model
        private B06ViewModel DbToB06Model(IFRS9_Foward_Looking_Parm data)
        {
            return new B06ViewModel()
            {
                CPD_Segment_Code = data.CPD_Segment_Code,
                Delta_Q = data.Delta_Q.ToString(),
                Processing_Date = data.Processing_Date.ToString("yyyy/MM/dd"),
                Product_Code = data.Product_Code,
                PRJID = data.PRJID,
                FLOWID = data.FLOWID,
                CompID = data.CompID
            };
        }
        #endregion

        #region  getB06
        public Tuple<bool, List<B06ViewModel>> getB06(B06ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.IFRS9_Foward_Looking_Parm.Any())
                {
                    var query = from q in db.IFRS9_Foward_Looking_Parm.AsNoTracking()
                                select q;

                    if (dataModel.Processing_Date.IsNullOrWhiteSpace() == false)
                    {
                        DateTime processingDate = DateTime.Parse(dataModel.Processing_Date);
                        query = query.Where(x => x.Processing_Date >= processingDate);
                    }
                    if (!dataModel.to.IsNullOrWhiteSpace())
                    {
                        DateTime to = DateTime.Parse(dataModel.to);
                        query = query.Where(x => x.Processing_Date <= to);
                    }

                    query = query.Where(x => x.Product_Code == dataModel.Product_Code, dataModel.Product_Code.IsNullOrWhiteSpace() == false);
                    query = query.Where(x => x.PRJID == dataModel.PRJID, dataModel.PRJID.IsNullOrWhiteSpace() == false);
                    query = query.Where(x => x.FLOWID == dataModel.FLOWID, dataModel.FLOWID.IsNullOrWhiteSpace() == false);
                    query = query.Where(x => x.CompID == dataModel.CompID, dataModel.CompID.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<B06ViewModel>>(query.Any(), query.AsEnumerable()
                                                                                 .Select(x => { return DbToB06Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<B06ViewModel>>(false, new List<B06ViewModel>());
        }
        #endregion

        #region saveB06
        public MSGReturnModel saveB06(string actionType, B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    IFRS9_Foward_Looking_Parm editData = new IFRS9_Foward_Looking_Parm();

                    if (actionType == "Add")
                    {
                        dataModel.Processing_Date = DateTime.Now.ToString("yyyy/MM/dd");
                        DateTime processingDate = DateTime.Parse(dataModel.Processing_Date);

                        List<IFRS9_Foward_Looking_Parm> Datas = db.IFRS9_Foward_Looking_Parm.AsNoTracking().ToList();

                        if (Datas
                            .Where(x => x.Processing_Date == processingDate
                                     && x.Product_Code == dataModel.Product_Code
                                     && x.PRJID == "無"
                                     && x.FLOWID == "無"
                                     && x.CompID == "無")
                            .Any() == true)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = $"資料重複：資料處理日期={dataModel.Processing_Date} 的資料 已存在";
                            return result;
                        }

                        editData.CPD_Segment_Code = "";
                        editData.Delta_Q = double.Parse(dataModel.Delta_Q);
                        editData.Processing_Date = processingDate;
                        editData.Product_Code = dataModel.Product_Code;
                        editData.PRJID = "無";
                        editData.FLOWID = "無";
                        editData.CompID = "無";
                    }
                    else if (actionType == "Modify")
                    {

                    }

                    if (actionType == "Add")
                    {
                        db.IFRS9_Foward_Looking_Parm.Add(editData);
                    }

                    db.SaveChanges();

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

        #region deleteB06
        public MSGReturnModel deleteB06(B06ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    DateTime processingDate = DateTime.Parse(dataModel.Processing_Date);

                    var query = db.IFRS9_Foward_Looking_Parm
                                  .Where(x => x.Processing_Date == processingDate
                                           && x.Product_Code == dataModel.Product_Code
                                           && x.PRJID == dataModel.PRJID
                                           && x.FLOWID == dataModel.FLOWID
                                           && x.CompID == dataModel.CompID);

                    db.IFRS9_Foward_Looking_Parm.RemoveRange(query);

                    db.SaveChanges();

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
    }
}