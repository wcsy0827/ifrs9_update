using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Repository
{
    public class TickerRatingRepository : ITickerRatingRepository
    {
        #region 其他
        public TickerRatingRepository()
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

        #region  getGuarantorTicker
        public Tuple<bool, List<GuarantorTickerViewModel>> getGuarantorTicker(string queryType, GuarantorTickerViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = from q in db.Guarantor_Ticker.AsNoTracking()
                            select q;

                if (dataModel.Guarantor_Ticker_ID.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Guarantor_Ticker_ID.ToString() == dataModel.Guarantor_Ticker_ID);
                }

                if (dataModel.Issuer.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer == dataModel.Issuer);
                }

                if (dataModel.GUARANTOR_NAME.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.GUARANTOR_NAME == dataModel.GUARANTOR_NAME);
                }

                if (dataModel.GUARANTOR_EQY_TICKER.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.GUARANTOR_EQY_TICKER == dataModel.GUARANTOR_EQY_TICKER);
                }

                bool item1 = true;

                if (queryType != "ALL")
                {
                    item1 = query.Any();
                }

                return new Tuple<bool, 
                                 List<GuarantorTickerViewModel>>(
                                 item1, query.AsEnumerable()
                                              .Select(x => { return DbToGuarantorTickerModel(x); }).ToList());
            }
        }
        #endregion

        #region DbToGuarantorTickerModel
        private GuarantorTickerViewModel DbToGuarantorTickerModel(Guarantor_Ticker data)
        {
            return new GuarantorTickerViewModel()
            {
                Table_Name = "Guarantor_Ticker",
                Guarantor_Ticker_ID = data.Guarantor_Ticker_ID.ToString(),
                Issuer = data.Issuer,
                GUARANTOR_NAME = data.GUARANTOR_NAME,
                GUARANTOR_EQY_TICKER = data.GUARANTOR_EQY_TICKER
            };
        }
        #endregion

        #region saveGuarantorTicker
        public MSGReturnModel saveGuarantorTicker(string actionType, GuarantorTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        if (db.Guarantor_Ticker.AsNoTracking()
                                               .Where(x => x.Issuer == dataModel.Issuer)
                                               .Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = string.Format("資料重複：{0} 已存在", dataModel.Issuer);
                            return result;
                        }

                        Guarantor_Ticker addData = new Guarantor_Ticker();

                        addData.Issuer = dataModel.Issuer;
                        addData.GUARANTOR_NAME = dataModel.GUARANTOR_NAME;
                        addData.GUARANTOR_EQY_TICKER = dataModel.GUARANTOR_EQY_TICKER;
                        addData.Create_User = _UserInfo._user;
                        addData.Create_Date = _UserInfo._date;
                        addData.Create_Time = _UserInfo._time;
                        db.Guarantor_Ticker.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                        Guarantor_Ticker oldData = db.Guarantor_Ticker
                                                     .Where(x => x.Issuer == dataModel.Issuer)
                                                     .FirstOrDefault();
                        oldData.GUARANTOR_NAME = dataModel.GUARANTOR_NAME;
                        oldData.GUARANTOR_EQY_TICKER = dataModel.GUARANTOR_EQY_TICKER;
                        oldData.LastUpdate_User = _UserInfo._user;
                        oldData.LastUpdate_Date = _UserInfo._date;
                        oldData.LastUpdate_Time = _UserInfo._time;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteGuarantorTicker
        public MSGReturnModel deleteGuarantorTicker(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Guarantor_Ticker
                                  .Where(x => x.Issuer == issuer);

                    db.Guarantor_Ticker.RemoveRange(query);

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region  getIssuerTicker
        public Tuple<bool, List<IssuerTickerViewModel>> getIssuerTicker(string queryType, IssuerTickerViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = from q in db.Issuer_Ticker.AsNoTracking()
                            select q;

                if (dataModel.Issuer_Ticker_ID.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer_Ticker_ID.ToString() == dataModel.Issuer_Ticker_ID);
                }

                if (dataModel.Issuer.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer == dataModel.Issuer);
                }

                if (dataModel.ISSUER_EQUITY_TICKER.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.ISSUER_EQUITY_TICKER == dataModel.ISSUER_EQUITY_TICKER);
                }

                bool item1 = true;

                if (queryType != "ALL")
                {
                    item1 = query.Any();
                }

                return new Tuple<bool,
                                 List<IssuerTickerViewModel>>(
                                 item1, query.AsEnumerable()
                                              .Select(x => { return DbToIssuerTickerModel(x); }).ToList());
            }
        }
        #endregion

        #region DbToIssuerTickerModel
        private IssuerTickerViewModel DbToIssuerTickerModel(Issuer_Ticker data)
        {
            return new IssuerTickerViewModel()
            {
                Table_Name = "Issuer_Ticker",
                Issuer_Ticker_ID = data.Issuer_Ticker_ID.ToString(),
                Issuer = data.Issuer,
                ISSUER_EQUITY_TICKER = data.ISSUER_EQUITY_TICKER
            };
        }
        #endregion

        #region saveIssuerTicker
        public MSGReturnModel saveIssuerTicker(string actionType, IssuerTickerViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        if (db.Issuer_Ticker.AsNoTracking()
                                            .Where(x => x.Issuer == dataModel.Issuer)
                                            .Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = string.Format("資料重複：{0} 已存在", dataModel.Issuer);
                            return result;
                        }

                        Issuer_Ticker addData = new Issuer_Ticker();

                        addData.Issuer = dataModel.Issuer;
                        addData.ISSUER_EQUITY_TICKER = dataModel.ISSUER_EQUITY_TICKER;
                        addData.Create_User = _UserInfo._user;
                        addData.Create_Time = _UserInfo._time;
                        addData.Create_Date = _UserInfo._date;
                        db.Issuer_Ticker.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                        Issuer_Ticker oldData = db.Issuer_Ticker
                                                  .Where(x => x.Issuer == dataModel.Issuer)
                                                  .FirstOrDefault();

                        oldData.ISSUER_EQUITY_TICKER = dataModel.ISSUER_EQUITY_TICKER;
                        oldData.LastUpdate_User = _UserInfo._user;
                        oldData.LastUpdate_Date = _UserInfo._date;
                        oldData.LastUpdate_Time = _UserInfo._time;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteIssuerTicker
        public MSGReturnModel deleteIssuerTicker(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Issuer_Ticker
                                  .Where(x => x.Issuer == issuer);

                    db.Issuer_Ticker.RemoveRange(query);

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region  getIssuerRating
        public Tuple<bool, List<IssuerRatingViewModel>> getIssuerRating(string queryType, IssuerRatingViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = from q in db.Issuer_Rating.AsNoTracking()
                            select q;

                if (dataModel.Issuer_Rating_ID.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer_Rating_ID.ToString() == dataModel.Issuer_Rating_ID);
                }

                if (dataModel.Issuer.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer == dataModel.Issuer);
                }

                if (dataModel.S_And_P.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.S_And_P == dataModel.S_And_P);
                }

                if (dataModel.Moodys.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Moodys == dataModel.Moodys);
                }

                if (dataModel.Fitch.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch == dataModel.Fitch);
                }

                if (dataModel.Fitch_TW.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch_TW == dataModel.Fitch_TW);
                }

                if (dataModel.TRC.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.TRC == dataModel.TRC);
                }

                bool item1 = true;

                if (queryType != "ALL")
                {
                    item1 = query.Any();
                }

                return new Tuple<bool,
                                 List<IssuerRatingViewModel>>(
                                 item1, query.AsEnumerable()
                                             .Select(x => { return DbToIssuerRatingModel(x); }).ToList());
            }
        }
        #endregion

        #region DbToIssuerRatingModel
        private IssuerRatingViewModel DbToIssuerRatingModel(Issuer_Rating data)
        {
            return new IssuerRatingViewModel()
            {
                Table_Name = "Issuer_Rating",
                Issuer_Rating_ID = data.Issuer_Rating_ID.ToString(),
                Issuer = data.Issuer,
                S_And_P = data.S_And_P,
                Moodys = data.Moodys,
                Fitch = data.Fitch,
                Fitch_TW = data.Fitch_TW,
                TRC = data.TRC
            };
        }
        #endregion

        #region saveIssuerRating
        public MSGReturnModel saveIssuerRating(string actionType, IssuerRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        if (db.Issuer_Rating.AsNoTracking()
                                            .Where(x => x.Issuer == dataModel.Issuer)
                                            .Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = string.Format("資料重複：{0} 已存在", dataModel.Issuer);
                            return result;
                        }

                        Issuer_Rating addData = new Issuer_Rating();

                        addData.Issuer = dataModel.Issuer;
                        addData.S_And_P = dataModel.S_And_P;
                        addData.Moodys = dataModel.Moodys;
                        addData.Fitch = dataModel.Fitch;
                        addData.Fitch_TW = dataModel.Fitch_TW;
                        addData.TRC = dataModel.TRC;
                        addData.Create_User = _UserInfo._user;
                        addData.Create_Date = _UserInfo._date;
                        addData.Create_Time = _UserInfo._time;
                        db.Issuer_Rating.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                        Issuer_Rating oldData = db.Issuer_Rating
                                                  .Where(x => x.Issuer == dataModel.Issuer)
                                                  .FirstOrDefault();

                        oldData.S_And_P = dataModel.S_And_P;
                        oldData.Moodys = dataModel.Moodys;
                        oldData.Fitch = dataModel.Fitch;
                        oldData.Fitch_TW = dataModel.Fitch_TW;
                        oldData.TRC = dataModel.TRC;
                        oldData.LastUpdate_User = _UserInfo._user;
                        oldData.LastUpdate_Date = _UserInfo._date;
                        oldData.LastUpdate_Time = _UserInfo._time;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteIssuerRating
        public MSGReturnModel deleteIssuerRating(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Issuer_Rating
                                  .Where(x => x.Issuer == issuer);

                    db.Issuer_Rating.RemoveRange(query);

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region  getGuarantorRating
        public Tuple<bool, List<GuarantorRatingViewModel>> getGuarantorRating(string queryType, GuarantorRatingViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = from q in db.Guarantor_Rating.AsNoTracking()
                            select q;

                if (dataModel.Guarantor_Rating_ID.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Guarantor_Rating_ID.ToString() == dataModel.Guarantor_Rating_ID);
                }

                if (dataModel.Issuer.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Issuer == dataModel.Issuer);
                }

                if (dataModel.S_And_P.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.S_And_P == dataModel.S_And_P);
                }

                if (dataModel.Moodys.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Moodys == dataModel.Moodys);
                }

                if (dataModel.Fitch.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch == dataModel.Fitch);
                }

                if (dataModel.Fitch_TW.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch_TW == dataModel.Fitch_TW);
                }

                if (dataModel.TRC.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.TRC == dataModel.TRC);
                }

                bool item1 = true;

                if (queryType != "ALL")
                {
                    item1 = query.Any();
                }

                return new Tuple<bool,
                                 List<GuarantorRatingViewModel>>(
                                 item1, query.AsEnumerable()
                                             .Select(x => { return DbToGuarantorRatingModel(x); }).ToList());
            }
        }
        #endregion

        #region DbToGuarantorRatingModel
        private GuarantorRatingViewModel DbToGuarantorRatingModel(Guarantor_Rating data)
        {
            return new GuarantorRatingViewModel()
            {
                Table_Name = "Guarantor_Rating",
                Guarantor_Rating_ID = data.Guarantor_Rating_ID.ToString(),
                Issuer = data.Issuer,
                S_And_P = data.S_And_P,
                Moodys = data.Moodys,
                Fitch = data.Fitch,
                Fitch_TW = data.Fitch_TW,
                TRC = data.TRC
            };
        }
        #endregion

        #region saveGuarantorRating
        public MSGReturnModel saveGuarantorRating(string actionType, GuarantorRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        if (db.Guarantor_Rating.AsNoTracking()
                                            .Where(x => x.Issuer == dataModel.Issuer)
                                            .Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = string.Format("資料重複：{0} 已存在", dataModel.Issuer);
                            return result;
                        }

                        Guarantor_Rating addData = new Guarantor_Rating();

                        addData.Issuer = dataModel.Issuer;
                        addData.S_And_P = dataModel.S_And_P;
                        addData.Moodys = dataModel.Moodys;
                        addData.Fitch = dataModel.Fitch;
                        addData.Fitch_TW = dataModel.Fitch_TW;
                        addData.TRC = dataModel.TRC;
                        addData.Create_User = _UserInfo._user;
                        addData.Create_Date = _UserInfo._date;
                        addData.Create_Time = _UserInfo._time;
                        db.Guarantor_Rating.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                        Guarantor_Rating oldData = db.Guarantor_Rating
                                                  .Where(x => x.Issuer == dataModel.Issuer)
                                                  .FirstOrDefault();

                        oldData.S_And_P = dataModel.S_And_P;
                        oldData.Moodys = dataModel.Moodys;
                        oldData.Fitch = dataModel.Fitch;
                        oldData.Fitch_TW = dataModel.Fitch_TW;
                        oldData.TRC = dataModel.TRC;
                        oldData.LastUpdate_User = _UserInfo._user;
                        oldData.LastUpdate_Date = _UserInfo._date;
                        oldData.LastUpdate_Time = _UserInfo._time;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteGuarantorRating
        public MSGReturnModel deleteGuarantorRating(string issuer)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Guarantor_Rating
                                  .Where(x => x.Issuer == issuer);

                    db.Guarantor_Rating.RemoveRange(query);

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region  getBondRating
        public Tuple<bool, List<BondRatingViewModel>> getBondRating(string queryType, BondRatingViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                var query = from q in db.Bond_Rating.AsNoTracking()
                            select q;

                if (dataModel.Bond_Rating_ID.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Bond_Rating_ID.ToString() == dataModel.Bond_Rating_ID);
                }

                if (dataModel.Bond_Number.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Bond_Number == dataModel.Bond_Number);
                }

                if (dataModel.S_And_P.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.S_And_P == dataModel.S_And_P);
                }

                if (dataModel.Moodys.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Moodys == dataModel.Moodys);
                }

                if (dataModel.Fitch.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch == dataModel.Fitch);
                }

                if (dataModel.Fitch_TW.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.Fitch_TW == dataModel.Fitch_TW);
                }

                if (dataModel.TRC.IsNullOrWhiteSpace() == false)
                {
                    query = query.Where(x => x.TRC == dataModel.TRC);
                }

                bool item1 = true;

                if (queryType != "ALL")
                {
                    item1 = query.Any();
                }

                return new Tuple<bool,
                                 List<BondRatingViewModel>>(
                                 item1, query.AsEnumerable()
                                             .Select(x => { return DbToBondRatingModel(x); }).ToList());
            }
        }
        #endregion

        #region DbToBondRatingModel
        private BondRatingViewModel DbToBondRatingModel(Bond_Rating data)
        {
            return new BondRatingViewModel()
            {
                Table_Name = "Bond_Rating",
                Bond_Rating_ID = data.Bond_Rating_ID.ToString(),
                Bond_Number = data.Bond_Number,
                S_And_P = data.S_And_P,
                Moodys = data.Moodys,
                Fitch = data.Fitch,
                Fitch_TW = data.Fitch_TW,
                TRC = data.TRC
            };
        }
        #endregion

        #region saveBondRating
        public MSGReturnModel saveBondRating(string actionType, BondRatingViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    if (actionType == "Add")
                    {
                        if (db.Bond_Rating.AsNoTracking()
                                          .Where(x => x.Bond_Number == dataModel.Bond_Number)
                                          .Count() > 0)
                        {
                            result.RETURN_FLAG = false;
                            result.DESCRIPTION = string.Format("資料重複：{0} 已存在", dataModel.Bond_Number);
                            return result;
                        }

                        Bond_Rating addData = new Bond_Rating();

                        addData.Bond_Number = dataModel.Bond_Number;
                        addData.S_And_P = dataModel.S_And_P;
                        addData.Moodys = dataModel.Moodys;
                        addData.Fitch = dataModel.Fitch;
                        addData.Fitch_TW = dataModel.Fitch_TW;
                        addData.TRC = dataModel.TRC;
                        addData.Create_User = _UserInfo._user;
                        addData.Create_Time = _UserInfo._time;
                        addData.Create_Date = _UserInfo._date;
                        db.Bond_Rating.Add(addData);
                    }
                    else if (actionType == "Modify")
                    {
                        Bond_Rating oldData = db.Bond_Rating
                                                .Where(x => x.Bond_Number == dataModel.Bond_Number)
                                                .FirstOrDefault();

                        oldData.S_And_P = dataModel.S_And_P;
                        oldData.Moodys = dataModel.Moodys;
                        oldData.Fitch = dataModel.Fitch;
                        oldData.Fitch_TW = dataModel.Fitch_TW;
                        oldData.TRC = dataModel.TRC;
                        oldData.LastUpdate_User = _UserInfo._user;
                        oldData.LastUpdate_Date = _UserInfo._date;
                        oldData.LastUpdate_Time = _UserInfo._time;
                    }

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
                {
                    result.RETURN_FLAG = false;
                    result.DESCRIPTION = ex.Message;
                }
            }

            return result;
        }
        #endregion

        #region deleteBondRating
        public MSGReturnModel deleteBondRating(string bondNumber)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Bond_Rating
                                  .Where(x => x.Bond_Number == bondNumber);

                    db.Bond_Rating.RemoveRange(query);

                    db.SaveChanges();

                    result.RETURN_FLAG = true;
                }
                catch (DbUpdateException ex)
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