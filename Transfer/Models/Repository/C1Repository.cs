using System;
using System.Collections.Generic;
using System.Linq;
using Transfer.Models.Interface;
using Transfer.Utility;
using Transfer.ViewModels;

namespace Transfer.Models.Repository
{
    public class C1Repository : IC1Repository
    {
        #region getC13All
        public Tuple<bool, List<C13ViewModel>> getC13All()
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Econ_D_Var.Any())
                {
                    return new Tuple<bool, List<C13ViewModel>>
                                (
                                    true,
                                    (
                                        from q in db.Econ_D_Var.AsNoTracking().AsEnumerable()
                                        select DbToC13Model(q)
                                    ).ToList()
                                );
                }
            }

            return new Tuple<bool, List<C13ViewModel>>(true, new List<C13ViewModel>());
        }
        #endregion

        #region DbToC13Model
        private C13ViewModel DbToC13Model(Econ_D_Var data)
        {
            return new C13ViewModel()
            {
                Year_Quartly = data.Year_Quartly,
                Consumer_Price_Index = data.Consumer_Price_Index.ToString(),
                Consumer_Price_Index_Pre_Ind = data.Consumer_Price_Index_Pre_Ind,
                Unemployment_Rate = data.Unemployment_Rate.ToString(),
                Unemployment_Rate_Pre_Ind = data.Unemployment_Rate_Pre_Ind
            };
        }
        #endregion

        #region  getC13
        public Tuple<bool, List<C13ViewModel>> getC13(C13ViewModel dataModel)
        {
            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                if (db.Econ_D_Var.Any())
                {
                    var query = db.Econ_D_Var.AsNoTracking()
                                  .Where(x => x.Year_Quartly.Contains(dataModel.Year_Quartly), dataModel.Year_Quartly.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Consumer_Price_Index_Pre_Ind == dataModel.Consumer_Price_Index_Pre_Ind, dataModel.Consumer_Price_Index_Pre_Ind.IsNullOrWhiteSpace() == false)
                                  .Where(x => x.Unemployment_Rate_Pre_Ind == dataModel.Unemployment_Rate_Pre_Ind, dataModel.Unemployment_Rate_Pre_Ind.IsNullOrWhiteSpace() == false);

                    return new Tuple<bool, List<C13ViewModel>>(query.Any(), query.AsEnumerable()
                                                                            .Select(x => { return DbToC13Model(x); }).ToList());
                }
            }

            return new Tuple<bool, List<C13ViewModel>>(false, new List<C13ViewModel>());
        }
        #endregion

        #region saveC13
        public MSGReturnModel saveC13(string actionType, C13ViewModel dataModel)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    Econ_D_Var dataEdit = new Econ_D_Var();

                    if (actionType == "Add")
                    {
                        dataEdit.Year_Quartly = dataModel.Year_Quartly;
                    }
                    else if (actionType == "Modify")
                    {
                        dataEdit = db.Econ_D_Var
                                     .Where(x => x.Year_Quartly == dataModel.Year_Quartly)
                                     .FirstOrDefault();
                    }

                    dataEdit.Consumer_Price_Index = double.Parse(dataModel.Consumer_Price_Index);
                    dataEdit.Consumer_Price_Index_Pre_Ind = dataModel.Consumer_Price_Index_Pre_Ind;
                    dataEdit.Unemployment_Rate = double.Parse(dataModel.Unemployment_Rate);
                    dataEdit.Unemployment_Rate_Pre_Ind = dataModel.Unemployment_Rate_Pre_Ind;

                    if (actionType == "Add")
                    {
                        db.Econ_D_Var.Add(dataEdit);
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

        #region deleteC13
        public MSGReturnModel deleteC13(string yearQuartly)
        {
            MSGReturnModel result = new MSGReturnModel();

            using (IFRS9DBEntities db = new IFRS9DBEntities())
            {
                try
                {
                    var query = db.Econ_D_Var
                                  .Where(x => x.Year_Quartly == yearQuartly);
                    db.Econ_D_Var.RemoveRange(query);

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